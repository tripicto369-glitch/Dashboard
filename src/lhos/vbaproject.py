"""Build ``vbaProject.bin`` files from VBA source text, without Office.

The produced container follows [MS-OVBA] (VBA project structures) and
[MS-CFB] (compound file), mirroring the layout Excel writes:

    PROJECT            text stream: project properties + module list
    PROJECTwm          module name map (MBCS <-> UTF-16)
    VBA/_VBA_PROJECT   7-byte header, Version 0xFFFF, no performance cache
    VBA/dir            compressed PROJECTINFORMATION/REFERENCES/MODULES records
    VBA/<module>       compressed module source (MODULEOFFSET = 0)

Because ``_VBA_PROJECT`` carries no compiled p-code (Version 0xFFFF is what
the spec requires writers to emit), Office compiles every module from the
stored source text when the workbook is opened.

Typical use with xlsxwriter::

    bin_data = build_vba_project([
        VbaModule("ThisWorkbook", "workbook", ""),
        VbaModule("Sheet1", "worksheet", ""),
        VbaModule("modImport", "standard", "Sub Hello()\\nEnd Sub\\n"),
    ])
    Path("vbaProject.bin").write_bytes(bin_data)
    workbook.add_vba_project("vbaProject.bin")
    workbook.set_vba_name("ThisWorkbook")
    worksheet.set_vba_name("Sheet1")

Limitations (by design): only the ``stdole`` reference is declared, so code
must not rely on Office-library (``mso*``) constants; no UserForms/designers;
modules have no docstrings or help contexts.
"""

from __future__ import annotations

import re
import struct
from dataclasses import dataclass
from itertools import islice
from typing import Iterable

from .cfb import write_cfb

__all__ = [
    "VbaModule",
    "VbaProjectError",
    "build_vba_project",
    "compress",
    "decompress",
    "encrypt_project_value",
    "decrypt_project_value",
]


class VbaProjectError(ValueError):
    """Invalid module definition or source that cannot be stored."""


@dataclass
class VbaModule:
    """One VBA module.

    ``name`` is an ASCII identifier (``modImport``, ``ThisWorkbook``,
    ``shDash``); for ``workbook``/``worksheet`` modules it must equal the code
    name given to xlsxwriter's ``set_vba_name``.  ``kind`` is one of
    ``"standard" | "workbook" | "worksheet" | "class"``.  ``code`` is the
    source WITHOUT ``Attribute VB_*`` header lines, in any newline style.
    """

    name: str
    kind: str
    code: str


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.1 compression
# ---------------------------------------------------------------------------

CHUNK_SIZE = 4096  # decompressed bytes per chunk
MAX_COMPRESSED_CHUNK = 4098  # 2-byte header + 4096 payload bytes
_SIGNATURE_BYTE = 0x01
_CHUNK_SIGNATURE = 0b011
_MAX_CANDIDATES = 512  # hash-chain search depth per position


def _copy_token_help(difference: int) -> tuple:
    """[MS-OVBA] 2.4.1.3.19.1 CopyToken Help.

    ``difference`` is DecompressedCurrent - DecompressedChunkStart.  Returns
    ``(length_mask, offset_mask, bit_count, maximum_length)``.
    """
    bit_count = max((difference - 1).bit_length(), 4)  # ceil(log2(difference)), min 4
    length_mask = 0xFFFF >> bit_count
    offset_mask = ~length_mask & 0xFFFF
    maximum_length = length_mask + 3
    return length_mask, offset_mask, bit_count, maximum_length


def _compress_chunk(data: bytes, start: int, end: int) -> bytes:
    """Compress ``data[start:end]`` (at most 4096 bytes) into one CompressedChunk.

    Token sequences follow [MS-OVBA] 2.4.1.3.8/2.4.1.3.9.  Matches are found
    with hash chains over 3-byte prefixes (the minimum copy length), picking
    the longest match and, among equally long ones, the nearest -- the same
    choice as the spec's exhaustive backwards search whenever a prefix has at
    most ``_MAX_CANDIDATES`` earlier occurrences (any choice is valid output).
    Falls back to a raw chunk (2.4.1.3.10) when compression would not fit.
    """
    out = bytearray(b"\x00\x00")  # header placeholder
    heads: dict = {}  # 3-byte prefix -> positions (ascending)
    pos = start

    def remember(p: int) -> None:
        if p + 3 <= end:
            heads.setdefault(data[p : p + 3], []).append(p)

    while pos < end:
        flag_index = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= end:
                break
            best_len = 0
            best_cand = 0
            candidates = heads.get(data[pos : pos + 3]) if pos + 3 <= end else None
            if candidates:
                limit = _copy_token_help(pos - start)[3]
                limit = min(limit, end - pos)
                # Nearest first; bounded so pathological inputs stay fast.
                for cand in islice(reversed(candidates), _MAX_CANDIDATES):
                    length = 3
                    while length < limit and data[cand + length] == data[pos + length]:
                        length += 1
                    if length > best_len:
                        best_len, best_cand = length, cand
                        if length == limit:
                            break
            if best_len >= 3:
                _, _, bit_count, _ = _copy_token_help(pos - start)
                token = ((pos - best_cand - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                for p in range(pos, pos + best_len):
                    remember(p)
                pos += best_len
            else:
                out.append(data[pos])
                remember(pos)
                pos += 1
        out[flag_index] = flags
        if len(out) > MAX_COMPRESSED_CHUNK:
            break

    if len(out) <= MAX_COMPRESSED_CHUNK and pos >= end:
        header = (len(out) - 3) | (_CHUNK_SIGNATURE << 12) | 0x8000
        out[0:2] = struct.pack("<H", header)
        return bytes(out)

    # Raw chunk: header + exactly 4096 bytes (zero padded, per 2.4.1.3.10).
    raw = data[start:end].ljust(CHUNK_SIZE, b"\x00")
    header = (CHUNK_SIZE + 2 - 3) | (_CHUNK_SIGNATURE << 12)
    return struct.pack("<H", header) + raw


def compress(data: bytes) -> bytes:
    """Compress ``data`` into a [MS-OVBA] 2.4.1 CompressedContainer.

    Note: if the *final* chunk is incompressible and shorter than 4096 bytes
    it must be stored raw, which the format pads with zeros; decompression
    then yields those padding bytes too.  VBA source text always compresses,
    and :func:`build_vba_project` verifies the round trip of every stream.
    """
    data = bytes(data)
    parts = [bytes([_SIGNATURE_BYTE])]
    for start in range(0, len(data), CHUNK_SIZE):
        parts.append(_compress_chunk(data, start, min(start + CHUNK_SIZE, len(data))))
    return b"".join(parts)


def decompress(data: bytes) -> bytes:
    """Decompress a [MS-OVBA] 2.4.1 CompressedContainer."""
    data = bytes(data)
    if not data or data[0] != _SIGNATURE_BYTE:
        raise VbaProjectError("compressed container must start with signature byte 0x01")
    out = bytearray()
    pos = 1
    while pos < len(data):
        if pos + 2 > len(data):
            raise VbaProjectError("truncated chunk header at offset %d" % pos)
        (header,) = struct.unpack_from("<H", data, pos)
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0b111 != _CHUNK_SIGNATURE:
            raise VbaProjectError("bad chunk signature at offset %d" % pos)
        compressed = bool(header & 0x8000)
        chunk_end = min(pos + size, len(data))
        pos += 2
        chunk_start = len(out)
        if not compressed:
            out += data[pos : pos + CHUNK_SIZE]
            pos += CHUNK_SIZE
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not flags & (1 << bit):
                    out.append(data[pos])
                    pos += 1
                    continue
                if pos + 2 > chunk_end:
                    raise VbaProjectError("truncated copy token at offset %d" % pos)
                (token,) = struct.unpack_from("<H", data, pos)
                pos += 2
                length_mask, offset_mask, bit_count, _ = _copy_token_help(len(out) - chunk_start)
                length = (token & length_mask) + 3
                offset = ((token & offset_mask) >> (16 - bit_count)) + 1
                src = len(out) - offset
                if src < chunk_start:
                    raise VbaProjectError("copy token points before chunk start")
                for i in range(length):  # byte-wise: copies may overlap
                    out.append(out[src + i])
    return bytes(out)


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.3 data encryption (CMG / DPB / GC values of the PROJECT stream)
# ---------------------------------------------------------------------------


def _project_key(project_id: str) -> int:
    """ProjKey: sum of the bytes of the project ID string, modulo 256."""
    return sum(project_id.encode("ascii")) & 0xFF


def encrypt_project_value(project_id: str, data: bytes, seed: int) -> str:
    """Encrypt ``data`` per [MS-OVBA] 2.4.3.2; returns uppercase hex.

    ``seed`` (0..255) normally comes from a random source; a fixed seed keeps
    the output deterministic.  The ignored filler bytes are zero.
    """
    return _encrypt(project_id, data, seed, b"\x00" * ((seed & 6) // 2))


def _encrypt(project_id: str, data: bytes, seed: int, filler: bytes) -> str:
    """Encryption with explicit filler ("ignored") bytes; see 2.4.3.2."""
    if not 0 <= seed <= 0xFF:
        raise ValueError("seed must be a byte value")
    if len(filler) != (seed & 6) // 2:
        raise ValueError("seed %#04x needs %d filler bytes" % (seed, (seed & 6) // 2))
    version = 2
    proj_key = _project_key(project_id)
    version_enc = seed ^ version
    proj_key_enc = seed ^ proj_key
    out = bytearray([seed, version_enc, proj_key_enc])

    unencrypted_1 = proj_key
    encrypted_1 = proj_key_enc
    encrypted_2 = version_enc
    for byte in bytes(filler) + struct.pack("<I", len(data)) + bytes(data):
        byte_enc = byte ^ ((encrypted_2 + unencrypted_1) & 0xFF)
        out.append(byte_enc)
        encrypted_2 = encrypted_1
        encrypted_1 = byte_enc
        unencrypted_1 = byte
    return out.hex().upper()


def decrypt_project_value(project_id: str, hexstr: str) -> bytes:
    """Decrypt a CMG/DPB/GC value per [MS-OVBA] 2.4.3.3.

    Raises ``VbaProjectError`` if the version byte is not 2, the embedded
    project key does not match ``project_id``, or the length is inconsistent.
    """
    return _decrypt(project_id, hexstr)[2]


def _decrypt(project_id: str, hexstr: str) -> tuple:
    """Return ``(seed, filler, data)`` of an encrypted value."""
    enc = bytes.fromhex(hexstr)
    if len(enc) < 7:
        raise VbaProjectError("encrypted value too short")
    seed, version_enc, proj_key_enc = enc[0], enc[1], enc[2]
    if seed ^ version_enc != 2:
        raise VbaProjectError("unsupported encryption version %d" % (seed ^ version_enc))
    proj_key = seed ^ proj_key_enc
    if proj_key != _project_key(project_id):
        raise VbaProjectError("project key does not match project id %s" % project_id)

    state = {"u1": proj_key, "e1": proj_key_enc, "e2": version_enc, "pos": 3}

    def take(count: int) -> bytes:
        pos = state["pos"]
        if pos + count > len(enc):
            raise VbaProjectError("encrypted value truncated")
        result = bytearray()
        for byte_enc in enc[pos : pos + count]:
            byte = byte_enc ^ ((state["e2"] + state["u1"]) & 0xFF)
            result.append(byte)
            state["e2"], state["e1"], state["u1"] = state["e1"], byte_enc, byte
        state["pos"] = pos + count
        return bytes(result)

    filler = take((seed & 6) // 2)
    (length,) = struct.unpack("<I", take(4))
    data = take(length)
    if state["pos"] != len(enc):
        raise VbaProjectError("trailing bytes after encrypted data")
    return seed, filler, data


# ---------------------------------------------------------------------------
# Module source
# ---------------------------------------------------------------------------

KINDS = ("standard", "workbook", "worksheet", "class")
_IDENT = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")
_RESERVED_STREAMS = {"DIR", "_VBA_PROJECT", "PROJECT", "PROJECTWM"}
_MAX_LINE = 1023  # VBA editor limit for one physical line

_DOCUMENT_ATTRIBUTES = (
    'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_Creatable = False',
    'Attribute VB_PredeclaredId = True',
    'Attribute VB_Exposed = True',
    'Attribute VB_TemplateDerived = False',
    'Attribute VB_Customizable = True',
)
_BASE_WORKBOOK = '0{00020819-0000-0000-C000-000000000046}'
_BASE_WORKSHEET = '0{00020820-0000-0000-C000-000000000046}'


def _attribute_header(module: VbaModule) -> list:
    """Attribute lines exactly as Excel stores them for each module kind."""
    lines = ['Attribute VB_Name = "%s"' % module.name]
    if module.kind == "workbook":
        lines.append('Attribute VB_Base = "%s"' % _BASE_WORKBOOK)
        lines.extend(_DOCUMENT_ATTRIBUTES)
    elif module.kind == "worksheet":
        lines.append('Attribute VB_Base = "%s"' % _BASE_WORKSHEET)
        lines.extend(_DOCUMENT_ATTRIBUTES)
    elif module.kind == "class":
        lines.extend(
            (
                'Attribute VB_GlobalNameSpace = False',
                'Attribute VB_Creatable = False',
                'Attribute VB_PredeclaredId = False',
                'Attribute VB_Exposed = False',
            )
        )
    return lines


def _source_lines(code: str) -> list:
    """Split source on any newline style, dropping one trailing empty line."""
    text = code.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    return lines


def module_source(module: VbaModule, codepage: int = 1251) -> bytes:
    """Return the exact bytes stored (before compression) for ``module``.

    Attribute header + code, CRLF line endings, terminated by CRLF, encoded
    strictly in ``codepage``.  Raises ``VbaProjectError`` naming the module
    and line for characters the code page cannot represent (use ``ChrW()``).
    """
    encoding = "cp%d" % codepage
    out = []
    code_lines = _source_lines(module.code)
    for number, line in enumerate(code_lines, start=1):
        if re.match(r"^\s*Attribute\s+VB_Name\b", line, re.IGNORECASE):
            raise VbaProjectError(
                "%s line %d: code must not contain 'Attribute VB_Name' (header is generated)"
                % (module.name, number)
            )
        try:
            encoded = line.encode(encoding)
        except UnicodeEncodeError as exc:
            bad = line[exc.start : exc.end]
            raise VbaProjectError(
                "%s line %d: character %r (U+%04X) is not representable in %s; "
                "use ChrW(%d) in VBA instead"
                % (module.name, number, bad, ord(bad[0]), encoding, ord(bad[0]))
            ) from None
        if len(encoded) > _MAX_LINE:
            raise VbaProjectError(
                "%s line %d: %d bytes exceeds the VBA line limit of %d (use line continuations)"
                % (module.name, number, len(encoded), _MAX_LINE)
            )
        out.append(encoded)
    header = [line.encode("ascii") for line in _attribute_header(module)]
    return b"".join(line + b"\r\n" for line in header + out)


def _validate_modules(modules: Iterable[VbaModule]) -> list:
    modules = list(modules)
    if not modules:
        raise VbaProjectError("a VBA project needs at least one module")
    seen = set()
    for m in modules:
        if m.kind not in KINDS:
            raise VbaProjectError("module %r: kind must be one of %s" % (m.name, ", ".join(KINDS)))
        if not _IDENT.match(m.name or ""):
            raise VbaProjectError(
                "module name %r must be an ASCII identifier of at most 31 characters" % m.name
            )
        key = m.name.upper()
        if key in _RESERVED_STREAMS or key.startswith("__SRP_"):
            raise VbaProjectError("module name %r collides with a reserved stream name" % m.name)
        if key in seen:
            raise VbaProjectError("duplicate module name %r (names are case-insensitive)" % m.name)
        seen.add(key)
    if sum(1 for m in modules if m.kind == "workbook") > 1:
        raise VbaProjectError("at most one workbook module is allowed")
    return modules


# ---------------------------------------------------------------------------
# dir stream ([MS-OVBA] 2.3.4.2)
# ---------------------------------------------------------------------------

_STDOLE_LIBID = (
    r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
    r"C:\Windows\System32\stdole2.tlb#OLE Automation"
)
# PROJECTVERSION values copied from an Excel-written project; any value is valid.
_VERSION_MAJOR = 0x52671F51
_VERSION_MINOR = 0x0020


def _rec(record_id: int, payload: bytes = b"") -> bytes:
    """Generic record: Id (2 bytes) + Size (4 bytes) + payload."""
    return struct.pack("<HI", record_id, len(payload)) + payload


def _dir_stream(modules: list, codepage: int, project_name: str) -> bytes:
    enc = "cp%d" % codepage
    out = bytearray()

    # PROJECTINFORMATION
    out += _rec(0x0001, struct.pack("<I", 1))  # SYSKIND: 1 = 32-bit Windows
    out += _rec(0x0002, struct.pack("<I", 0x0409))  # LCID
    out += _rec(0x0014, struct.pack("<I", 0x0409))  # LCIDINVOKE
    out += _rec(0x0003, struct.pack("<H", codepage))  # CODEPAGE
    out += _rec(0x0004, project_name.encode(enc))  # NAME
    out += _rec(0x0005) + _rec(0x0040)  # DOCSTRING (MBCS + Unicode)
    out += _rec(0x0006) + _rec(0x003D)  # HELPFILEPATH (HelpFile1 + HelpFile2)
    out += _rec(0x0007, struct.pack("<I", 0))  # HELPCONTEXT
    out += _rec(0x0008, struct.pack("<I", 0))  # LIBFLAGS
    # VERSION: Id, Reserved (=4), VersionMajor (4 bytes), VersionMinor (2 bytes)
    out += struct.pack("<HIIH", 0x0009, 4, _VERSION_MAJOR, _VERSION_MINOR)
    out += _rec(0x000C) + _rec(0x003C)  # CONSTANTS (MBCS + Unicode)

    # PROJECTREFERENCES: stdole only
    ref_name = "stdole"
    out += _rec(0x0016, ref_name.encode(enc))  # REFERENCENAME
    out += _rec(0x003E, ref_name.encode("utf-16-le"))  # ... NameUnicode
    libid = _STDOLE_LIBID.encode("ascii")
    out += _rec(0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))

    # PROJECTMODULES
    out += _rec(0x000F, struct.pack("<H", len(modules)))  # count
    out += _rec(0x0013, struct.pack("<H", 0xFFFF))  # PROJECTCOOKIE
    for m in modules:
        name_mbcs = m.name.encode(enc)
        name_uni = m.name.encode("utf-16-le")
        out += _rec(0x0019, name_mbcs)  # MODULENAME
        out += _rec(0x0047, name_uni)  # MODULENAMEUNICODE
        out += _rec(0x001A, name_mbcs) + _rec(0x0032, name_uni)  # MODULESTREAMNAME
        out += _rec(0x001C) + _rec(0x0048)  # MODULEDOCSTRING
        out += _rec(0x0031, struct.pack("<I", 0))  # MODULEOFFSET: source at offset 0
        out += _rec(0x001E, struct.pack("<I", 0))  # MODULEHELPCONTEXT
        out += _rec(0x002C, struct.pack("<H", 0xFFFF))  # MODULECOOKIE
        out += _rec(0x0021 if m.kind == "standard" else 0x0022)  # MODULETYPE
        out += _rec(0x002B)  # module terminator
    out += _rec(0x0010)  # dir stream terminator
    return bytes(out)


# ---------------------------------------------------------------------------
# PROJECT / PROJECTwm streams ([MS-OVBA] 2.3.1, 2.3.3)
# ---------------------------------------------------------------------------

# Fixed seeds keep output deterministic (the spec only asks for "random").
_SEED_CMG, _SEED_DPB, _SEED_GC = 0xDB, 0x55, 0xCF


def _project_stream(modules: list, codepage: int, project_name: str, project_id: str) -> bytes:
    lines = ['ID="%s"' % project_id]
    for m in modules:
        if m.kind in ("workbook", "worksheet"):
            lines.append("Document=%s/&H00000000" % m.name)
        elif m.kind == "class":
            lines.append("Class=%s" % m.name)
        else:
            lines.append("Module=%s" % m.name)
    lines += [
        'Name="%s"' % project_name,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="%s"' % encrypt_project_value(project_id, b"\x00\x00\x00\x00", _SEED_CMG),
        'DPB="%s"' % encrypt_project_value(project_id, b"\x00", _SEED_DPB),
        'GC="%s"' % encrypt_project_value(project_id, b"\xff", _SEED_GC),
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    lines += ["%s=0, 0, 0, 0, C" % m.name for m in modules]
    return "".join(line + "\r\n" for line in lines).encode("cp%d" % codepage)


def _projectwm_stream(modules: list, codepage: int) -> bytes:
    out = bytearray()
    for m in modules:
        out += m.name.encode("cp%d" % codepage) + b"\x00"
        out += m.name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


# ---------------------------------------------------------------------------
# public entry point
# ---------------------------------------------------------------------------

_VBA_PROJECT_HEADER = bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])
_GUID = re.compile(r"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$")


def _compress_checked(data: bytes, what: str) -> bytes:
    packed = compress(data)
    if decompress(packed) != data:  # only possible for incompressible tails
        raise VbaProjectError("%s cannot be stored losslessly (incompressible tail)" % what)
    return packed


def build_vba_project(
    modules: list,
    *,
    codepage: int = 1251,
    project_name: str = "VBAProject",
    project_id: str = "{5A1C2D3E-4F50-4A6B-8C7D-9E0F1A2B3C4D}",
) -> bytes:
    """Return the bytes of a ``vbaProject.bin`` holding ``modules``.

    Module order is preserved in the ``dir`` and ``PROJECT`` streams.  The
    output is deterministic for identical input.
    """
    modules = _validate_modules(modules)
    if not _IDENT.match(project_name):
        raise VbaProjectError("project name %r must be an ASCII identifier" % project_name)
    if not _GUID.match(project_id):
        raise VbaProjectError("project id must look like {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}")
    try:
        "".encode("cp%d" % codepage)
    except LookupError:
        raise VbaProjectError("unsupported code page %d" % codepage) from None

    entries = {
        "PROJECT": _project_stream(modules, codepage, project_name, project_id),
        "PROJECTwm": _projectwm_stream(modules, codepage),
        "VBA/_VBA_PROJECT": _VBA_PROJECT_HEADER,
        "VBA/dir": _compress_checked(_dir_stream(modules, codepage, project_name), "dir stream"),
    }
    for m in modules:
        entries["VBA/" + m.name] = _compress_checked(
            module_source(m, codepage), "module %s" % m.name
        )
    return write_cfb(entries)
