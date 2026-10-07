"""Tests for src/lhos/vbaproject.py and src/lhos/cfb.py.

Run with ``python3 -I tests/test_vbaproject.py`` (exits non-zero on failure)
or with pytest.  Independent checks:

* ``oletools.olevba.decompress_stream`` cross-checks the compressor;
* ``olefile`` (strict defect level) and an in-file CFB parser validate the
  container: header, FAT/MiniFAT chains, red-black directory trees;
* ``oletools.olevba.VBA_Parser`` extracts the modules back;
* ``tests/fixtures/excel_vbaProject.bin`` is a project written by Excel
  (from the XlsxWriter 3.2.9 examples, BSD-2-Clause) used as the reference
  for record layout and for decrypting real CMG/DPB/GC values.

The LibreOffice end-to-end test lives in ``tests/test_vba_lo.py``.
"""

from __future__ import annotations

import random
import struct
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import olefile  # noqa: E402
from oletools.olevba import VBA_Parser, decompress_stream  # noqa: E402

from lhos.cfb import CfbError, cfb_name_key, write_cfb  # noqa: E402
from lhos.vbaproject import (  # noqa: E402
    VbaModule,
    VbaProjectError,
    _decrypt,
    _encrypt,
    build_vba_project,
    compress,
    decompress,
    decrypt_project_value,
    encrypt_project_value,
    module_source,
)

REFERENCE = ROOT / "tests" / "fixtures" / "excel_vbaProject.bin"
REF_ID = "{8D807122-0657-42C8-BC6F-4B5FD08031C9}"
DEFAULT_ID = "{5A1C2D3E-4F50-4A6B-8C7D-9E0F1A2B3C4D}"

CYR_CODE = """Option Explicit

' Импорт исходных данных ЛХОС
Public Const APP_TITLE As String = "Содержание ЛХОС по СИКН"

Sub WriteOk()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "OK-" & ChrW(1055)
End Sub

Function Greeting(ByVal who As String) As String
    Greeting = "Привет, " & who & "! Ёё №"
End Function
"""


def sample_modules() -> list:
    return [
        VbaModule("ThisWorkbook", "workbook", "Option Explicit\r\n"),
        VbaModule("shDash", "worksheet", "Option Explicit\rPrivate Sub Worksheet_Activate()\rEnd Sub"),
        VbaModule("modImport", "standard", CYR_CODE),
        VbaModule("clsRecord", "class", "Option Explicit\nPublic Value As Double\nPublic Name As String\n"),
    ]


def big_code(target_bytes: int = 60_000, seed: int = 7) -> str:
    """Pseudo-random but valid VBA, poorly compressible enough to exceed 4096."""
    rng = random.Random(seed)
    words = ["Объект", "Дата", "Значение", "Порог", "Лист", "Строка", "Импорт", "Сумма"]
    lines = ["Option Explicit", ""]
    n = 0
    while sum(len(x) + 2 for x in lines) < target_bytes:
        n += 1
        a, b = rng.randrange(10**6), rng.randrange(10**6)
        word = rng.choice(words)
        lines += [
            f"Function F{n}_{a:06d}(ByVal x As Double) As Double",
            f'    \' {word} {a} {b} {rng.random():.9f}',
            f"    F{n}_{a:06d} = x * {rng.random():.12f} + {b}",
            "End Function",
            "",
        ]
    return "\n".join(lines) + "\n"


def stored_code(source: str) -> str:
    """olevba's output minus our generated Attribute header, LF newlines."""
    lines = source.replace("\r\n", "\n").split("\n")
    while lines and lines[0].startswith("Attribute VB_"):
        lines.pop(0)
    return "\n".join(lines)


def normalised(code: str) -> str:
    text = code.replace("\r\n", "\n").replace("\r", "\n")
    return text if text.endswith("\n") or not text else text + "\n"


# ---------------------------------------------------------------------------
# compression
# ---------------------------------------------------------------------------

EDGE_SIZES = [0, 1, 7, 8, 9, 4095, 4096, 4097, 8192, 20000]


def _samples(size: int, rng: random.Random) -> dict:
    vba = (CYR_CODE.encode("cp1251") + big_code(25_000).encode("cp1251")) * 2
    return {
        "random": rng.randbytes(size),
        "repetitive": (b"abcabcab" * (size // 8 + 1))[:size],
        "zeros": b"\x00" * size,
        "vba": vba[:size],
    }


def test_compress_round_trip_edge_sizes() -> None:
    rng = random.Random(1234)
    for size in EDGE_SIZES:
        for label, data in _samples(size, rng).items():
            packed = compress(data)
            ours = decompress(packed)
            theirs = bytes(decompress_stream(bytearray(packed)))
            assert ours == theirs, f"{label}/{size}: our decompressor disagrees with olevba"
            if ours != data:
                # Only allowed case: incompressible final chunk shorter than
                # 4096 stored raw, which the format zero-pads ([MS-OVBA] 2.4.1.3.10).
                assert label == "random" and size % 4096, f"{label}/{size}: round trip failed"
                assert ours[:size] == data and not any(ours[size:]), f"{label}/{size}: bad padding"
                assert len(ours) % 4096 == 0
            if label in ("repetitive", "zeros", "vba") and size >= 64:
                assert len(packed) < size, f"{label}/{size}: no compression achieved"


def test_compress_matches_are_real() -> None:
    data = b"Sub Test()\r\n" * 300
    packed = compress(data)
    assert len(packed) < 40, len(packed)  # long overlapping copy tokens
    assert decompress(packed) == data


def test_compress_chunk_headers() -> None:
    data = big_code(20_000).encode("cp1251")
    packed = compress(data)
    assert packed[0] == 0x01
    pos, chunks = 1, 0
    while pos < len(packed):
        (header,) = struct.unpack_from("<H", packed, pos)
        assert (header >> 12) & 0b111 == 0b011, "chunk signature"
        assert header & 0x8000, "VBA text should produce compressed chunks"
        pos += (header & 0x0FFF) + 3
        chunks += 1
    assert pos == len(packed)
    assert chunks == -(-len(data) // 4096)


def test_decompress_reference_streams() -> None:
    ole = olefile.OleFileIO(str(REFERENCE))
    for path in ("VBA/dir", "VBA/Module1"):
        raw = ole.openstream(path).read()
        if path == "VBA/Module1":
            raw = raw[_ref_module_offsets()["Module1"]:]
        assert decompress(raw) == bytes(decompress_stream(bytearray(raw)))


# ---------------------------------------------------------------------------
# encryption (CMG / DPB / GC)
# ---------------------------------------------------------------------------

REF_VALUES = {
    "CMG": ("DBD966ECE4F0E4F0E4F0E4F0", b"\x00\x00\x00\x00"),  # ProjectProtectionState: none
    "DPB": ("5557E86E18E919E919E9", b"\x00"),  # ProjectPassword: no password
    "GC": ("CFCD72F0961011111111EE", b"\xff"),  # ProjectVisibilityState: visible
}


def test_decrypt_reference_values() -> None:
    project = _ref_project_text()
    for key, (hexstr, expected) in REF_VALUES.items():
        assert f'{key}="{hexstr}"' in project
        assert decrypt_project_value(REF_ID, hexstr) == expected, key
        # With the filler bytes Excel chose, re-encryption is byte-identical.
        seed, filler, data = _decrypt(REF_ID, hexstr)
        assert _encrypt(REF_ID, data, seed, filler) == hexstr, key


def test_encrypt_round_trip() -> None:
    for seed in range(256):
        for data in (b"", b"\x00", b"\xff", b"\x00" * 4, bytes(range(40))):
            hexstr = encrypt_project_value(DEFAULT_ID, data, seed)
            assert hexstr == hexstr.upper()
            assert len(hexstr) == 2 * (3 + (seed & 6) // 2 + 4 + len(data))
            assert decrypt_project_value(DEFAULT_ID, hexstr) == data


def test_decrypt_rejects_wrong_project_id() -> None:
    hexstr = encrypt_project_value(DEFAULT_ID, b"\x00", 0x55)
    try:
        decrypt_project_value(REF_ID, hexstr)
    except VbaProjectError:
        return
    raise AssertionError("project key mismatch not detected")


# ---------------------------------------------------------------------------
# module source rules
# ---------------------------------------------------------------------------


def test_module_source_format() -> None:
    src = module_source(VbaModule("shDash", "worksheet", "A\rB\nC\r\nD"))
    assert src.endswith(b"\r\nD\r\n")
    assert b"\r\r" not in src and src.count(b"\n") == src.count(b"\r\n")
    assert src.startswith(
        b'Attribute VB_Name = "shDash"\r\n'
        b'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r\n'
        b"Attribute VB_GlobalNameSpace = False\r\n"
        b"Attribute VB_Creatable = False\r\n"
        b"Attribute VB_PredeclaredId = True\r\n"
        b"Attribute VB_Exposed = True\r\n"
        b"Attribute VB_TemplateDerived = False\r\n"
        b"Attribute VB_Customizable = True\r\n"
    )
    wb = module_source(VbaModule("ThisWorkbook", "workbook", ""))
    assert b'VB_Base = "0{00020819-0000-0000-C000-000000000046}"' in wb
    cls = module_source(VbaModule("clsX", "class", "Public A\n"))
    assert cls == (
        b'Attribute VB_Name = "clsX"\r\n'
        b"Attribute VB_GlobalNameSpace = False\r\n"
        b"Attribute VB_Creatable = False\r\n"
        b"Attribute VB_PredeclaredId = False\r\n"
        b"Attribute VB_Exposed = False\r\n"
        b"Public A\r\n"
    )
    std = module_source(VbaModule("modA", "standard", "Sub X()\nEnd Sub\n\n"))
    assert std == b'Attribute VB_Name = "modA"\r\nSub X()\r\nEnd Sub\r\n\r\n'
    assert module_source(VbaModule("m", "standard", 'x = "Ж"'), 1251).endswith(b'"\xc6"\r\n')


def test_unencodable_character_is_reported() -> None:
    code = 'Sub A()\n    s = "a <= b"\n    t = "a ≤ b"\nEnd Sub\n'
    try:
        module_source(VbaModule("modBad", "standard", code))
    except VbaProjectError as exc:
        msg = str(exc)
        assert "modBad" in msg and "line 3" in msg and "U+2264" in msg and "ChrW(8804)" in msg, msg
    else:
        raise AssertionError("'≤' must not be encodable in cp1251")


def test_invalid_modules_rejected() -> None:
    bad_sets = [
        [],
        [VbaModule("1abc", "standard", "")],
        [VbaModule("Модуль", "standard", "")],
        [VbaModule("a" * 32, "standard", "")],
        [VbaModule("dir", "standard", "")],
        [VbaModule("modA", "form", "")],
        [VbaModule("modA", "standard", ""), VbaModule("MODA", "standard", "")],
        [VbaModule("modA", "standard", 'Attribute VB_Name = "x"\n')],
        [VbaModule("modA", "standard", "x = 1 '" + "y" * 1100)],
        [VbaModule("A", "workbook", ""), VbaModule("B", "workbook", "")],
    ]
    for mods in bad_sets:
        try:
            build_vba_project(mods)
        except VbaProjectError:
            continue
        raise AssertionError(f"accepted invalid modules {mods!r}")


# ---------------------------------------------------------------------------
# container structure
# ---------------------------------------------------------------------------


def _dir_records(data: bytes) -> list:
    """Parse a decompressed dir stream into ``(id, payload)`` tuples."""
    out, pos = [], 0
    while pos < len(data):
        rid, size = struct.unpack_from("<HI", data, pos)
        if rid == 0x0009:  # PROJECTVERSION: Reserved=4 but 6 bytes follow
            assert size == 4
            out.append((rid, data[pos + 6 : pos + 12]))
            pos += 12
            continue
        out.append((rid, data[pos + 6 : pos + 6 + size]))
        pos += 6 + size
    assert pos == len(data), "dir stream not fully consumed"
    return out


def _ref_dir() -> list:
    ole = olefile.OleFileIO(str(REFERENCE))
    return _dir_records(decompress(ole.openstream("VBA/dir").read()))


def _ref_module_offsets() -> dict:
    offsets, current = {}, None
    for rid, payload in _ref_dir():
        if rid == 0x0019:
            current = payload.decode("cp1252")
        elif rid == 0x0031:
            offsets[current] = struct.unpack("<I", payload)[0]
    return offsets


def _ref_project_text() -> str:
    ole = olefile.OleFileIO(str(REFERENCE))
    return ole.openstream("PROJECT").read().decode("cp1252")


def test_dir_stream_matches_reference_layout() -> None:
    ole = olefile.OleFileIO(build_vba_project(sample_modules()))
    ours = _dir_records(decompress(ole.openstream("VBA/dir").read()))
    ref = _ref_dir()

    # The reference also references the Office library (a second 16/3E/0D
    # triple) and has five modules; reduce both to their record "shape".
    def shape(records: list) -> list:
        ids = [rid for rid, _ in records]
        head_end = ids.index(0x000F)
        head = ids[:head_end]
        refs = head[head.index(0x0016):]
        head = head[: head.index(0x0016)] + refs[:3]  # first reference only
        first_mod = ids.index(0x0019)
        mod_end = ids.index(0x002B, first_mod) + 1
        return head + ids[head_end:first_mod] + ids[first_mod:mod_end] + ids[-1:]

    assert shape(ours) == shape(ref), (shape(ours), shape(ref))

    rec = dict()
    for rid, payload in ours:
        rec.setdefault(rid, []).append(payload)
    assert rec[0x0001] == [struct.pack("<I", 1)]
    assert rec[0x0002] == [struct.pack("<I", 0x409)] and rec[0x0014] == [struct.pack("<I", 0x409)]
    assert rec[0x0003] == [struct.pack("<H", 1251)]
    assert rec[0x0004] == [b"VBAProject"]
    assert rec[0x0016] == [b"stdole"] and rec[0x003E] == ["stdole".encode("utf-16-le")]
    libid = rec[0x000D][0]
    (n,) = struct.unpack_from("<I", libid)
    assert libid[4 : 4 + n] == (
        rb"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
        rb"C:\Windows\System32\stdole2.tlb#OLE Automation"
    )
    assert libid[4 + n :] == b"\x00" * 6
    assert rec[0x000F] == [struct.pack("<H", 4)] and rec[0x0013] == [b"\xff\xff"]
    names = [p.decode("ascii") for p in rec[0x0019]]
    assert names == ["ThisWorkbook", "shDash", "modImport", "clsRecord"]
    assert [p.decode("utf-16-le") for p in rec[0x0047]] == names
    assert [p.decode("ascii") for p in rec[0x001A]] == names
    assert [p.decode("utf-16-le") for p in rec[0x0032]] == names
    assert rec[0x0031] == [b"\x00" * 4] * 4  # MODULEOFFSET = 0
    assert rec[0x002C] == [b"\xff\xff"] * 4  # MODULECOOKIE
    types = [rid for rid, _ in ours if rid in (0x0021, 0x0022)]
    assert types == [0x0022, 0x0022, 0x0021, 0x0022]
    assert ours[-1] == (0x0010, b"")


def test_project_streams() -> None:
    data = build_vba_project(sample_modules())
    ole = olefile.OleFileIO(data)
    assert ole.openstream("VBA/_VBA_PROJECT").read() == bytes.fromhex("CC61FFFF000000")
    text = ole.openstream("PROJECT").read().decode("cp1251")
    lines = text.split("\r\n")
    assert text.endswith("\r\n") and "\n" not in text.replace("\r\n", "")
    assert lines[:5] == [
        f'ID="{DEFAULT_ID}"',
        "Document=ThisWorkbook/&H00000000",
        "Document=shDash/&H00000000",
        "Module=modImport",
        "Class=clsRecord",
    ]
    assert lines[5:8] == ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"']
    for key, expected in (("CMG", b"\x00" * 4), ("DPB", b"\x00"), ("GC", b"\xff")):
        line = next(x for x in lines if x.startswith(key + "="))
        assert decrypt_project_value(DEFAULT_ID, line.split('"')[1]) == expected
    assert lines[11:] == [
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
        "ThisWorkbook=0, 0, 0, 0, C",
        "shDash=0, 0, 0, 0, C",
        "modImport=0, 0, 0, 0, C",
        "clsRecord=0, 0, 0, 0, C",
        "",
    ]
    wm = ole.openstream("PROJECTwm").read()
    expected = b"".join(
        n.encode("ascii") + b"\x00" + n.encode("utf-16-le") + b"\x00\x00"
        for n in ("ThisWorkbook", "shDash", "modImport", "clsRecord")
    ) + b"\x00\x00"
    assert wm == expected


def test_olefile_listing_and_olevba_extraction() -> None:
    mods = sample_modules()
    data = build_vba_project(mods)
    ole = olefile.OleFileIO(data, raise_defects=olefile.DEFECT_POTENTIAL)
    listing = sorted("/".join(e) for e in ole.listdir(streams=True, storages=True))
    assert listing == sorted(
        ["PROJECT", "PROJECTwm", "VBA", "VBA/_VBA_PROJECT", "VBA/dir"]
        + ["VBA/" + m.name for m in mods]
    ), listing
    for m in mods:  # module streams hold only the compressed source
        raw = ole.openstream("VBA/" + m.name).read()
        assert decompress(raw) == module_source(m)

    parser = VBA_Parser("vbaProject.bin", data=data)
    assert parser.detect_vba_macros()
    found = {}
    for _sub, stream_path, _vba_name, code in parser.extract_macros():
        found[stream_path] = code
    parser.close()
    assert sorted(found) == sorted("VBA/" + m.name for m in mods)
    for m in mods:
        got = stored_code(found["VBA/" + m.name])
        assert got == normalised(m.code), (m.name, got)
    assert "Содержание ЛХОС по СИКН" in found["VBA/modImport"]


def test_large_module_uses_regular_sectors() -> None:
    code = big_code(60_000)
    mods = [VbaModule("ThisWorkbook", "workbook", ""), VbaModule("modBig", "standard", code)]
    data = build_vba_project(mods)
    ole = olefile.OleFileIO(data, raise_defects=olefile.DEFECT_POTENTIAL)
    size = ole.get_size("VBA/modBig")
    assert size >= 4096, size
    assert ole.get_size("VBA/ThisWorkbook") < 4096
    report = check_cfb(data)
    assert report["VBA/modBig"] == "regular" and report["VBA/ThisWorkbook"] == "mini"
    parser = VBA_Parser("big.bin", data=data)
    codes = {p: c for _s, p, _n, c in parser.extract_macros()}
    parser.close()
    assert stored_code(codes["VBA/modBig"]) == normalised(code)


def test_deterministic_output() -> None:
    assert build_vba_project(sample_modules()) == build_vba_project(sample_modules())
    other = build_vba_project(sample_modules(), project_id="{11111111-2222-3333-4444-555555555555}")
    assert other != build_vba_project(sample_modules())


# ---------------------------------------------------------------------------
# independent CFB validator
# ---------------------------------------------------------------------------

ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


def check_cfb(data: bytes, zero_metadata: bool = True) -> dict:
    """Strictly validate a CFB v3 file; return ``{path: "mini"|"regular"|"empty"}``.

    ``zero_metadata`` additionally requires CLSIDs, state bits and timestamps
    to be zero (true for our writer; Excel stamps storages with times).
    """
    assert data[:8] == bytes.fromhex("D0CF11E0A1B11AE1")
    assert data[8:24] == b"\x00" * 16
    minor, major, order, ssh, mssh = struct.unpack_from("<HHHHH", data, 24)
    assert (minor, major, order, ssh, mssh) == (0x3E, 3, 0xFFFE, 9, 6)
    assert data[34:40] == b"\x00" * 6
    (ndirsec, nfat, dir_start, txn, cutoff, mf_start, n_mf, difat_start, n_difat) = struct.unpack_from(
        "<9I", data, 40
    )
    assert ndirsec == 0 and txn == 0 and cutoff == 4096
    assert difat_start == ENDOFCHAIN and n_difat == 0
    assert len(data) % 512 == 0
    nsec = len(data) // 512 - 1
    difat = struct.unpack_from("<109I", data, 76)
    fat_secs = list(difat[:nfat])
    assert all(x == FREESECT for x in difat[nfat:])
    sector = lambda i: data[512 + i * 512 : 1024 + i * 512]  # noqa: E731
    fat = []
    for s in fat_secs:
        fat += struct.unpack("<128I", sector(s))
    assert len(fat) >= nsec
    assert all(x == FREESECT for x in fat[nsec:]), "FAT entries beyond file end"
    for s in fat_secs:
        assert fat[s] == FATSECT

    used = {s: "fat" for s in fat_secs}

    def chain(start: int, owner: str, table: list, limit: int) -> list:
        out, s = [], start
        while s != ENDOFCHAIN:
            assert 0 <= s < limit, f"{owner}: sector {s} out of range"
            out.append(s)
            assert len(out) <= limit, f"{owner}: loop"
            s = table[s]
        return out

    dir_chain = chain(dir_start, "dir", fat, nsec)
    mf_chain = chain(mf_start, "minifat", fat, nsec) if n_mf else []
    assert len(mf_chain) == n_mf
    for s in dir_chain + mf_chain:
        assert s not in used, f"sector {s} used twice"
        used[s] = "meta"

    raw_dir = b"".join(sector(s) for s in dir_chain)
    entries = []
    for i in range(len(raw_dir) // 128):
        e = raw_dir[i * 128 : (i + 1) * 128]
        (nlen,) = struct.unpack_from("<H", e, 64)
        typ, color = e[66], e[67]
        left, right, child = struct.unpack_from("<3I", e, 68)
        start, size = struct.unpack_from("<IQ", e, 116)
        if typ == 0:
            assert e[:64] == b"\x00" * 64 and nlen == 0 and (left, right, child) == (NOSTREAM,) * 3
            assert e[80:] == b"\x00" * 48
            entries.append(None)
            continue
        name = e[: nlen - 2].decode("utf-16-le")
        assert e[nlen - 2 : 64] == b"\x00" * (66 - nlen), "name padding"
        if zero_metadata:
            assert e[80:116] == b"\x00" * 36, "CLSID/state/times must be zero"
        entries.append(dict(name=name, typ=typ, color=color, left=left, right=right,
                            child=child, start=start, size=size))
    root = entries[0]
    assert root["name"] == "Root Entry" and root["typ"] == 5
    # The Root Entry is not part of any sibling tree; we write it black (as
    # most writers do), Excel's reference file has it red.
    assert root["color"] == 1 or not zero_metadata

    mini_chain = chain(root["start"], "ministream", fat, nsec) if root["size"] else []
    assert len(mini_chain) == -(-root["size"] // 512)
    for s in mini_chain:
        assert s not in used
        used[s] = "mini"
    minifat = []
    for s in mf_chain:
        minifat += struct.unpack("<128I", sector(s))
    mini_used = set()

    result, seen = {}, set()

    def walk_tree(sid: int, prefix: str, lo, hi) -> int:
        if sid == NOSTREAM:
            return 1
        assert sid not in seen, "directory entry reachable twice"
        seen.add(sid)
        e = entries[sid]
        key = cfb_name_key(e["name"])
        assert (lo is None or key > lo) and (hi is None or key < hi), "BST order"
        for c in (e["left"], e["right"]):
            if e["color"] == 0 and c != NOSTREAM:
                assert entries[c]["color"] == 1, "red-red"
        bl = walk_tree(e["left"], prefix, lo, key)
        br = walk_tree(e["right"], prefix, key, hi)
        assert bl == br, "black height"
        path = prefix + e["name"]
        if e["typ"] == 1:
            assert e["start"] == 0 and e["size"] == 0
            top = e["child"]
            if top != NOSTREAM:
                assert entries[top]["color"] == 1, "subtree root must be black"
            walk_tree(top, path + "/", None, None)
        else:
            assert e["typ"] == 2 and e["child"] == NOSTREAM
            if e["size"] == 0:
                result[path] = "empty"
            elif e["size"] < 4096:
                secs = chain(e["start"], path, minifat, len(minifat))
                assert len(secs) == -(-e["size"] // 64)
                assert not mini_used.intersection(secs)
                mini_used.update(secs)
                result[path] = "mini"
            else:
                secs = chain(e["start"], path, fat, nsec)
                assert len(secs) == -(-e["size"] // 512)
                for s in secs:
                    assert s not in used, f"sector {s} used twice"
                    used[s] = path
                result[path] = "regular"
        return bl + (e["color"] == 1)

    if root["child"] != NOSTREAM:
        assert entries[root["child"]]["color"] == 1
    walk_tree(root["child"], "", None, None)
    assert len(seen) == sum(1 for e in entries[1:] if e), "orphan directory entries"
    assert sorted(used) == list(range(nsec)), "unreferenced sectors"
    assert root["size"] == 64 * len(mini_used) or not mini_used
    assert all(minifat[i] == FREESECT for i in range(len(mini_used), len(minifat)))
    return result


def test_cfb_validator_on_reference() -> None:
    # The validator must accept Excel's own file ...
    report = check_cfb(REFERENCE.read_bytes(), zero_metadata=False)
    assert report["VBA/dir"] == "mini" and report["VBA/_VBA_PROJECT"] == "mini"
    assert len(report) == 13
    # ... and reject a corrupted red-black tree in ours (flip the colour of
    # the Root Entry's child, which must be black).
    data = bytearray(build_vba_project(sample_modules()))
    _, _, dir_start = struct.unpack_from("<III", data, 40)
    root = 512 + dir_start * 512
    (child,) = struct.unpack_from("<I", data, root + 76)
    data[root + child * 128 + 67] ^= 1
    try:
        check_cfb(bytes(data))
    except AssertionError:
        return
    raise AssertionError("validator missed a red-black violation")


def test_cfb_layouts() -> None:
    rng = random.Random(99)
    entries = {}
    for size in (0, 1, 63, 64, 65, 511, 512, 513, 4095, 4096, 4097, 70_000):
        entries[f"S/stream{size}"] = rng.randbytes(size)
    for i in range(1, 40):  # many siblings -> deep red-black tree
        entries[f"A/B/n{i:02d}{'x' * (i % 5)}"] = bytes([i]) * (i * 31)
    entries["top"] = b"hello"
    entries["EmptyStorage"] = None
    data = write_cfb(entries)
    assert data == write_cfb(dict(reversed(list(entries.items())))), "order-independent"
    report = check_cfb(data)
    ole = olefile.OleFileIO(data, raise_defects=olefile.DEFECT_POTENTIAL)
    for path, payload in entries.items():
        if payload is None:
            assert ole.get_type(path) == olefile.STGTY_STORAGE
            continue
        assert ole.openstream(path).read() == payload, path
        expected = "empty" if not payload else ("mini" if len(payload) < 4096 else "regular")
        assert report[path] == expected, (path, report[path])


def test_cfb_name_rules() -> None:
    assert cfb_name_key("Zz") < cfb_name_key("aaa")  # shorter first
    assert cfb_name_key("abc") == cfb_name_key("ABC")
    for bad in ({"a/" + "x" * 32: b""}, {"a:b": b""}, {"x": b"1", "X": b"2"}, {"s": b"1", "s/t": b"2"}):
        try:
            write_cfb(bad)
        except CfbError:
            continue
        raise AssertionError(f"accepted {bad}")


def test_built_project_passes_cfb_validator() -> None:
    report = check_cfb(build_vba_project(sample_modules()))
    assert set(report.values()) == {"mini"}
    assert report["VBA/_VBA_PROJECT"] == "mini"


# ---------------------------------------------------------------------------


def main() -> int:
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    failed = 0
    for name, func in tests:
        try:
            func()
            print(f"PASS {name}")
        except Exception:
            failed += 1
            print(f"FAIL {name}")
            traceback.print_exc()
    print(f"{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
