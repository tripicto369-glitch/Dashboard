"""Minimal, deterministic writer for Compound File Binary (CFB) containers.

Implements the subset of [MS-CFB] needed to produce ``vbaProject.bin`` files:

* version 3 files (512-byte sectors, 64-byte mini sectors, 4096-byte cutoff);
* streams smaller than the cutoff live in the mini stream owned by the Root
  Entry, larger streams use regular sectors;
* FAT, MiniFAT and the 109 header DIFAT slots (no DIFAT sector chain, which
  limits the container to roughly 6.8 MB -- far more than a VBA project needs);
* every storage's children form a red-black tree ordered with the CFB name
  ordering (shorter names first, then case-insensitive UTF-16 comparison).

Output is byte-for-byte deterministic: all timestamps and CLSIDs are zero and
directory entries are emitted in a stable order.

Usage::

    data = write_cfb({
        "PROJECT": b"...",
        "VBA/dir": b"...",
        "VBA/_VBA_PROJECT": b"...",
    })

Storages are implied by the paths.  A value of ``None`` declares an (empty)
storage explicitly.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Mapping, Optional

__all__ = ["write_cfb", "cfb_name_key", "CfbError"]

# --- [MS-CFB] constants -----------------------------------------------------

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 4096
DIR_ENTRY_SIZE = 128
ENTRIES_PER_SECTOR = SECTOR_SIZE // DIR_ENTRY_SIZE  # 4
IDS_PER_SECTOR = SECTOR_SIZE // 4  # 128 FAT/MiniFAT slots per sector
HEADER_DIFAT_SLOTS = 109

FATSECT = 0xFFFFFFFD
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF

TYPE_UNUSED = 0
TYPE_STORAGE = 1
TYPE_STREAM = 2
TYPE_ROOT = 5

RED = 0
BLACK = 1

SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
MAX_NAME_CHARS = 31  # 32 UTF-16 code units including the terminating NUL
ILLEGAL_NAME_CHARS = frozenset("/\\:!")


class CfbError(ValueError):
    """Raised for invalid input (bad names, duplicate paths, too much data)."""


# --- directory model --------------------------------------------------------


@dataclass
class _Node:
    name: str
    kind: int  # TYPE_ROOT / TYPE_STORAGE / TYPE_STREAM
    data: bytes = b""
    children: dict = field(default_factory=dict)  # upper-cased name -> _Node
    # Filled in during layout.
    sid: int = -1
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    color: int = BLACK
    start: int = ENDOFCHAIN
    size: int = 0


def _upper_unit(ch: str) -> str:
    """Upper-case one character, keeping it unchanged if the mapping is not 1:1."""
    up = ch.upper()
    return up if len(up) == 1 else ch


def cfb_name_key(name: str) -> tuple:
    """Sort key implementing the CFB directory name ordering.

    [MS-CFB] 2.6.4: a shorter name (in UTF-16 code units) is less than a longer
    one; names of equal length are compared code unit by code unit after
    upper-casing each character.
    """
    upper = "".join(_upper_unit(ch) for ch in name).encode("utf-16-le")
    units = struct.unpack("<%dH" % (len(upper) // 2), upper)
    return (len(units), units)


def _check_name(name: str, path: str) -> None:
    if not name:
        raise CfbError(f"empty path component in {path!r}")
    if len(name.encode("utf-16-le")) // 2 > MAX_NAME_CHARS:
        raise CfbError(f"name {name!r} in {path!r} exceeds {MAX_NAME_CHARS} UTF-16 units")
    bad = ILLEGAL_NAME_CHARS.intersection(name)
    if bad:
        raise CfbError(f"name {name!r} in {path!r} contains illegal characters {sorted(bad)}")


def _build_tree(entries: Mapping[str, Optional[bytes]]) -> _Node:
    """Turn ``{"A/B/stream": data}`` into a tree of _Node objects."""
    root = _Node("Root Entry", TYPE_ROOT)
    for path in sorted(entries):  # sorted -> deterministic error messages
        data = entries[path]
        parts = path.split("/")
        node = root
        for depth, part in enumerate(parts):
            _check_name(part, path)
            key = "".join(_upper_unit(c) for c in part)
            last = depth == len(parts) - 1
            existing = node.children.get(key)
            if last and data is not None:
                if existing is not None:
                    raise CfbError(f"duplicate entry {path!r} (names are case-insensitive)")
                if not isinstance(data, (bytes, bytearray, memoryview)):
                    raise CfbError(f"stream {path!r} must be bytes, got {type(data).__name__}")
                node.children[key] = _Node(part, TYPE_STREAM, bytes(data))
                break
            if existing is None:
                existing = _Node(part, TYPE_STORAGE)
                node.children[key] = existing
            elif existing.kind != TYPE_STORAGE:
                raise CfbError(f"{path!r}: {part!r} is a stream, not a storage")
            node = existing
    return root


def _assign_sids(root: _Node) -> list:
    """Pre-order numbering: Root Entry gets SID 0, children follow their parent."""
    order = []

    def visit(node: _Node) -> None:
        node.sid = len(order)
        order.append(node)
        for child in sorted(node.children.values(), key=lambda n: cfb_name_key(n.name)):
            visit(child)

    visit(root)
    return order


def _link_red_black(siblings: list) -> int:
    """Arrange sorted siblings as a balanced BST coloured as a red-black tree.

    The tree is built by recursive median split, so all NIL leaves lie at
    depth ``m`` or ``m + 1``.  Nodes at depth ``m`` (only leaves can be that
    deep) are coloured red, everything above black: every root-to-NIL path
    then contains exactly ``m`` black nodes and no red node has a red child.
    Returns the SID of the subtree root (or NOSTREAM for no siblings).
    """
    if not siblings:
        return NOSTREAM

    depths: dict = {}

    def build(lo: int, hi: int, depth: int) -> int:
        if lo >= hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        node = siblings[mid]
        depths[node.sid] = depth
        node.left = build(lo, mid, depth + 1)
        node.right = build(mid + 1, hi, depth + 1)
        return node.sid

    top = build(0, len(siblings), 0)

    # m = number of nodes on the shortest root-to-NIL path.
    by_sid = {n.sid: n for n in siblings}

    def min_nil_depth(sid: int) -> int:
        if sid == NOSTREAM:
            return 0
        n = by_sid[sid]
        return 1 + min(min_nil_depth(n.left), min_nil_depth(n.right))

    m = min_nil_depth(top)
    for node in siblings:
        node.color = RED if depths[node.sid] >= m else BLACK
    _verify_red_black(top, by_sid)
    return top


def _verify_red_black(top: int, by_sid: dict) -> None:
    """Defensive check of the red-black invariants (cheap; trees are tiny)."""
    if by_sid[top].color != BLACK:
        raise AssertionError("red-black root must be black")

    def walk(sid: int, lo, hi) -> int:
        if sid == NOSTREAM:
            return 1
        n = by_sid[sid]
        key = cfb_name_key(n.name)
        if (lo is not None and key <= lo) or (hi is not None and key >= hi):
            raise AssertionError("directory BST order violated")
        for c in (n.left, n.right):
            if n.color == RED and c != NOSTREAM and by_sid[c].color == RED:
                raise AssertionError("red node with red child")
        bl = walk(n.left, lo, key)
        br = walk(n.right, key, hi)
        if bl != br:
            raise AssertionError("unequal black height")
        return bl + (1 if n.color == BLACK else 0)

    walk(top, None, None)


# --- sector helpers ---------------------------------------------------------


def _ceil_div(a: int, b: int) -> int:
    return -(-a // b)


def _pad(data: bytes, unit: int, fill: bytes = b"\x00") -> bytes:
    rem = len(data) % unit
    return data if rem == 0 else data + fill * (unit - rem)


def _chain(fat: list, first: int, count: int) -> None:
    """Write a contiguous chain ``first .. first+count-1`` into ``fat``."""
    for i in range(count - 1):
        fat[first + i] = first + i + 1
    if count:
        fat[first + count - 1] = ENDOFCHAIN


def _dir_entry(node: Optional[_Node]) -> bytes:
    """Serialise one 128-byte directory entry ([MS-CFB] 2.6.1)."""
    if node is None:  # unused entry: zeros, sibling/child IDs = NOSTREAM
        return (
            b"\x00" * 64
            + struct.pack("<HBB", 0, TYPE_UNUSED, RED)
            + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
            + b"\x00" * 16  # CLSID
            + struct.pack("<I", 0)  # state bits
            + struct.pack("<QQ", 0, 0)  # creation / modified time
            + struct.pack("<IQ", 0, 0)  # start sector, size
        )
    name = node.name.encode("utf-16-le") + b"\x00\x00"
    is_storage = node.kind == TYPE_STORAGE
    return (
        name.ljust(64, b"\x00")
        + struct.pack("<HBB", len(name), node.kind, node.color)
        + struct.pack("<III", node.left, node.right, node.child)
        + b"\x00" * 16
        + struct.pack("<I", 0)
        + struct.pack("<QQ", 0, 0)
        # Storages carry start=0/size=0 (as Excel writes them).
        + struct.pack("<IQ", 0 if is_storage else node.start, 0 if is_storage else node.size)
    )


# --- public API -------------------------------------------------------------


def write_cfb(entries: Mapping[str, Optional[bytes]]) -> bytes:
    """Serialise ``entries`` (``"storage/.../stream" -> bytes``) as a CFB v3 file.

    A ``None`` value declares an empty storage.  Paths use ``/`` as the
    separator; names are limited to 31 UTF-16 code units and must not contain
    ``/ \\ : !``.  Sibling names are compared case-insensitively.
    """
    root = _build_tree(entries)
    nodes = _assign_sids(root)

    # Red-black trees, one per storage (the Root Entry included).
    for node in nodes:
        if node.kind in (TYPE_ROOT, TYPE_STORAGE):
            siblings = sorted(node.children.values(), key=lambda n: cfb_name_key(n.name))
            node.child = _link_red_black(siblings)
    root.color = BLACK

    streams = [n for n in nodes if n.kind == TYPE_STREAM]
    mini_streams = [n for n in streams if 0 < len(n.data) < MINI_STREAM_CUTOFF]
    big_streams = [n for n in streams if len(n.data) >= MINI_STREAM_CUTOFF]
    for n in streams:
        n.size = len(n.data)
        if not n.data:
            n.start = ENDOFCHAIN  # empty stream owns no sectors

    # Mini stream: concatenated 64-byte-aligned small streams + MiniFAT.
    minifat: list = []
    mini_parts = []
    for n in mini_streams:
        count = _ceil_div(len(n.data), MINI_SECTOR_SIZE)
        n.start = len(minifat)
        minifat.extend([0] * count)
        _chain(minifat, n.start, count)
        mini_parts.append(_pad(n.data, MINI_SECTOR_SIZE))
    mini_stream = b"".join(mini_parts)

    n_dir = _ceil_div(len(nodes), ENTRIES_PER_SECTOR)
    n_minifat = _ceil_div(len(minifat), IDS_PER_SECTOR)
    n_ministream = _ceil_div(len(mini_stream), SECTOR_SIZE)
    n_big = sum(_ceil_div(len(n.data), SECTOR_SIZE) for n in big_streams)
    payload = n_dir + n_minifat + n_ministream + n_big

    # The FAT must also describe its own sectors; iterate to a fixed point.
    n_fat = 1
    while _ceil_div(payload + n_fat, IDS_PER_SECTOR) > n_fat:
        n_fat += 1
    if n_fat > HEADER_DIFAT_SLOTS:
        raise CfbError("container too large: DIFAT sectors are not supported")
    total = n_fat + payload

    # Layout: [FAT][directory][MiniFAT][mini stream][big streams...]
    fat = [FREESECT] * (n_fat * IDS_PER_SECTOR)
    for i in range(n_fat):
        fat[i] = FATSECT
    cursor = n_fat
    dir_start = cursor
    _chain(fat, dir_start, n_dir)
    cursor += n_dir
    minifat_start = cursor if n_minifat else ENDOFCHAIN
    _chain(fat, cursor, n_minifat)
    cursor += n_minifat
    if n_ministream:
        root.start = cursor
        root.size = len(mini_stream)
        _chain(fat, cursor, n_ministream)
        cursor += n_ministream
    else:
        root.start, root.size = ENDOFCHAIN, 0
    for n in big_streams:
        count = _ceil_div(len(n.data), SECTOR_SIZE)
        n.start = cursor
        _chain(fat, cursor, count)
        cursor += count
    assert cursor == total

    # Header ([MS-CFB] 2.2).
    difat = list(range(n_fat)) + [FREESECT] * (HEADER_DIFAT_SLOTS - n_fat)
    header = (
        SIGNATURE
        + b"\x00" * 16  # header CLSID
        + struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
        + b"\x00" * 6  # reserved
        + struct.pack(
            "<IIIIIIIII",
            0,  # number of directory sectors (MUST be 0 for v3)
            n_fat,
            dir_start,
            0,  # transaction signature
            MINI_STREAM_CUTOFF,
            minifat_start,
            n_minifat,
            ENDOFCHAIN,  # first DIFAT sector
            0,  # number of DIFAT sectors
        )
        + struct.pack("<109I", *difat)
    )
    assert len(header) == SECTOR_SIZE

    # Root Entry's start/size describe the mini stream.
    dir_bytes = b"".join(_dir_entry(n) for n in nodes)
    dir_bytes += _dir_entry(None) * (n_dir * ENTRIES_PER_SECTOR - len(nodes))

    minifat_bytes = struct.pack(
        "<%dI" % (n_minifat * IDS_PER_SECTOR),
        *(minifat + [FREESECT] * (n_minifat * IDS_PER_SECTOR - len(minifat))),
    )

    body = [
        struct.pack("<%dI" % len(fat), *fat),
        dir_bytes,
        minifat_bytes,
        _pad(mini_stream, SECTOR_SIZE),
    ]
    body.extend(_pad(n.data, SECTOR_SIZE) for n in big_streams)
    out = header + b"".join(body)
    assert len(out) == SECTOR_SIZE * (1 + total)
    return out
