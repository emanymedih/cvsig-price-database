from __future__ import annotations

import struct
from dataclasses import dataclass
from pathlib import Path


END_OF_CHAIN = 0xFFFFFFFE
FREE_SECTOR = 0xFFFFFFFF


@dataclass
class DirectoryEntry:
    name: str
    object_type: int
    start_sector: int
    size: int


class CompoundFile:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.data = self.path.read_bytes()
        if self.data[:8] != bytes.fromhex("d0 cf 11 e0 a1 b1 1a e1"):
            raise ValueError("Not an OLE compound file")

        self.sector_size = 1 << struct.unpack_from("<H", self.data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", self.data, 32)[0]
        self.first_dir_sector = struct.unpack_from("<I", self.data, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", self.data, 56)[0]
        self.first_minifat_sector = struct.unpack_from("<I", self.data, 60)[0]
        self.num_minifat_sectors = struct.unpack_from("<I", self.data, 64)[0]
        self.num_fat_sectors = struct.unpack_from("<I", self.data, 44)[0]

        difat = list(struct.unpack_from("<109I", self.data, 76))
        fat_sector_ids = [sid for sid in difat if sid != FREE_SECTOR]
        self.fat = []
        for sid in fat_sector_ids[: self.num_fat_sectors]:
            self.fat.extend(struct.unpack("<" + "I" * (self.sector_size // 4), self.sector(sid)))

        self.directory = self._read_directory()
        self.root = next(entry for entry in self.directory if entry.object_type == 5)
        self.minifat = self._read_minifat()
        self.mini_stream = self._read_regular_stream(self.root.start_sector, self.root.size)

    def sector(self, sid: int) -> bytes:
        start = (sid + 1) * self.sector_size
        return self.data[start : start + self.sector_size]

    def sector_chain(self, start_sector: int):
        sid = start_sector
        seen = set()
        while sid not in (END_OF_CHAIN, FREE_SECTOR) and sid not in seen:
            seen.add(sid)
            yield sid
            sid = self.fat[sid]

    def _read_regular_stream(self, start_sector: int, size: int) -> bytes:
        chunks = [self.sector(sid) for sid in self.sector_chain(start_sector)]
        return b"".join(chunks)[:size]

    def _read_minifat(self) -> list[int]:
        if self.first_minifat_sector in (END_OF_CHAIN, FREE_SECTOR) or self.num_minifat_sectors == 0:
            return []
        raw = b"".join(self.sector(sid) for sid in self.sector_chain(self.first_minifat_sector))
        return list(struct.unpack("<" + "I" * (len(raw) // 4), raw))

    def _read_mini_stream(self, start_sector: int, size: int) -> bytes:
        chunks = []
        sid = start_sector
        seen = set()
        while sid not in (END_OF_CHAIN, FREE_SECTOR) and sid not in seen:
            seen.add(sid)
            start = sid * self.mini_sector_size
            chunks.append(self.mini_stream[start : start + self.mini_sector_size])
            sid = self.minifat[sid]
        return b"".join(chunks)[:size]

    def _read_directory(self) -> list[DirectoryEntry]:
        raw = b"".join(self.sector(sid) for sid in self.sector_chain(self.first_dir_sector))
        entries = []
        for pos in range(0, len(raw), 128):
            item = raw[pos : pos + 128]
            name_len = struct.unpack_from("<H", item, 64)[0]
            if name_len < 2:
                continue
            name = item[: name_len - 2].decode("utf-16le", errors="replace")
            object_type = item[66]
            start_sector = struct.unpack_from("<I", item, 116)[0]
            size = struct.unpack_from("<Q", item, 120)[0]
            entries.append(DirectoryEntry(name, object_type, start_sector, size))
        return entries

    def open_stream(self, name: str) -> bytes:
        wanted = name.lower()
        entry = next((item for item in self.directory if item.name.lower() == wanted), None)
        if entry is None:
            raise KeyError(name)
        if entry.size < self.mini_cutoff and entry.object_type == 2:
            return self._read_mini_stream(entry.start_sector, entry.size)
        return self._read_regular_stream(entry.start_sector, entry.size)


class SSTReader:
    def __init__(self, records: list[tuple[int, bytes]]):
        chunks = []
        for sid, payload in records:
            if sid == 0x00FC:
                chunks.append(payload[8:])
            elif sid == 0x003C:
                chunks.append(payload)
        self.data = b"".join(chunks)
        self.pos = 0

    def read(self, size: int) -> bytes:
        out = self.data[self.pos : self.pos + size]
        self.pos += size
        return out

    def read_string(self) -> str:
        cch = struct.unpack("<H", self.read(2))[0]
        flags = self.read(1)[0]
        is_16bit = flags & 0x01
        has_ext = flags & 0x04
        has_rich = flags & 0x08
        rich_runs = struct.unpack("<H", self.read(2))[0] if has_rich else 0
        ext_size = struct.unpack("<I", self.read(4))[0] if has_ext else 0
        byte_count = cch * (2 if is_16bit else 1)
        raw = self.read(byte_count)
        text = raw.decode("utf-16le" if is_16bit else "latin1", errors="replace")
        if rich_runs:
            self.read(rich_runs * 4)
        if ext_size:
            self.read(ext_size)
        return text


def iter_biff_records(stream: bytes, start: int = 0):
    pos = start
    while pos + 4 <= len(stream):
        sid, size = struct.unpack_from("<HH", stream, pos)
        pos += 4
        payload = stream[pos : pos + size]
        pos += size
        yield sid, payload, pos
        if sid == 0x000A:
            break


def decode_rk(raw: int) -> float:
    divide = raw & 0x01
    is_int = raw & 0x02
    if is_int:
        value = raw >> 2
        if value & 0x20000000:
            value -= 0x40000000
        result = float(value)
    else:
        packed = struct.pack("<Q", (raw & 0xFFFFFFFC) << 32)
        result = struct.unpack("<d", packed)[0]
    return result / 100 if divide else result


def formula_result(payload: bytes):
    raw = payload[6:14]
    if raw[:2] == b"\xff\xff":
        return None
    return struct.unpack("<d", raw)[0]


def read_xls(path: str | Path) -> dict[str, list[list[object]]]:
    cfb = CompoundFile(path)
    workbook_stream = cfb.open_stream("Workbook")

    boundsheets = []
    sst_records = []
    for sid, payload, _ in iter_biff_records(workbook_stream):
        if sid == 0x0085:
            offset = struct.unpack_from("<I", payload, 0)[0]
            name_len = payload[6]
            is_16bit = payload[7] & 0x01
            raw_name = payload[8 : 8 + name_len * (2 if is_16bit else 1)]
            name = raw_name.decode("utf-16le" if is_16bit else "latin1", errors="replace")
            boundsheets.append((name, offset))
        elif sid in (0x00FC, 0x003C):
            sst_records.append((sid, payload))

    sst = []
    if sst_records:
        first = sst_records[0][1]
        unique_count = struct.unpack_from("<I", first, 4)[0]
        reader = SSTReader(sst_records)
        for _ in range(unique_count):
            sst.append(reader.read_string())

    sheets: dict[str, list[list[object]]] = {}
    for sheet_name, offset in boundsheets:
        cells: dict[tuple[int, int], object] = {}
        max_row = -1
        max_col = -1
        for sid, payload, _ in iter_biff_records(workbook_stream, offset):
            row = col = value = None
            if sid == 0x00FD and len(payload) >= 10:
                row, col = struct.unpack_from("<HH", payload, 0)
                idx = struct.unpack_from("<I", payload, 6)[0]
                value = sst[idx] if idx < len(sst) else None
            elif sid == 0x0203 and len(payload) >= 14:
                row, col = struct.unpack_from("<HH", payload, 0)
                value = struct.unpack_from("<d", payload, 6)[0]
            elif sid == 0x027E and len(payload) >= 10:
                row, col = struct.unpack_from("<HH", payload, 0)
                value = decode_rk(struct.unpack_from("<I", payload, 6)[0])
            elif sid == 0x00BD and len(payload) >= 6:
                row = struct.unpack_from("<H", payload, 0)[0]
                first_col = struct.unpack_from("<H", payload, 2)[0]
                last_col = struct.unpack_from("<H", payload, len(payload) - 2)[0]
                pos = 4
                for item_col in range(first_col, last_col + 1):
                    if pos + 6 > len(payload) - 2:
                        break
                    value = decode_rk(struct.unpack_from("<I", payload, pos + 2)[0])
                    cells[(row, item_col)] = value
                    max_row = max(max_row, row)
                    max_col = max(max_col, item_col)
                    pos += 6
                continue
            elif sid == 0x0006 and len(payload) >= 14:
                row, col = struct.unpack_from("<HH", payload, 0)
                value = formula_result(payload)
            elif sid == 0x0204 and len(payload) >= 8:
                row, col = struct.unpack_from("<HH", payload, 0)
                text_len = struct.unpack_from("<H", payload, 6)[0]
                value = payload[8 : 8 + text_len].decode("latin1", errors="replace")

            if row is not None and col is not None and value is not None:
                cells[(row, col)] = value
                max_row = max(max_row, row)
                max_col = max(max_col, col)

        matrix = []
        for row_idx in range(max_row + 1):
            matrix.append([cells.get((row_idx, col_idx)) for col_idx in range(max_col + 1)])
        sheets[sheet_name] = matrix
    return sheets


if __name__ == "__main__":
    import json
    import sys

    sheets = read_xls(sys.argv[1])
    summary = {
        name: {
            "rows": len(values),
            "cols": max((len(row) for row in values), default=0),
            "preview": values[:12],
        }
        for name, values in sheets.items()
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
