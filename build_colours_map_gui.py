from __future__ import annotations
import hashlib
import re
import zipfile
import ctypes
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, Button, Entry, Frame, Label, StringVar, Tk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from xml.etree import ElementTree as ET


NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
CODE_RE = re.compile(r"^[0-9A-Fa-f]{8}$")
COLOUR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
SCRIPT_DIRECTORY = Path(__file__).resolve().parent


def enable_windows_dpi_awareness() -> None:
    """Prevent Windows from bitmap-scaling the Tkinter window on HiDPI displays."""
    if __import__("sys").platform != "win32":
        return
    try:
        # Windows 10 version 1703+: per-monitor DPI awareness, including monitors
        # with different scaling factors.
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except (AttributeError, OSError):
        try:
            # Windows 7/8 and older Windows 10 fallback.
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


@dataclass(frozen=True)
class ColourRecord:
    row: int
    mode: str
    code: str
    colour: str

    @property
    def element(self) -> int:
        return int(self.code[:4], 16)

    @property
    def scheme(self) -> int:
        return int(self.code[4:], 16)

    def to_bytes(self) -> bytes:
        red, green, blue = bytes.fromhex(self.colour[1:])
        return (
            self.element.to_bytes(2, "big")
            + self.scheme.to_bytes(2, "big")
            + b"\x00"
            + bytes((red, green, blue))
        )


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.itertext()) for node in root.findall("x:si", NS)]


def worksheet_cells(workbook: Path) -> dict[str, str]:
    with zipfile.ZipFile(workbook) as archive:
        strings = shared_strings(archive)
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

    cells: dict[str, str] = {}
    for cell in root.findall(".//x:sheetData/x:row/x:c", NS):
        reference = cell.attrib["r"]
        value = cell.find("x:v", NS)
        inline = cell.find("x:is", NS)
        if value is not None and value.text is not None:
            if cell.attrib.get("t") == "s":
                cells[reference] = strings[int(value.text)]
            else:
                cells[reference] = value.text
        elif inline is not None:
            cells[reference] = "".join(inline.itertext())
        else:
            cells[reference] = ""
    return cells


def extract_records(cells: dict[str, str], day_code_column: str, day_colour_column: str, night_code_column: str, night_colour_column: str) -> list[ColourRecord]:
    records: list[ColourRecord] = []
    columns = (("Day", day_code_column, day_colour_column), ("Night", night_code_column, night_colour_column))

    for row in range(1, 1000):
        for mode, code_column, colour_column in columns:
            code = cells.get(f"{code_column}{row}", "").strip().upper()
            colour = cells.get(f"{colour_column}{row}", "").strip().upper()
            if CODE_RE.fullmatch(code) and COLOUR_RE.fullmatch(colour):
                records.append(ColourRecord(row, mode, code, colour))
    return records


def ordered_records(records: list[ColourRecord], ground_start_row: int | None) -> list[tuple[str, ColourRecord]]:
    if ground_start_row is None:
        groups = [("all", records)]
    else:
        groups = [
            ("roads", [record for record in records if record.row < ground_start_row]),
            ("ground", [record for record in records if record.row >= ground_start_row]),
        ]

    result: list[tuple[str, ColourRecord]] = []
    for group_name, group in groups:
        for mode in ("Day", "Night"):
            result.extend(
                (group_name, record)
                for record in sorted(
                    (record for record in group if record.mode == mode),
                    key=lambda record: (record.element, record.scheme),
                )
            )
    return result


def build_map(workbook: Path, output: Path, ground_start_row: int | None, columns: tuple[str, str, str, str]) -> tuple[list[tuple[str, ColourRecord]], str]:
    workbook = Path(workbook)
    output = Path(output)
    cells = worksheet_cells(workbook)
    records = extract_records(cells, *columns)
    if not records:
        raise ValueError("No valid replacement pairs found. Check the four column settings.")

    duplicate_codes = sorted({record.code for record in records if sum(other.code == record.code for other in records) > 1})
    if duplicate_codes:
        raise ValueError(f"Duplicate replacement code(s): {', '.join(duplicate_codes)}")

    ordered = ordered_records(records, ground_start_row)
    data = b"".join(record.to_bytes() for _, record in ordered)
    if len(data) % 8 != 0 or any(data[offset + 4] != 0 for offset in range(0, len(data), 8)):
        raise RuntimeError("Internal validation failed: invalid record layout.")

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)
    return ordered, hashlib.sha256(data).hexdigest().upper()


class BuilderWindow:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("RNS-E colours.map builder")
        self.root.resizable(True, True)

        self.workbook = StringVar(value=str(SCRIPT_DIRECTORY / "RNSE_Map_Colour_Legend.xlsx"))
        self.output = StringVar(value=str(SCRIPT_DIRECTORY / "colours.map"))
        self.ground_row = StringVar(value="23")
        self.day_code_column = StringVar(value="D")
        self.day_colour_column = StringVar(value="H")
        self.night_code_column = StringVar(value="K")
        self.night_colour_column = StringVar(value="O")
        self.status = StringVar(value="Choose a workbook, then build the map file.")

        self._path_row("Excel workbook", self.workbook, self.choose_workbook)
        self._path_row("Output map", self.output, self.choose_output)

        row = Frame(root, padx=10, pady=4)
        row.pack(fill=X)
        Label(row, text="Ground starts at row (blank = no split)", width=36, anchor="w").pack(side=LEFT)
        Entry(row, textvariable=self.ground_row, width=12).pack(side=LEFT)

        row = Frame(root, padx=10, pady=4)
        row.pack(fill=X)
        Label(row, text="Columns", width=14, anchor="w").pack(side=LEFT)
        for label, variable in (
            ("Day code", self.day_code_column),
            ("Day colour", self.day_colour_column),
            ("Night code", self.night_code_column),
            ("Night colour", self.night_colour_column),
        ):
            Label(row, text=label).pack(side=LEFT, padx=(8, 2))
            Entry(row, textvariable=variable, width=4).pack(side=LEFT)

        Button(root, text="Build colours.map", command=self.build, padx=18, pady=6).pack(pady=(8, 4))
        Label(root, textvariable=self.status, anchor="w", padx=10).pack(fill=X)
        self.log = ScrolledText(root, width=104, height=20, state="disabled")
        self.log.pack(fill=BOTH, expand=True, padx=10, pady=(4, 10))

    def _path_row(self, label: str, variable: StringVar, browse_command) -> None:
        row = Frame(self.root, padx=10, pady=4)
        row.pack(fill=X)
        Label(row, text=label, width=14, anchor="w").pack(side=LEFT)
        Entry(row, textvariable=variable).pack(side=LEFT, fill=X, expand=True)
        Button(row, text="Browse…", command=browse_command).pack(side=RIGHT, padx=(6, 0))

    def choose_workbook(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose RNSE map-colour workbook",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.workbook.set(selected)

    def choose_output(self) -> None:
        current = Path(self.output.get() or SCRIPT_DIRECTORY / "colours.map")
        selected = filedialog.asksaveasfilename(
            title="Save colours.map as",
            initialdir=current.parent,
            initialfile=current.name,
            defaultextension=".map",
            filetypes=[("Map file", "*.map"), ("All files", "*.*")],
        )
        if selected:
            self.output.set(selected)

    def build(self) -> None:
        try:
            workbook = Path(self.workbook.get().strip())
            output = Path(self.output.get().strip())
            if not workbook.is_file():
                raise FileNotFoundError("Choose a valid .xlsx workbook.")
            if not output.name:
                raise ValueError("Choose an output .map file.")

            row_text = self.ground_row.get().strip()
            ground_row = int(row_text) if row_text else None
            if ground_row is not None and ground_row < 1:
                raise ValueError("Ground row must be a positive number or blank.")

            columns = tuple(value.get().strip().upper() for value in (
                self.day_code_column,
                self.day_colour_column,
                self.night_code_column,
                self.night_colour_column,
            ))
            if any(not re.fullmatch(r"[A-Z]{1,3}", column) for column in columns):
                raise ValueError("Each column setting must be an Excel column letter, for example D or AA.")

            ordered, digest = build_map(workbook, output, ground_row, columns)
            lines = [
                f"{group:6} {record.mode:5} {record.code} {record.colour}  {record.to_bytes().hex(' ').upper()}"
                for group, record in ordered
            ]
            lines += ["", f"Wrote {len(ordered)} records / {len(ordered) * 8} bytes to: {output}", f"SHA-256: {digest}"]
            self._set_log("\n".join(lines))
            self.status.set("Map file created successfully.")
        except (OSError, ValueError, zipfile.BadZipFile) as error:
            self.status.set("Could not build the map file.")
            messagebox.showerror("RNS-E colours.map builder", str(error))

    def _set_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", END)
        self.log.insert("1.0", text)
        self.log.configure(state="disabled")


if __name__ == "__main__":
    enable_windows_dpi_awareness()
    app = Tk()
    BuilderWindow(app)
    app.mainloop()