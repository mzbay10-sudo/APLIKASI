from __future__ import annotations

import argparse
import base64
import copy
import csv
import json
import logging
import os
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree
from typing import TYPE_CHECKING, Any, Iterable
from urllib.parse import urljoin

from openpyxl import load_workbook

if TYPE_CHECKING:
    from playwright.sync_api import Page, Playwright


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
EMPTY_VALUES = {"", "0", "none", "null", "nil", "nan", "n/a", "-"}


class FlowLimitError(RuntimeError):
    """Google Flow tidak dapat melanjutkan karena kredit/kuota habis."""


class FlowBlockedError(RuntimeError):
    """Flow menolak sementara: "We noticed some unusual activity"."""


class FlowPolicyRejected(RuntimeError):
    """Semua hasil satu prompt ditolak Flow karena kebijakan konten."""


class FlowProjectNotFound(RuntimeError):
    """URL project Flow tidak ada / bukan milik akun Google profil ini."""


def project_missing(page: "Page") -> bool:
    """Deteksi halaman 'Project not found' (flow.google.com/404?reason=project)."""
    url = (page.url or "").casefold()
    if "/404" in url and "project" in url:
        return True
    try:
        text = page.locator("body").inner_text(timeout=1500)
    except Exception:
        return False
    return bool(re.search(r"project not found|proyek tidak ditemukan|project tidak ditemukan", text, re.IGNORECASE))


LIMIT_PATTERNS = [
    r"kredit (?:anda )?(?:tidak cukup|habis)", r"kehabisan kredit",
    r"kuota (?:anda )?(?:habis|terlampaui)", r"batas penggunaan",
    r"(?:telah|sudah) mencapai batas", r"tidak cukup token",
    r"not enough credits?", r"insufficient credits?", r"out of credits?",
    r"quota exceeded", r"usage limit", r"(?:you.?ve )?reached (?:your )?limit",
    r"reached (?:your |the )?(?:daily |monthly |weekly )?(?:usage |generation |image )?limit",
    r"(?:daily|monthly) (?:generation |usage |image )?limit", r"limit (?:reached|exceeded)",
    r"batas (?:harian|bulanan|penggunaan)", r"mencapai batas",
]
POLICY_PATTERNS = [
    r"prompt (?:ini )?(?:tidak diizinkan|ditolak)", r"konten (?:ini )?tidak diperbolehkan",
    r"melanggar (?:kebijakan|pedoman)", r"tidak dapat membuat (?:gambar|konten) ini",
    r"could(?:n.?t| not) generate (?:this )?(?:image|content)",
    r"prompt (?:is )?(?:not allowed|rejected)", r"violates? (?:our )?(?:policy|guidelines)",
    r"content policy", r"safety (?:policy|filter)",
]


def limit_snippet(text: str) -> str:
    """Potongan kalimat pesan limit dari Flow untuk dicatat di log."""
    normalized = " ".join(str(text or "").split())
    for pattern in LIMIT_PATTERNS:
        match = re.search(pattern, normalized, re.IGNORECASE)
        if match:
            start = max(0, match.start() - 80)
            return normalized[start:match.end() + 80]
    return normalized[:160]


def classify_flow_message(text: str) -> str | None:
    """Bedakan limit akun dari penolakan prompt berdasarkan pesan Flow."""
    normalized = " ".join(text.casefold().split())
    if any(re.search(pattern, normalized, re.IGNORECASE) for pattern in LIMIT_PATTERNS):
        return "limit"
    if any(re.search(pattern, normalized, re.IGNORECASE) for pattern in POLICY_PATTERNS):
        return "policy"
    return None


def log_file_handler(log_dir: Path) -> logging.Handler:
    """bot.log dibatasi ~5 MB (bot tetap ringan): saat mulai, log lama digeser ke
    bot.log.1 .. bot.log.3. Tanpa logging.handlers (tidak ada di EXE lama)."""
    log = log_dir / "bot.log"
    try:
        if log.exists() and log.stat().st_size > 5 * 1024 * 1024:
            for index in range(3, 0, -1):
                older = log_dir / f"bot.log.{index}"
                newer = log_dir / (f"bot.log.{index - 1}" if index > 1 else "bot.log")
                if newer.exists():
                    if older.exists():
                        older.unlink()
                    newer.rename(older)
    except OSError:
        pass  # file sedang dipakai proses lain: lanjut menulis ke file yang sama
    return logging.FileHandler(log, encoding="utf-8")


def is_empty(value: Any) -> bool:
    if value is None or value is False:
        return True
    if isinstance(value, (int, float)) and value == 0:
        return True
    return str(value).strip().casefold() in EMPTY_VALUES


def resolve_path(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else APP_DIR / path


# Kolom penanda yang ditulis bot ke Excel. Kolom ini tidak pernah dipakai
# sebagai prompt/karakter, walaupun letaknya di posisi "5", "6", dst.
MARK_STATUS = "STATUS"
MARK_FILES = "FILE GAMBAR"
MARK_NARRATION = "NARASI"
MARK_TIME = "WAKTU"
MARK_COLUMNS = [MARK_STATUS, MARK_FILES, MARK_NARRATION, MARK_TIME]
MARK_KEYS = {name.casefold() for name in MARK_COLUMNS}
DONE_PREFIXES = ("selesai", "sebagian", "ditolak")
SCENE_HEADERS = ("scene #", "scene", "scene no", "no scene", "adegan")
SOURCE_REF_HEADERS = ("source ref", "source", "ref", "sumber")


def read_rows(path: Path, sheet_name: str | None = None) -> Iterable[tuple[int, dict[str, Any]]]:
    if path.suffix.casefold() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = [str(value or "").strip() for value in (reader.fieldnames or [])]
            for number, row in enumerate(reader, start=2):
                result = dict(row)
                for index, value in enumerate(row.values(), start=1):
                    if index <= len(headers) and headers[index - 1].casefold() in MARK_KEYS:
                        continue
                    result.setdefault(str(index), value)
                yield number, result
        return
    if path.suffix.casefold() not in {".xlsx", ".xlsm"}:
        raise ValueError("Input harus berupa .csv, .xlsx, atau .xlsm")
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    iterator = sheet.iter_rows(values_only=True)
    headers = [str(value).strip() if value is not None else "" for value in next(iterator)]
    for number, values in enumerate(iterator, start=2):
        result = dict(zip(headers, values))
        # Selalu sediakan alias berdasarkan posisi kolom. Ini membuat file
        # tanpa header (baris pertama kosong) tetap terbaca sebagai kolom
        # "1", "2", "3", dan seterusnya. Kolom penanda bot (STATUS, dst.)
        # tidak diberi alias agar tidak terbaca sebagai Character.
        for index, value in enumerate(values, start=1):
            if index <= len(headers) and headers[index - 1].casefold() in MARK_KEYS:
                continue
            result.setdefault(str(index), value)
        yield number, result
    workbook.close()


def read_headers(path: Path, sheet_name: str | None = None) -> list[str]:
    """Baca baris header (baris 1) dari CSV/XLSX."""
    if path.suffix.casefold() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return [str(value or "").strip() for value in next(csv.reader(handle), [])]
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name] if sheet_name else workbook.active
        first = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        return [str(value).strip() if value is not None else "" for value in first]
    finally:
        workbook.close()


def find_header(headers: list[str], names: Iterable[str]) -> str | None:
    lookup = {header.casefold(): header for header in headers if header}
    for name in names:
        if name.casefold() in lookup:
            return lookup[name.casefold()]
    return None


def with_detected_columns(config: dict[str, Any], headers: list[str]) -> dict[str, Any]:
    """Sesuaikan kolom prompt/karakter untuk Excel yang memakai nama header.

    Format lama (RG_xxx) menaruh prompt di kolom 1 dan tetap dibaca per posisi.
    Format STORY MODE (Scene #, Source Ref, Prompt, Character 1, ...) menaruh
    prompt di kolom lain, sehingga kolom dicari berdasarkan nama header.
    """
    result = copy.deepcopy(config)
    prompt_header = find_header(headers, ["Prompt"])
    if prompt_header is None or headers.index(prompt_header) == 0:
        return result
    result.setdefault("prompt", {})["column"] = prompt_header
    fields = {}
    for name, settings in result.get("fields", {}).items():
        header = find_header(headers, [name])
        if header is not None:
            fields[name] = {**settings, "column": header}
    result["fields"] = fields
    result["_header_mode"] = True
    return result


def is_marked_done(row: dict[str, Any]) -> bool:
    for key, value in row.items():
        if str(key).casefold() == MARK_STATUS.casefold():
            return str(value or "").strip().casefold().startswith(DONE_PREFIXES)
    return False


def prompt_rows(path: Path, config: dict[str, Any]) -> list[tuple[int, dict[str, Any]]]:
    """Ambil baris berprompt yang belum bertanda SELESAI/SEBAGIAN di Excel."""
    config = with_detected_columns(config, read_headers(path, config.get("sheet_name")))
    start_row = int(config.get("start_row", 2))
    prompt_column = str(config.get("prompt", {}).get("column", "1"))
    return [
        (number, row)
        for number, row in read_rows(path, config.get("sheet_name"))
        if number >= start_row and not is_empty(row.get(prompt_column)) and not is_marked_done(row)
    ]


def config_for_input(config: dict[str, Any], input_path: Path) -> dict[str, Any]:
    """Buat folder dan nama hasil berdasarkan nama file input."""
    headers = read_headers(input_path, config.get("sheet_name"))
    result = with_detected_columns(config, headers)
    safe_name = re.sub(r'[<>:"/\\|?*]+', "_", input_path.stem).strip(" .") or "hasil"
    root = resolve_path(str(result.get("download", {}).get("root_folder", "downloads")))
    result["download"]["folder"] = str(root / safe_name)
    result["download"]["base_name"] = safe_name
    result["_scene_column"] = find_header(headers, SCENE_HEADERS)
    return result


def scene_label(row_number: int, row: dict[str, Any], config: dict[str, Any]) -> str:
    """Label scene untuk nama file: S001, S002, ... (dari kolom Scene # bila ada)."""
    column = config.get("_scene_column")
    value = row.get(column) if column else None
    if not is_empty(value):
        text = str(value).strip()
        if re.fullmatch(r"\d+(?:\.0+)?", text):
            return f"S{int(float(text)):03d}"
        safe = re.sub(r'[<>:"/\\|?*\s]+', "_", text).strip("_.")
        if safe:
            return f"S{safe}"
    # Tanpa kolom Scene: pakai nomor baris data (baris Excel 2 = S001).
    return f"S{max(1, row_number - 1):03d}"


# ---------------------------------------------------------------------------
# Narasi dari .docx
# ---------------------------------------------------------------------------

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_paragraphs(path: Path) -> list[str]:
    """Ambil paragraf tidak kosong dari badan .docx (tanpa python-docx)."""
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("word/document.xml"))
    body = root.find(f"{W_NS}body")
    paragraphs = []
    for paragraph in (body if body is not None else []):
        if paragraph.tag != f"{W_NS}p":
            continue
        text = "".join(node.text or "" for node in paragraph.iter(f"{W_NS}t")).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def find_narration_docx(input_path: Path) -> Path | None:
    """Cari .docx narasi di folder yang sama dengan Excel.

    Contoh: Episode_041.xlsx cocok dengan BTS_MODE_FINAL_Episode_041_TTS_SAFE.docx.
    """
    folder = input_path.parent
    docs = [path for path in folder.glob("*.docx") if not path.name.startswith("~$")]
    stem = input_path.stem.casefold()
    matches = [path for path in docs if stem in path.stem.casefold()]
    if not matches:
        episode = re.search(r"(?:episode|eps?|ep)[ _-]*0*(\d+)", stem)
        if episode:
            number = episode.group(1)
            matches = [
                path for path in docs
                if re.search(rf"(?:episode|eps?|ep)[ _-]*0*{number}(?!\d)", path.stem.casefold())
            ]
    if not matches:
        return None
    matches.sort(key=lambda path: ("tts" not in path.stem.casefold(), -path.stat().st_mtime))
    return matches[0]


def narration_for_ref(ref: Any, paragraphs: list[str]) -> str:
    """¶0 -> paragraf 0, ¶16b -> paragraf 16, ¶102-103 -> paragraf 102 s.d. 103."""
    match = re.search(r"¶\s*(\d+)(?:\s*-\s*(\d+))?", str(ref or ""))
    if not match:
        return ""
    start = int(match.group(1))
    end = int(match.group(2) or start)
    if end < start:
        start, end = end, start
    return "\n".join(paragraphs[index] for index in range(start, end + 1) if index < len(paragraphs))


# ---------------------------------------------------------------------------
# Penanda Excel: STATUS, FILE GAMBAR, NARASI, WAKTU
# ---------------------------------------------------------------------------

class ExcelMarker:
    """Tulis status tiap baris langsung ke Excel sumber setelah baris selesai.

    Jika Excel sedang terbuka (terkunci), penanda disimpan ke salinan
    `<nama>_TANDA.xlsx` di folder hasil agar proses generate tidak terhenti.
    """

    FILLS = {
        "selesai": "C6EFCE",   # hijau
        "sebagian": "FFEB9C",  # kuning
        "gagal": "FFC7CE",     # merah
        "ditolak": "F8CBAD",   # oranye
    }

    def __init__(self, input_path: Path, config: dict[str, Any]) -> None:
        from openpyxl.styles import Alignment, Font, PatternFill

        self._alignment, self._font, self._fill = Alignment, Font, PatternFill
        self.path = input_path
        self.enabled = input_path.suffix.casefold() in {".xlsx", ".xlsm"}
        self.expected = int(config.get("download", {}).get("outputs_per_row", 2))
        self.fallback = resolve_path(config["download"]["folder"]) / f"{input_path.stem}_TANDA{input_path.suffix}"
        self.warned_locked = False
        self.backed_up = False
        if not self.enabled:
            logging.warning("Penanda Excel hanya untuk .xlsx/.xlsm; %s tidak ditandai", input_path.name)
            return
        self.workbook = load_workbook(input_path, keep_vba=input_path.suffix.casefold() == ".xlsm")
        sheet_name = config.get("sheet_name")
        self.sheet = self.workbook[sheet_name] if sheet_name else self.workbook.active
        headers = {}
        last_column = 0
        for cell in self.sheet[1]:
            if cell.value is not None and str(cell.value).strip():
                headers[str(cell.value).strip().casefold()] = cell.column
                last_column = max(last_column, cell.column)
        self.headers = headers
        self.columns: dict[str, int] = {}
        widths = {MARK_STATUS: 16, MARK_FILES: 34, MARK_NARRATION: 70, MARK_TIME: 18}
        for name in MARK_COLUMNS:
            column = headers.get(name.casefold())
            if column is None:
                last_column += 1
                column = last_column
                header = self.sheet.cell(row=1, column=column, value=name)
                header.font = Font(bold=True, color="FFFFFF")
                header.fill = PatternFill("solid", fgColor="305496")
                header.alignment = Alignment(horizontal="center", vertical="center")
                letter = header.column_letter
                self.sheet.column_dimensions[letter].width = widths[name]
            self.columns[name] = column

    def fill_narration(self) -> None:
        """Isi kolom NARASI untuk semua baris dari .docx yang cocok (sekali di awal)."""
        if not self.enabled:
            return
        docx_path = find_narration_docx(self.path)
        if docx_path is None:
            logging.info(
                "NARASI | tidak ada .docx yang cocok di folder %s; kolom NARASI dibiarkan",
                self.path.parent,
            )
            return
        ref_column = None
        for name in SOURCE_REF_HEADERS:
            if name in self.headers:
                ref_column = self.headers[name]
                break
        if ref_column is None:
            logging.info("NARASI | kolom 'Source Ref' tidak ada; kolom NARASI dibiarkan")
            return
        paragraphs = docx_paragraphs(docx_path)
        filled = missing = 0
        column = self.columns[MARK_NARRATION]
        for row in range(2, self.sheet.max_row + 1):
            ref = self.sheet.cell(row=row, column=ref_column).value
            if is_empty(ref):
                continue
            cell = self.sheet.cell(row=row, column=column)
            if not is_empty(cell.value):
                continue
            text = narration_for_ref(ref, paragraphs)
            if text:
                cell.value = text
                cell.alignment = self._alignment(wrap_text=True, vertical="top")
                filled += 1
            else:
                missing += 1
        logging.info(
            "NARASI | %s | %s baris diisi dari %s (%s paragraf)%s",
            self.path.name, filled, docx_path.name, len(paragraphs),
            f"; {missing} Source Ref tidak dikenali" if missing else "",
        )
        self.save()

    def mark(self, row_number: int, files: list[str], note: str = "", rejected: bool = False) -> None:
        if not self.enabled:
            return
        files_cell = self.sheet.cell(row=row_number, column=self.columns[MARK_FILES])
        previous = [line for line in str(files_cell.value or "").splitlines() if line.strip()]
        all_files = previous + [name for name in files if name not in previous]
        # Hitung dari semua file scene ini, termasuk hasil run sebelumnya.
        count = len(all_files)
        if count >= self.expected:
            status, kind = f"SELESAI ({count}/{self.expected})", "selesai"
        elif count > 0:
            status, kind = f"SEBAGIAN ({count}/{self.expected})", "sebagian"
        elif rejected:
            status, kind = "DITOLAK KEBIJAKAN FLOW - ubah prompt lalu kosongkan STATUS", "ditolak"
        else:
            status, kind = f"GAGAL (0/{self.expected})", "gagal"
        if note:
            status = f"{status} - {note}"[:250]
        fill = self._fill("solid", fgColor=self.FILLS[kind])
        status_cell = self.sheet.cell(row=row_number, column=self.columns[MARK_STATUS], value=status)
        status_cell.fill = fill
        status_cell.font = self._font(bold=True)
        status_cell.alignment = self._alignment(vertical="top", wrap_text=True)
        files_cell.value = "\n".join(all_files) if all_files else None
        files_cell.fill = fill
        files_cell.alignment = self._alignment(vertical="top", wrap_text=True)
        self.sheet.cell(row=row_number, column=1).fill = fill
        time_cell = self.sheet.cell(
            row=row_number, column=self.columns[MARK_TIME], value=time.strftime("%Y-%m-%d %H:%M:%S")
        )
        time_cell.alignment = self._alignment(vertical="top")
        self.save()
        logging.info("TANDA EXCEL | baris %s | %s | %s", row_number, status, ", ".join(files) or "-")

    def backup_original(self) -> None:
        """Simpan salinan Excel asli satu kali sebelum bot menulis tanda."""
        backup = self.fallback.with_name(f"{self.path.stem}_ASLI{self.path.suffix}")
        if backup.exists():
            return
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.path, backup)
        logging.info("TANDA EXCEL | cadangan Excel asli disimpan: %s", backup)

    def save(self) -> None:
        if not self.enabled:
            return
        if not self.backed_up:
            self.backup_original()
            self.backed_up = True
        # Tulis ke file sementara lalu ganti, supaya Excel tidak rusak jika
        # bot dihentikan tepat saat menyimpan.
        temp = self.path.with_name(self.path.name + ".tmp")
        try:
            self.workbook.save(temp)
            os.replace(temp, self.path)
            return
        except PermissionError:
            try:
                temp.unlink(missing_ok=True)
            except OSError:
                pass
        self.fallback.parent.mkdir(parents=True, exist_ok=True)
        self.workbook.save(self.fallback)
        if not self.warned_locked:
            logging.warning(
                "TANDA EXCEL | %s sedang dibuka/terkunci; tanda disimpan ke %s. "
                "Tutup Excel agar file asli ikut diperbarui.",
                self.path.name, self.fallback,
            )
            self.warned_locked = True


def open_marker(input_path: Path, config: dict[str, Any]) -> ExcelMarker | None:
    try:
        marker = ExcelMarker(input_path, config)
        marker.fill_narration()
        return marker
    except Exception:
        logging.exception("Penanda Excel tidak dapat disiapkan untuk %s; generate tetap berjalan", input_path.name)
        return None


def safe_mark(
    marker: ExcelMarker | None, row_number: int, files: list[str], note: str = "", rejected: bool = False
) -> None:
    if marker is None:
        return
    try:
        marker.mark(row_number, files, note, rejected)
    except Exception:
        logging.exception("Baris %s: gagal menulis tanda ke Excel", row_number)


def safe_screenshot(page: "Page", path: Path | str, full_page: bool = False) -> bool:
    """Screenshot hanya untuk catatan; TIDAK boleh menggagalkan baris.
    Saat jendela Chrome diminimalkan, screenshot bisa macet -> batasi 8 detik."""
    try:
        page.screenshot(path=str(path), full_page=full_page, timeout=8000)
        return True
    except Exception as exc:
        logging.warning("Screenshot dilewati (%s). Jangan minimalkan jendela Chrome bot.", str(exc).splitlines()[0][:80])
        return False


def first_visible(page: "Page", selectors: list[str], timeout_ms: int = 30000):
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for selector in selectors:
            candidates = page.locator(selector)
            for index in range(candidates.count()):
                candidate = candidates.nth(index)
                if candidate.is_visible():
                    return candidate
        # Pada varian UI tertentu tulisan pencarian adalah elemen kustom,
        # bukan atribut placeholder pada <input>.
        labels = page.get_by_text(re.compile(r"Telusuri aset|Search assets", re.IGNORECASE), exact=False)
        for index in range(labels.count()):
            candidate = labels.nth(index)
            if candidate.is_visible():
                return candidate
        page.wait_for_timeout(500)
    raise RuntimeError(f"Elemen tidak ditemukan. Selector yang dicoba: {selectors}")


_LOGGED_ACTION_SELECTORS: set[tuple[str, str]] = set()


def prompt_action_button(page: "Page", config: dict[str, Any], side: str):
    """Cari tombol aksi prompt lewat identitas elemen, container, lalu posisi lama."""
    prompt = first_visible(page, config["prompt"]["selectors"])
    prompt_box = prompt.bounding_box()
    if prompt_box is None:
        raise RuntimeError("Posisi kotak prompt tidak dapat dibaca")

    # Cara utama: atribut aksesibilitas/ikon. Ini tidak bergantung pada posisi
    # tombol dan tetap bekerja saat Flow memindahkan baris aksi prompt.
    selector_key = "left_action_selectors" if side == "left" else "right_action_selectors"
    stable_selectors = config.get("prompt", {}).get(selector_key, [])
    for selector in stable_selectors:
        matches = page.locator(selector)
        for index in range(matches.count()):
            button = matches.nth(index)
            try:
                if button.is_visible(timeout=300) and button.is_enabled(timeout=300):
                    box = button.bounding_box()
                    # Tombol harus berada di kotak prompt, bukan tombol "+" di header.
                    if box is not None and not (
                        prompt_box["y"] - 160 <= box["y"] + box["height"] / 2
                        <= prompt_box["y"] + prompt_box["height"] + 120
                    ):
                        continue
                    identity = " ".join([
                        button.inner_text(timeout=300) or "",
                        button.get_attribute("aria-label", timeout=300) or "",
                        button.get_attribute("title", timeout=300) or "",
                    ]).casefold()
                    if "add media menu" in identity:
                        continue
                    identity_words = set(re.findall(r"[a-z0-9_]+|\+", identity))
                    if side == "left" and identity_words.intersection({"arrow_forward", "send", "generate"}):
                        continue
                    if side == "right" and identity_words.intersection({"+", "add", "add_2", "add_circle"}):
                        continue
                    key = (side, selector)
                    if key not in _LOGGED_ACTION_SELECTORS:
                        _LOGGED_ACTION_SELECTORS.add(key)
                        logging.info("Tombol aksi %s ditemukan dengan selector stabil: %s", side, selector)
                    return button
            except Exception:
                continue

    # Varian Flow lain memakai komponen ikon tanpa class material-symbols.
    # Baca nama/teks tombol lalu pilih kandidat semantik yang paling dekat
    # dengan editor prompt agar tidak tertukar dengan tombol + di header.
    semantic_candidates = []
    semantic_tokens = (
        {"+", "add", "add_2", "add_circle", "tambah"}
        if side == "left"
        else {"arrow_forward", "send", "generate"}
    )
    all_buttons = page.locator("button")
    prompt_center_x = prompt_box["x"] + prompt_box["width"] / 2
    prompt_center_y = prompt_box["y"] + prompt_box["height"] / 2
    for index in range(all_buttons.count()):
        button = all_buttons.nth(index)
        try:
            if not button.is_visible(timeout=200) or not button.is_enabled(timeout=200):
                continue
            parts = [
                button.inner_text(timeout=300) or "",
                button.get_attribute("aria-label", timeout=300) or "",
                button.get_attribute("title", timeout=300) or "",
                button.get_attribute("data-tooltip", timeout=300) or "",
            ]
            normalized = " ".join(parts).strip().casefold()
            # Sertakan angka karena ikon prompt pada UI baru bernama `add_2`.
            words = set(re.findall(r"[a-z0-9_]+|\+", normalized))
            if not words.intersection(semantic_tokens):
                continue
            box = button.bounding_box()
            if box is None:
                continue
            center_x = box["x"] + box["width"] / 2
            center_y = box["y"] + box["height"] / 2
            distance = (center_x - prompt_center_x) ** 2 + (center_y - prompt_center_y) ** 2
            semantic_candidates.append((distance, button, normalized[:80]))
        except Exception:
            continue
    if semantic_candidates:
        semantic_candidates.sort(key=lambda item: item[0])
        logging.info(
            "Tombol aksi %s ditemukan dari identitas ikon: %s",
            side, semantic_candidates[0][2],
        )
        return semantic_candidates[0][1]

    # Cara kedua: cari container terkecil yang membungkus editor prompt dan
    # tombol-tombolnya. Ini tahan terhadap prompt yang berubah tinggi/lebar.
    containers = prompt.locator("xpath=ancestor::*")
    container_candidates = []
    for container_index in range(containers.count()):
        container = containers.nth(container_index)
        try:
            if not container.is_visible(timeout=200):
                continue
            container_box = container.bounding_box()
            if container_box is None:
                continue
            local_buttons = container.locator("button")
            visible_buttons = []
            for button_index in range(local_buttons.count()):
                button = local_buttons.nth(button_index)
                if not button.is_visible(timeout=200) or not button.is_enabled(timeout=200):
                    continue
                box = button.bounding_box()
                if box is not None:
                    visible_buttons.append((box["x"] + box["width"] / 2, button))
            if 2 <= len(visible_buttons) <= 20:
                area = container_box["width"] * container_box["height"]
                container_candidates.append((area, visible_buttons))
        except Exception:
            continue
    if container_candidates:
        container_candidates.sort(key=lambda item: item[0])
        local = sorted(container_candidates[0][1], key=lambda item: item[0])
        logging.info("Tombol aksi %s ditemukan dari container prompt", side)
        return local[0][1] if side == "left" else local[-1][1]

    # Cara lama tetap dipertahankan sebagai cadangan untuk UI sebelumnya.
    candidates = []
    buttons = page.locator("button")
    for index in range(buttons.count()):
        button = buttons.nth(index)
        try:
            visible = button.is_visible(timeout=500)
            enabled = button.is_enabled(timeout=500)
        except Exception:
            continue
        if not visible or not enabled:
            continue
        box = button.bounding_box()
        if box is None:
            continue
        center_x = box["x"] + box["width"] / 2
        center_y = box["y"] + box["height"] / 2
        within_x = prompt_box["x"] - 30 <= center_x <= prompt_box["x"] + prompt_box["width"] + 30
        near_bottom = prompt_box["y"] + prompt_box["height"] - 100 <= center_y <= prompt_box["y"] + prompt_box["height"] + 100
        if within_x and near_bottom:
            candidates.append((center_x, button))
    if not candidates:
        raise RuntimeError("Tombol aksi prompt tidak ditemukan oleh selector, container, maupun posisi")
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1] if side == "left" else candidates[-1][1]


def click_prompt_action(page: "Page", config: dict[str, Any], side: str) -> None:
    """Klik tombol prompt dengan locator segar dan batas waktu pendek.

    Flow kadang mengganti tombol +/Generate tepat setelah locator ditemukan.
    Klik Playwright biasa lalu menunggu 30 detik pada elemen lama yang sudah
    disabled. Cari ulang tombol setiap percobaan supaya perubahan render itu
    pulih cepat tanpa melewati karakter.
    """
    prompt_config = config.get("prompt", {})
    attempts = max(1, int(prompt_config.get("action_click_retry_count", 4)))
    timeout_ms = max(500, int(prompt_config.get("action_click_timeout_ms", 2500)))
    last_error: Exception | None = None

    for attempt in range(1, attempts + 1):
        try:
            button = prompt_action_button(page, config, side)
            if not button.is_visible(timeout=400) or not button.is_enabled(timeout=400):
                raise RuntimeError("tombol berubah nonaktif sebelum diklik")
            button.click(timeout=timeout_ms)
            logging.info("Tombol aksi %s berhasil diklik pada percobaan %d/%d", side, attempt, attempts)
            return
        except Exception as exc:
            last_error = exc
            logging.warning(
                "Klik tombol aksi %s belum siap pada percobaan %d/%d: %s",
                side, attempt, attempts, exc,
            )
            page.wait_for_timeout(350 * attempt)

    raise RuntimeError(
        f"Tombol aksi {side} gagal diklik setelah {attempts} percobaan; "
        f"penyebab terakhir: {last_error}"
    )


def click_nearest_text(page: "Page", texts: list[str], anchor, timeout_ms: int = 30000) -> None:
    """Tunggu kategori modal selesai dimuat, lalu klik yang terdekat."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        try:
            anchor_box = anchor.bounding_box()
        except Exception:
            anchor_box = None
        if anchor_box is None:
            time.sleep(0.5)
            continue
        anchor_x = anchor_box["x"] + anchor_box["width"] / 2
        anchor_y = anchor_box["y"] + anchor_box["height"] / 2
        candidates = []
        for text_value in texts:
            matches = page.get_by_text(text_value, exact=True)
            if matches.count() == 0:
                matches = page.get_by_text(text_value, exact=False)
            for index in range(matches.count()):
                candidate = matches.nth(index)
                try:
                    if not candidate.is_visible():
                        continue
                    candidate_text = (candidate.inner_text() or "").strip()
                    if len(candidate_text) > 100:
                        continue
                    box = candidate.bounding_box()
                except Exception:
                    continue
                if box is None:
                    continue
                center_x = box["x"] + box["width"] / 2
                center_y = box["y"] + box["height"] / 2
                distance = (center_x - anchor_x) ** 2 + (center_y - anchor_y) ** 2
                candidates.append((distance, candidate))
        if candidates:
            candidates.sort(key=lambda item: item[0])
            candidates[0][1].click()
            return
        time.sleep(0.5)
    raise RuntimeError(f"Kategori aset tidak ditemukan setelah menunggu panel dimuat: {texts}")


def modal_asset_search(page: "Page", asset: dict[str, Any], timeout_ms: int = 30000):
    """Ambil pencarian panel @; tombol Add tidak selalu ada pada UI baru."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for selector in asset["search_selectors"]:
            matches = page.locator(selector)
            for index in range(matches.count()):
                candidate = matches.nth(index)
                if candidate.is_visible():
                    return candidate
        page.wait_for_timeout(400)
    raise RuntimeError("Kotak pencarian dalam dialog @ tidak ditemukan")


def open_asset_picker(page: "Page", config: dict[str, Any]):
    """Buka panel media dengan retry karena UI Flow kadang terlambat merender dialog."""
    asset = config["asset_picker"]
    attempts = max(1, int(asset.get("open_retry_count", 3)))
    wait_ms = max(3000, int(asset.get("open_retry_timeout_ms", 12000)))
    last_error: Exception | None = None

    for attempt in range(1, attempts + 1):
        try:
            # Jika panel sebenarnya sudah terbuka tetapi baru selesai dirender,
            # jangan klik tombol + sekali lagi karena itu justru dapat menutupnya.
            try:
                return modal_asset_search(page, asset, 1200)
            except RuntimeError:
                pass

            close_settings_panel(page)  # panel model yang terbuka menghalangi klik prompt
            prompt = first_visible(page, config["prompt"]["selectors"], 10000)
            prompt.click(timeout=10000)
            click_prompt_action(page, config, "left")
            search = modal_asset_search(page, asset, wait_ms)
            logging.info("Panel aset siap pada percobaan %d/%d", attempt, attempts)
            return search
        except Exception as exc:
            last_error = exc
            logging.warning(
                "Panel aset belum siap pada percobaan %d/%d: %s",
                attempt, attempts, exc,
            )
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            page.wait_for_timeout(1000 + (attempt * 500))

    raise RuntimeError(
        f"Panel aset gagal dibuka setelah {attempts} percobaan; "
        f"penyebab terakhir: {last_error}"
    )


def click_character_category(page: "Page", timeout_ms: int = 30000) -> None:
    """Klik tab Karakter di dialog aset, bukan tombol sidebar belakang."""
    deadline = time.monotonic() + timeout_ms / 1000
    pattern = re.compile(r"Karakter|Character", re.IGNORECASE)
    while time.monotonic() < deadline:
        tabs = page.locator("[role='tab']").filter(has_text=pattern)
        for index in range(tabs.count()):
            tab = tabs.nth(index)
            try:
                if not tab.is_visible(timeout=300):
                    continue
                if (tab.get_attribute("aria-selected", timeout=300) or "").casefold() == "true":
                    return
                if tab.is_enabled(timeout=300):
                    tab.click()
                    return
            except Exception:
                continue
        page.wait_for_timeout(500)
    raise RuntimeError("Tab Karakter pada dialog aset tidak ditemukan atau belum aktif")


def asset_panel(page: "Page", search):
    """Cari panel modal terdekat tanpa bergantung pada tombol bagian bawah."""
    # Pilih leluhur terkecil yang mencakup tiga bagian dialog: pencarian,
    # kategori kiri, dan tombol Add. role=dialog Flow kadang hanya membungkus
    # sisi kanan sehingga tidak cukup untuk mencari kategori.
    ancestors = search.locator("xpath=ancestor::*")
    candidates = []
    for index in range(ancestors.count()):
        candidate = ancestors.nth(index)
        try:
            if not candidate.is_visible():
                continue
            text = (candidate.inner_text() or "").casefold()
            box = candidate.bounding_box()
        except Exception:
            continue
        has_category = "karakter" in text or "character" in text
        has_add = "tambahkan ke perintah" in text or "add to prompt" in text
        if has_category and has_add and box is not None:
            candidates.append((box["width"] * box["height"], candidate))
    if candidates:
        candidates.sort(key=lambda item: item[0])
        return candidates[0][1]
    return page.locator("body")


def prompt_reference_count(page: "Page", config: dict[str, Any]) -> int:
    """Hitung thumbnail/chip referensi yang benar-benar terpasang di prompt.

    Flow menutup dialog aset sebelum thumbnail selesai dirender. Karena itu status
    dialog tertutup tidak boleh dipakai sebagai tanda bahwa karakter sudah masuk.
    """
    prompt = first_visible(page, config["prompt"]["selectors"], 10000)
    return int(prompt.evaluate(
        """
        editor => {
          const visible = el => {
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          };
          const identity = el => [
            el.innerText || '', el.getAttribute('aria-label') || '',
            el.getAttribute('title') || '', el.getAttribute('data-tooltip') || ''
          ].join(' ').trim().toLowerCase();
          let node = editor;
          let best = null;
          for (let depth = 0; node && depth < 12; depth++, node = node.parentElement) {
            if (!visible(node)) continue;
            const buttons = [...node.querySelectorAll('button')].filter(visible);
            const ids = buttons.map(identity);
            const hasLeft = ids.some(v => /(^|\\s)(add|add_2|add_circle|tambah)(\\s|$)|tambahkan media/.test(v));
            const hasRight = ids.some(v => /arrow_forward|generate|buat|kirim|send/.test(v));
            if (!hasLeft || !hasRight) continue;
            const r = node.getBoundingClientRect();
            const area = r.width * r.height;
            if (!best || area < best.area) best = {node, area};
          }
          const root = best ? best.node : editor.parentElement;
          if (!root) return 0;
          const rootRect = root.getBoundingClientRect();
          const media = [...root.querySelectorAll('img, [style*="background-image"]')].filter(el => {
            if (!visible(el)) return false;
            const r = el.getBoundingClientRect();
            return r.width >= 18 && r.height >= 18 && r.width <= 140 && r.height <= 140 &&
              r.left >= rootRect.left - 2 && r.right <= rootRect.right + 2 &&
              r.top >= rootRect.top - 2 && r.bottom <= rootRect.bottom + 2;
          });
          // Tombol "Hapus perintah" selalu ada walaupun belum ada referensi,
          // sehingga tidak boleh dihitung sebagai chip. Thumbnail referensi Flow
          // memakai <img alt="Gambar referensi karakter"> dan sudah tercakup di atas.
          return media.length;
        }
        """
    ))


def wait_reference_attached(
    page: "Page", config: dict[str, Any], before_count: int, name: str, text: str
) -> int:
    """Tunggu sampai jumlah thumbnail referensi bertambah; jangan kirim prompt lebih awal."""
    timeout_ms = int(config.get("asset_picker", {}).get("attach_timeout_ms", 15000))
    deadline = time.monotonic() + timeout_ms / 1000
    last_count = before_count
    while time.monotonic() < deadline:
        try:
            last_count = prompt_reference_count(page, config)
            if last_count > before_count:
                logging.info(
                    "Referensi %s terverifikasi pada prompt: %s (thumbnail %d -> %d)",
                    name, text, before_count, last_count,
                )
                return last_count
        except Exception:
            pass
        page.wait_for_timeout(300)
    debug_dir = APP_DIR / "runtime" / "screenshots"
    debug_dir.mkdir(parents=True, exist_ok=True)
    try:
        safe_screenshot(page, debug_dir / "reference-attach-timeout.png")
        prompt = first_visible(page, config["prompt"]["selectors"], 3000)
        (debug_dir / "reference-attach-timeout.html").write_text(
            prompt.evaluate("el => el.parentElement.parentElement.parentElement.outerHTML"),
            encoding="utf-8",
        )
    except Exception:
        pass
    raise RuntimeError(
        f"Referensi {name} belum terpasang setelah {timeout_ms} ms: {text} "
        f"(thumbnail tetap {last_count})"
    )


def resolve_alias(value: Any, config: dict[str, Any]) -> str:
    """Nama karakter setelah character_aliases (bila diatur di config)."""
    text = " ".join(str(value).split())
    aliases = {
        " ".join(str(key).split()).casefold(): str(target).strip()
        for key, target in (config.get("character_aliases") or {}).items()
        if str(target or "").strip()
    }
    return aliases.get(text.casefold(), text)


def add_reference(page: "Page", name: str, value: Any, config: dict[str, Any]) -> bool:
    text = str(value).strip()
    # Nama karakter di Excel bisa berbeda dengan nama aset di Flow,
    # mis. "Kapten Lei" (Excel) = "Chen Erniu" (aset Flow).
    aliases = {
        str(key).strip().casefold(): str(target).strip()
        for key, target in (config.get("character_aliases") or {}).items()
        if str(target or "").strip()
    }
    alias = aliases.get(text.casefold())
    if alias and alias.casefold() != text.casefold():
        logging.info("Referensi %s: '%s' memakai aset Flow '%s' (character_aliases)", name, text, alias)
        text = alias
    asset = config["asset_picker"]
    project_url = page.url
    before_count = prompt_reference_count(page, config)
    # UI Flow terbaru tidak lagi membuka panel aset saat karakter "@" diketik.
    # Tombol + / Buat di kiri bawah prompt adalah pembuka panel media resmi.
    search = open_asset_picker(page, config)
    search_box = search.bounding_box()
    if search_box is None:
        raise RuntimeError("Posisi pencarian dialog @ tidak dapat dibaca")
    click_character_category(page)
    page.wait_for_timeout(300)
    search = modal_asset_search(page, asset)
    try:
        search.fill(text)
    except Exception:
        search.click()
        page.keyboard.press("Control+A")
        page.keyboard.type(text)
    page.wait_for_timeout(int(asset.get("search_wait_ms", 700)))

    panel = asset_panel(page, search)

    # Hasil pencarian biasanya menampilkan nama aset sebagai teks. Hindari input
    # pencarian itu sendiri dengan membatasi ke elemen yang terlihat dan bukan input.
    result = panel.get_by_text(text, exact=True)
    if result.count() == 0:
        result = panel.get_by_text(text, exact=False)
    chosen = None
    for index in range(result.count()):
        candidate = result.nth(index)
        if candidate.is_visible() and candidate.evaluate("el => el.tagName !== 'INPUT'"):
            candidate_text = candidate.inner_text().strip()
            if text.casefold() in candidate_text.casefold() and len(candidate_text) <= 160:
                chosen = candidate
                break
    if chosen is not None:
        chosen.click()
        # Dialog dapat tertutup lebih dulu sementara thumbnail masih diproses.
        # Tunggu render awal sebelum memeriksa status pemasangan.
        page.wait_for_timeout(1000)
    else:
        logging.warning(
            "Referensi %s tidak tersedia di panel Karakter: %s — referensi dilewati, prompt tetap diproses",
            name, text,
        )
        page.keyboard.press("Escape")
        first_visible(page, config["prompt"]["selectors"], 10000)
        return False
    if "/project/" not in page.url:
        try:
            page.go_back(wait_until="domcontentloaded")
        finally:
            raise RuntimeError(f"Pemilihan {name} keluar dari halaman project: {text}")
    # Pada UI Flow saat ini, klik hasil di dalam panel Karakter langsung
    # memasukkan thumbnail ke prompt dan menutup panel tanpa tombol Add.
    prompt_candidates = page.locator(", ".join(config["prompt"]["selectors"]))
    prompt_visible = any(prompt_candidates.nth(i).is_visible() for i in range(prompt_candidates.count()))
    if not search.is_visible() and prompt_visible:
        wait_reference_attached(page, config, before_count, name, text)
        return True
    try:
        add_button = first_visible(page, asset["add_selectors"], 3000)
    except RuntimeError:
        add_button = None
    if add_button is None:
        raise RuntimeError(f"Aset {name} belum masuk ke prompt dan tombol Tambahkan tidak tersedia: {text}")
    if not add_button.is_enabled():
        raise RuntimeError(f"Aset {name} tidak dapat dipilih: {text}")
    add_button.click()
    first_visible(page, config["prompt"]["selectors"], 10000)
    if page.url != project_url and "/project/" not in page.url:
        raise RuntimeError(f"Referensi {name} gagal kembali ke prompt: {text}")
    wait_reference_attached(page, config, before_count, name, text)
    return True


def fill_prompt(page: "Page", value: Any, config: dict[str, Any]):
    prompt = first_visible(
        page,
        config["prompt"]["selectors"],
        int(config.get("ready_timeout_ms", 120000)),
    )
    prompt.fill(str(value).strip())
    return prompt


def normalize_ws(value: Any) -> str:
    """Bandingkan teks tanpa peduli spasi/baris baru (editor Flow memecah paragraf)."""
    return " ".join(str(value or "").split())


def prompt_text_value(prompt) -> str:
    """Baca teks textarea maupun editor contenteditable milik Flow."""
    tag_name = (prompt.evaluate("el => el.tagName") or "").casefold()
    if tag_name in {"textarea", "input"}:
        return (prompt.input_value() or "").strip()
    return (prompt.inner_text() or "").strip()


def wait_prompt_bundle_ready(
    page: "Page",
    config: dict[str, Any],
    expected_prompt: Any,
    expected_references: int,
    row_number: int,
) -> None:
    """Pastikan teks dan seluruh thumbnail stabil sebelum tombol Generate ditekan."""
    expected_text = str(expected_prompt).strip()
    timeout_ms = int(config.get("prompt", {}).get("bundle_timeout_ms", 20000))
    stable_checks_required = int(config.get("prompt", {}).get("stable_checks", 5))
    deadline = time.monotonic() + timeout_ms / 1000
    stable_checks = 0
    refill_count = 0
    last_text = ""
    last_references = -1

    while time.monotonic() < deadline:
        prompt = first_visible(page, config["prompt"]["selectors"], 3000)
        try:
            last_text = prompt_text_value(prompt)
            last_references = prompt_reference_count(page, config)
        except Exception:
            stable_checks = 0
            page.wait_for_timeout(400)
            continue

        if normalize_ws(last_text) != normalize_ws(expected_text):
            stable_checks = 0
            if refill_count < 3:
                logging.warning(
                    "Baris %s: teks prompt belum lengkap/hilang; diisi ulang (%d/3)",
                    row_number, refill_count + 1,
                )
                prompt.fill(expected_text)
                refill_count += 1
            page.wait_for_timeout(500)
            continue

        if last_references != expected_references:
            stable_checks = 0
            page.wait_for_timeout(400)
            continue

        stable_checks += 1
        if stable_checks >= stable_checks_required:
            logging.info(
                "Baris %s: paket prompt siap dan stabil — teks %d karakter, "
                "%d thumbnail, %d/%d pemeriksaan lolos",
                row_number, len(last_text), last_references,
                stable_checks, stable_checks_required,
            )
            return
        page.wait_for_timeout(400)

    raise RuntimeError(
        f"Baris {row_number}: paket prompt belum stabil setelah {timeout_ms} ms "
        f"(teks cocok={normalize_ws(last_text) == normalize_ws(expected_text)}, thumbnail "
        f"{last_references}/{expected_references})"
    )


NEW_UI_IMAGE = "img[data-media-id]:not([alt*='Character' i]):not([alt*='Karakter' i])"


def new_ui_tile_order(page: "Page", limit: int = 12) -> list[str]:
    """media-id gambar di urutan teratas grid Flow baru (hasil terbaru di depan)."""
    try:
        return page.evaluate(
            """limit => {
                const tiles = [...document.querySelectorAll('flow-grid-tile-container')].slice(0, limit);
                const ids = [];
                for (const tile of tiles) {
                    const img = tile.querySelector('img[data-media-id]');
                    if (img && !/character|karakter/i.test(img.alt || '')) ids.push(img.dataset.mediaId);
                }
                return ids;
            }""",
            limit,
        )
    except Exception:
        return []


def overlay_text(page: "Page") -> str:
    """Teks toast/dialog Flow saja. Kartu grid lama (mis. "usage limit" kemarin)
    tidak ikut dibaca agar tidak salah dianggap limit baru."""
    try:
        return page.evaluate(
            """() => [...document.querySelectorAll(
                '.cdk-overlay-container, [role=alert], [role=alertdialog], [role=dialog], [role=status], mat-snack-bar-container'
            )].map(el => el.innerText || '').join('\\n')"""
        )
    except Exception:
        return ""


def top_tile_states(page: "Page", limit: int) -> list[dict[str, Any]]:
    """Status kartu teratas grid: selesai (id), gagal (Failed), atau masih proses (%)."""
    try:
        return page.evaluate(
            """limit => [...document.querySelectorAll('flow-grid-tile-container')].slice(0, limit).map(tile => {
                const text = tile.innerText || '';
                const img = tile.querySelector('img[data-media-id]');
                return {
                    id: img ? img.dataset.mediaId : null,
                    failed: /\\bFailed\\b|Gagal|violate our polic|melanggar kebijakan/i.test(text),
                    limit: /usage limit|reached your (?:\\w+ )?limit|daily limit|batas penggunaan|kuota/i.test(text),
                    blocked: /unusual activity|aktivitas (?:yang )?(?:tidak biasa|mencurigakan)/i.test(text),
                    pending: /\\b\\d{1,3}%/.test(text) && !img,
                };
            })""",
            limit,
        )
    except Exception:
        return []


def scroll_grid_top(page: "Page") -> None:
    try:
        page.evaluate(
            """() => {
                document.querySelectorAll('cdk-virtual-scroll-viewport, .cdk-virtual-scrollable')
                    .forEach(el => { el.scrollTop = 0; });
                window.scrollTo(0, 0);
            }"""
        )
    except Exception:
        pass


def generated_media(page: "Page") -> dict[str, Any]:
    media = {}
    # UI Flow baru (flow.google.com): <img data-media-id="..."> pada setiap kartu.
    images = page.locator(NEW_UI_IMAGE)
    for index in range(images.count()):
        image = images.nth(index)
        try:
            media_id = image.get_attribute("data-media-id", timeout=500)
        except Exception:
            continue
        if media_id:
            media[media_id] = page.locator(f"img[data-media-id='{media_id}']").first
    if media:
        return media
    images = page.locator("img[alt='Gambar yang dihasilkan'], img[alt='Generated image']")
    for index in range(images.count()):
        image = images.nth(index)
        if not image.is_visible():
            continue
        src = image.get_attribute("src") or ""
        match = re.search(r"[?&]name=([a-zA-Z0-9-]+)", src)
        if match:
            media[match.group(1)] = image
    return media


def has_character_badge(page: "Page", image) -> bool:
    """Deteksi ikon orang/Karakter di pojok kiri atas sebuah kartu aset."""
    try:
        if image.get_attribute("data-media-id", timeout=500):
            # UI baru: kartu karakter memakai alt "Character thumbnail" dan
            # ikon accessibility_new; hasil generate tidak.
            alt = (image.get_attribute("alt", timeout=500) or "").casefold()
            tile_text = image.evaluate(
                "el => (el.closest('flow-grid-tile-container') || el.parentElement).innerText || ''"
            ).casefold()
            return "character" in alt or "karakter" in alt or "accessibility_new" in tile_text
    except Exception:
        return True
    try:
        box = image.bounding_box(timeout=500)
    except Exception:
        return True  # aman: tunggu polling berikutnya bila kartu sedang dirender ulang
    if box is None:
        return True  # aman: kartu yang tidak dapat diverifikasi jangan dihapus
    keywords = re.compile(
        r"person|people|character|karakter|accessibility|man|woman|orang",
        re.IGNORECASE,
    )
    candidates = page.locator("[aria-label], [title], [data-icon], span, i, svg")
    for index in range(candidates.count()):
        candidate = candidates.nth(index)
        try:
            if not candidate.is_visible(timeout=100):
                continue
            badge_box = candidate.bounding_box(timeout=100)
        except Exception:
            continue
        if badge_box is None:
            continue
        center_x = badge_box["x"] + badge_box["width"] / 2
        center_y = badge_box["y"] + badge_box["height"] / 2
        # Ikon karakter pada Flow berada di area kecil pojok kiri atas kartu.
        if not (
            box["x"] - 8 <= center_x <= box["x"] + 75
            and box["y"] - 8 <= center_y <= box["y"] + 75
        ):
            continue
        try:
            marker = " ".join(
                filter(
                    None,
                    [
                        candidate.get_attribute("aria-label", timeout=100),
                        candidate.get_attribute("title", timeout=100),
                        candidate.get_attribute("data-icon", timeout=100),
                        (candidate.inner_text(timeout=100) or "").strip(),
                    ],
                )
            )
            tag_name = candidate.evaluate("el => el.tagName")
        except Exception:
            continue
        is_small_svg_badge = (
            tag_name.casefold() == "svg"
            and badge_box["width"] <= 48
            and badge_box["height"] <= 48
        )
        if keywords.search(marker) or is_small_svg_badge:
            return True
    return False


def tile_more_button(page: "Page", image):
    """Tombol "More options" milik kartu (UI Flow baru)."""
    tile = image.locator("xpath=ancestor::flow-grid-tile-container[1]")
    if tile.count() == 0:
        return None
    tile.first.scroll_into_view_if_needed(timeout=5000)
    tile.first.hover()
    page.wait_for_timeout(400)
    button = tile.first.locator(
        "button[aria-label='More options'], button[aria-label*='Opsi lainnya' i], button[aria-label*='More' i]"
    )
    for index in range(button.count()):
        if button.nth(index).is_visible():
            return button.nth(index)
    return None


def nearest_more_button(page: "Page", image):
    new_ui = tile_more_button(page, image)
    if new_ui is not None:
        return new_ui
    image_box = image.bounding_box()
    if image_box is None:
        raise RuntimeError("Posisi gambar hasil tidak dapat dibaca")
    # Tombol tiga titik milik kartu baru tampil sesudah kartu di-hover.
    image.hover()
    page.wait_for_timeout(400)
    target_x = image_box["x"] + image_box["width"] - 25
    target_y = image_box["y"] + 25
    candidates = []
    buttons = page.locator("button")
    for index in range(buttons.count()):
        button = buttons.nth(index)
        if not button.is_visible():
            continue
        label = (button.inner_text() or "").casefold()
        if "lainnya" not in label and "more_vert" not in label and "more" not in label:
            continue
        box = button.bounding_box()
        if box is None:
            continue
        center_x = box["x"] + box["width"] / 2
        center_y = box["y"] + box["height"] / 2
        # Abaikan menu tiga titik global di header. Kandidat wajib berada di
        # bagian atas kartu gambar yang sedang diproses.
        if not (
            image_box["x"] <= center_x <= image_box["x"] + image_box["width"]
            and image_box["y"] <= center_y <= image_box["y"] + 90
        ):
            continue
        distance = (center_x - target_x) ** 2 + (center_y - target_y) ** 2
        candidates.append((distance, button))
    if not candidates:
        raise RuntimeError("Menu tiga titik pada gambar tidak ditemukan")
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def close_menus(page: "Page") -> None:
    for _ in range(3):
        if page.locator("[role='menu']").count() == 0:
            return
        try:
            page.keyboard.press("Escape")
        except Exception:
            return
        page.wait_for_timeout(250)


CAPTURE_DOWNLOAD_JS = """() => {
    window.__fbDl = null;
    window.__fbCapture = true;
    if (window.__fbPatched) return;
    window.__fbPatched = true;
    window.__fbBlobs = {};
    const originalCreate = URL.createObjectURL.bind(URL);
    URL.createObjectURL = function (obj) {
        const url = originalCreate(obj);
        try {
            if (window.__fbCapture && obj instanceof Blob) window.__fbBlobs[url] = obj;
        } catch (e) {}
        return url;
    };
    const record = a => {
        if (!window.__fbCapture || !a || !a.download) return false;
        window.__fbDl = {href: a.href, name: a.download};
        return true;
    };
    const originalClick = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function () {
        if (record(this)) return;
        return originalClick.apply(this, arguments);
    };
    document.addEventListener('click', event => {
        const a = event.target && event.target.closest ? event.target.closest('a[download]') : null;
        if (a && record(a)) event.preventDefault();
    }, true);
}"""

READ_BLOB_JS = """async href => {
    // Flow langsung mencabut blob URL setelah klik; pakai objek Blob yang disimpan.
    let blob = (window.__fbBlobs || {})[href];
    if (!blob) blob = await (await fetch(href)).blob();
    if (window.__fbBlobs) window.__fbBlobs = {};
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let text = '';
    for (let i = 0; i < bytes.length; i += 0x8000) {
        text += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return {b64: btoa(text), type: blob.type || ''};
}"""


def download_via_menu(page: "Page", image, row_number: int) -> tuple[bytes, str] | None:
    """Ambil file ukuran asli (1K) lewat menu kartu: More options > Download > 1K.

    Flow membuat file lewat blob lalu memicu unduhan browser. Unduhan browser
    itu dicegat dan isinya dibaca langsung dari halaman, sehingga Chrome tidak
    perlu menjalankan proses download (pada profil tertentu download membuat
    Chrome tertutup).
    """
    import base64

    try:
        more = tile_more_button(page, image)
        if more is None:
            return None
        page.evaluate(CAPTURE_DOWNLOAD_JS)
        more.click()
        page.wait_for_timeout(600)
        download_item = page.get_by_role("menuitem", name=re.compile(r"Download|Unduh", re.IGNORECASE)).first
        download_item.click()
        page.wait_for_timeout(700)
        size_item = page.get_by_role("menuitem", name=re.compile(r"^\s*1K", re.IGNORECASE)).first
        if size_item.count() == 0:
            size_item = page.locator("[role='menu']").last.get_by_text(re.compile(r"^1K$")).first
        size_item.click()
        captured = None
        deadline = time.monotonic() + 45
        while time.monotonic() < deadline:
            captured = page.evaluate("() => window.__fbDl")
            if captured:
                break
            page.wait_for_timeout(300)
        if not captured:
            raise RuntimeError("file 1K tidak muncul dari menu Download")
        payload = page.evaluate(READ_BLOB_JS, captured["href"])
        data = base64.b64decode(payload["b64"])
        name = str(captured.get("name") or "")
        suffix = Path(name).suffix.casefold()
        content_type = str(payload.get("type") or "").casefold()
        if suffix not in {".png", ".jpg", ".jpeg", ".webp"}:
            suffix = ".png" if "png" in content_type else ".webp" if "webp" in content_type else ".jpeg"
        if suffix == ".jpg":
            suffix = ".jpeg"
        if len(data) < 1000:
            raise RuntimeError(f"file 1K terlalu kecil ({len(data)} byte)")
        logging.info("Baris %s: file 1K diambil lewat menu (%s, %s KB)", row_number, name, len(data) // 1024)
        return data, suffix
    except Exception as exc:
        logging.warning("Baris %s: ambil file lewat menu gagal, memakai URL gambar: %s", row_number, exc)
        return None
    finally:
        try:
            page.evaluate("() => { window.__fbCapture = false; }")
        except Exception:
            pass
        try:
            close_menus(page)
        except Exception:
            pass


def delete_generated_output(page: "Page", media_id: str, row_number: int) -> bool:
    """Hapus satu hasil yang sudah berhasil disimpan dari galeri Flow."""
    current = generated_media(page)
    image = current.get(media_id)
    if image is None:
        logging.warning("Baris %s: hasil %s sudah tidak terlihat, penghapusan dilewati", row_number, media_id)
        return False
    if has_character_badge(page, image):
        logging.warning(
            "Baris %s: KARAKTER DILINDUNGI (%s) — ikon orang terdeteksi, kartu tidak dihapus",
            row_number, media_id,
        )
        return False
    logging.info("Baris %s: HASIL GENERATE (%s) aman untuk dihapus", row_number, media_id)
    first = None
    for attempt in range(1, 4):
        image = generated_media(page).get(media_id)
        if image is None:
            return False
        nearest_more_button(page, image).click()
        page.wait_for_timeout(700)
        trash_candidates = [
            page.get_by_role("menuitem", name=re.compile(r"Pindahkan ke sampah|Move to trash", re.IGNORECASE)),
            page.get_by_text("Pindahkan ke sampah", exact=False),
            page.get_by_text("Move to trash", exact=False),
            page.locator("[role='menu']").get_by_text(re.compile(r"sampah|trash", re.IGNORECASE)),
        ]
        for trash in trash_candidates:
            for index in range(trash.count()):
                candidate = trash.nth(index)
                if candidate.is_visible():
                    first = candidate
                    break
            if first is not None:
                break
        if first is not None:
            break
        page.keyboard.press("Escape")
        logging.warning("Baris %s: menu hapus belum muncul, percobaan %s/3", row_number, attempt)
    if first is None:
        logging.error(
            "Baris %s: GAGAL HAPUS (%s) — kartu dibiarkan agar karakter tidak berisiko terhapus",
            row_number, media_id,
        )
        return False
    first.click()
    page.wait_for_timeout(800)
    logging.info("Baris %s: hasil yang sudah diunduh dipindahkan ke sampah", row_number)
    return True


def download_new_outputs(
    page: "Page",
    before_ids: set[str],
    row_number: int,
    config: dict[str, Any],
    scene: str,
) -> list[str]:
    download = config["download"]
    expected = int(download.get("outputs_per_row", 2))
    deadline = time.monotonic() + int(download.get("generation_timeout_ms", 300000)) / 1000
    started_waiting = time.monotonic()
    last_activity = started_waiting
    new_media = {}
    saw_progress = False
    previous_media_count = 0
    terminal_reason: str | None = None
    logged_generated: set[str] = set()
    logged_protected: set[str] = set()
    failed_new = 0
    while time.monotonic() < deadline:
        page_text = overlay_text(page)
        message_type = classify_flow_message(page_text)
        if message_type == "limit":
            snippet = limit_snippet(page_text)
            logging.critical("LIMIT GOOGLE FLOW TERDETEKSI | pesan Flow: %s", snippet)
            try:
                shots = APP_DIR / "runtime" / "screenshots"
                shots.mkdir(parents=True, exist_ok=True)
                safe_screenshot(page, shots / f"limit-row-{row_number}.png")
            except Exception:
                pass
            raise FlowLimitError(f"Kredit/kuota Google Flow habis: {snippet}")
        scroll_grid_top(page)
        current = generated_media(page)
        new_media = {}
        top_row_max_y = int(download.get("top_row_max_y", 250))
        newest = set(new_ui_tile_order(page, expected * 2 + 2))
        for key, value in current.items():
            if key in before_ids:
                continue
            if newest:
                # UI baru: hasil terbaru selalu di urutan paling depan grid.
                is_top = key in newest
            else:
                box = value.bounding_box()
                # Hanya hasil baru yang sudah selesai dan tampil di baris paling atas.
                # Gambar lama yang sekadar masuk viewport setelah grid bergeser diabaikan.
                is_top = box is not None and box["y"] <= top_row_max_y
            if is_top:
                if has_character_badge(page, value):
                    if key not in logged_protected:
                        logging.info(
                            "Baris %s: KARAKTER DILINDUNGI (%s) — tidak dianggap hasil generate",
                            row_number, key,
                        )
                        logged_protected.add(key)
                    continue
                new_media[key] = value
                if key not in logged_generated:
                    logging.info("Baris %s: HASIL GENERATE terdeteksi (%s)", row_number, key)
                    logged_generated.add(key)
        progress = page.get_by_text(re.compile(r"^\d+%$"))
        visible_progress = 0
        for index in range(progress.count()):
            if progress.nth(index).is_visible():
                visible_progress += 1
        if visible_progress:
            saw_progress = True
            last_activity = time.monotonic()
        if len(new_media) != previous_media_count:
            previous_media_count = len(new_media)
            last_activity = time.monotonic()

        if len(new_media) >= expected:
            terminal_reason = "semua hasil selesai"
            break

        # UI baru: kartu hasil yang ditolak kebijakan tampil "Failed ... might
        # violate our policies" di posisi teratas grid. Hitung hanya setelah
        # kartu baru muncul (progres terlihat atau sudah > 20 detik).
        if saw_progress or time.monotonic() - started_waiting >= 20:
            states = top_tile_states(page, expected)
            if states and not any(item.get("pending") for item in states):
                blocked_now = sum(1 for item in states if item.get("blocked"))
                if blocked_now and len(new_media) + blocked_now >= expected:
                    logging.critical(
                        "FLOW MENOLAK SEMENTARA | \"We noticed some unusual activity\" (%s/%s) — "
                        "bukan masalah prompt; terlalu banyak generate otomatis dalam waktu singkat",
                        blocked_now, expected,
                    )
                    try:
                        shots = APP_DIR / "runtime" / "screenshots"
                        shots.mkdir(parents=True, exist_ok=True)
                        safe_screenshot(page, shots / f"unusual-activity-row-{row_number}.png")
                    except Exception:
                        pass
                    raise FlowBlockedError("Flow: We noticed some unusual activity")
                limit_now = sum(1 for item in states if item.get("limit"))
                if limit_now and len(new_media) + limit_now >= expected:
                    logging.critical(
                        "LIMIT GOOGLE FLOW TERDETEKSI | kartu hasil: \"You've reached your usage limit\" (%s/%s)",
                        limit_now, expected,
                    )
                    try:
                        shots = APP_DIR / "runtime" / "screenshots"
                        shots.mkdir(parents=True, exist_ok=True)
                        safe_screenshot(page, shots / f"limit-row-{row_number}.png")
                    except Exception:
                        pass
                    raise FlowLimitError("You've reached your usage limit (kartu hasil Flow)")
                failed_now = sum(
                    1 for item in states
                    if item.get("failed") and not item.get("limit") and not item.get("blocked")
                )
                if failed_now and len(new_media) + failed_now >= expected:
                    failed_new = failed_now
                    terminal_reason = "hasil ditolak kebijakan Flow"
                    logging.warning(
                        "Baris %s: %s hasil DITOLAK KEBIJAKAN FLOW (might violate our policies); "
                        "%s hasil berhasil", row_number, failed_now, len(new_media),
                    )
                    break

        if message_type == "policy" and not visible_progress:
            quiet_seconds = time.monotonic() - last_activity
            if quiet_seconds >= int(download.get("failure_settle_ms", 12000)) / 1000:
                terminal_reason = "prompt ditolak kebijakan"
                logging.warning(
                    "Baris %s: PROMPT TIDAK DIIZINKAN terdeteksi; %s hasil berhasil tetap akan diunduh",
                    row_number, len(new_media),
                )
                break

        # Indikator persen bisa hilang sesaat. Tunggu masa tenang panjang sejak
        # aktivitas/hasil terakhir sebelum menyimpulkan slot lainnya gagal.
        quiet_seconds = time.monotonic() - last_activity
        elapsed = time.monotonic() - started_waiting
        settle_seconds = int(download.get("completion_settle_ms", 60000)) / 1000
        minimum_seconds = int(download.get("minimum_generation_wait_ms", 45000)) / 1000
        if not visible_progress and elapsed >= minimum_seconds and quiet_seconds >= settle_seconds:
            terminal_reason = "slot tersisa gagal/tidak menghasilkan gambar"
            break
        page.wait_for_timeout(2000)
    else:
        terminal_reason = "batas waktu generate tercapai"

    logging.info(
        "Baris %s: proses generate terminal (%s), hasil final %s/%s",
        row_number, terminal_reason, len(new_media), expected,
    )
    if not new_media:
        logging.warning("Baris %s: semua hasil Create gagal; tidak ada file yang diunduh", row_number)
        if failed_new:
            raise FlowPolicyRejected(
                "Prompt ditolak kebijakan Flow (might violate our policies); ubah isi prompt scene ini"
            )
        return []
    if len(new_media) < expected:
        logging.warning("Baris %s: hanya %s dari %s hasil yang berhasil", row_number, len(new_media), expected)

    items = list(new_media.items())
    order = new_ui_tile_order(page, expected * 2 + 2)
    if order:
        # Urutan grid: hasil pertama di depan -> -1, berikutnya -2.
        items.sort(key=lambda item: order.index(item[0]) if item[0] in order else 999)
    else:
        items.sort(key=lambda item: ((item[1].bounding_box() or {}).get("y", 99999), (item[1].bounding_box() or {}).get("x", 99999)))
    output_dir = resolve_path(download["folder"])
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = str(download["base_name"])

    downloaded_ids = []
    saved_names: list[str] = []
    for offset, (media_id, image) in enumerate(items[:expected]):
        menu_file = download_via_menu(page, image, row_number)
        body = None
        if menu_file is not None:
            body, suffix = menu_file
        else:
            src = image.get_attribute("src")
            if not src:
                raise RuntimeError("URL gambar hasil tidak ditemukan")
            response = page.request.get(urljoin(page.url, src))
            if not response.ok:
                raise RuntimeError(f"Download gambar gagal dengan status HTTP {response.status}")
            content_type = response.headers.get("content-type", "image/jpeg").casefold()
            suffix = ".png" if "png" in content_type else ".webp" if "webp" in content_type else ".jpeg"
            body = response.body()
        # Nama per scene: "Episode_041 S001-1.jpeg", "Episode_041 S001-2.jpeg".
        # Jika scene diulang, nomor lanjut (S001-3, ...) tanpa menimpa file lama.
        existing_numbers = []
        pattern = re.compile(
            rf"^{re.escape(base_name)} {re.escape(scene)}-(\d+)\.(?:jpe?g|png|webp)$", re.IGNORECASE
        )
        for existing in output_dir.iterdir():
            match = pattern.match(existing.name)
            if match:
                existing_numbers.append(int(match.group(1)))
        number = max(existing_numbers, default=0) + 1
        destination = output_dir / f"{base_name} {scene}-{number}{suffix}"
        destination.write_bytes(body)
        logging.info("Baris %s: hasil disimpan sebagai %s", row_number, destination.name)
        downloaded_ids.append(media_id)
        saved_names.append(destination.name)

    # Hanya hasil yang sudah benar-benar tersimpan yang boleh dihapus.
    for media_id in downloaded_ids:
        delete_generated_output(page, media_id, row_number)
    return saved_names


def process_row(
    page: "Page | None",
    row_number: int,
    row: dict[str, Any],
    config: dict[str, Any],
    dry_run: bool,
    run_state: dict[str, Any] | None = None,
) -> list[str]:
    """Proses satu baris dan kembalikan daftar nama file gambar yang tersimpan."""
    scene = scene_label(row_number, row, config)
    logging.info("Memproses baris %s (%s)", row_number, scene)
    prompt_config = config.get("prompt", {})
    prompt_value = row.get(str(prompt_config.get("column", "Prompt")))
    if is_empty(prompt_value):
        logging.warning("Baris %s: Prompt kosong, seluruh baris dilewati", row_number)
        return []
    if dry_run:
        logging.info(
            "[SIMULASI] Baris %s: isi prompt (%s karakter) -> file %s %s-1.jpeg dst.",
            row_number, len(str(prompt_value)), config["download"]["base_name"], scene,
        )
        prompt = None
        before_ids: set[str] = set()
    else:
        scroll_grid_top(page)  # type: ignore[arg-type]
        before_ids = set(generated_media(page).keys())  # type: ignore[arg-type]
        before_ids.update(new_ui_tile_order(page, 12))  # type: ignore[arg-type]
        # Referensi dipasang saat editor masih kosong. Pada beberapa versi Flow,
        # membuka panel aset dapat menghapus teks yang sudah diisi lebih dahulu.
        prompt = fill_prompt(page, "", config)  # type: ignore[arg-type]

    attached_names: list[str] = []
    for name, settings in config.get("fields", {}).items():
        source_column = str(settings.get("column", name))
        value = row.get(source_column)
        if is_empty(value):
            logging.info("Baris %s: %s kosong/0, dilewati", row_number, name)
            continue
        if dry_run:
            logging.info("[SIMULASI] Baris %s: tambah referensi %s = %s", row_number, name, value)
        elif norm_name(resolve_alias(value, config)) in set(config.get("_missing_characters", [])):
            logging.info(
                "Baris %s: %s = %s tidak ada di profil ini; dilewati (karakter dari prompt)", row_number, name, value
            )
        else:
            try:
                refs_before = prompt_reference_count(page, config)  # type: ignore[arg-type]
            except Exception:
                refs_before = None
            try:
                if add_reference(page, name, value, config):  # type: ignore[arg-type]
                    attached_names.append(name)
            except FlowProjectNotFound:
                raise
            except Exception as exc:
                # Cara cadangan: generate tetap jalan tanpa referensi ini.
                logging.warning(
                    "Baris %s: referensi %s (%s) gagal dipasang (%s); lanjut tanpa referensi",
                    row_number, name, value, str(exc).splitlines()[0][:120],
                )
                try:
                    page.keyboard.press("Escape")  # type: ignore[union-attr]
                    page.wait_for_timeout(1500)  # type: ignore[union-attr]
                    first_visible(page, config["prompt"]["selectors"], 10000)  # type: ignore[arg-type]
                    # Thumbnail kadang tetap masuk terlambat -> hitung sebagai terpasang.
                    if refs_before is not None and prompt_reference_count(page, config) > refs_before:  # type: ignore[arg-type]
                        logging.info("Baris %s: referensi %s ternyata terpasang", row_number, value)
                        attached_names.append(name)
                except Exception:
                    pass

    if dry_run:
        logging.info("[SIMULASI] Baris %s: klik Generate", row_number)
        return []
    else:
        # Isi teks paling akhir setelah seluruh dialog karakter selesai. Kemudian
        # tunggu teks + thumbnail tetap lengkap selama beberapa polling beruntun.
        fill_prompt(page, prompt_value, config)  # type: ignore[arg-type]
        expected_references = len(attached_names)
        wait_prompt_bundle_ready(
            page, config, prompt_value, expected_references, row_number  # type: ignore[arg-type]
        )
        attached_references = prompt_reference_count(page, config)  # type: ignore[arg-type]
        logging.info(
            "Baris %s: pemeriksaan akhir referensi sebelum Generate = %s thumbnail "
            "(%s karakter tersedia berhasil dipasang)",
            row_number, attached_references, len(attached_names),
        )
        pre_generate_dir = APP_DIR / "runtime" / "screenshots"
        pre_generate_dir.mkdir(parents=True, exist_ok=True)
        if config.get("debug_screenshots", False):
            safe_screenshot(page, pre_generate_dir / f"before-generate-row-{row_number}.png")
        try:
            # Utamakan tombol Generate yang berada dalam konteks prompt dan
            # sudah dipastikan aktif. Selector global lama tetap menjadi fallback.
            click_prompt_action(page, config, "right")  # type: ignore[arg-type]
        except RuntimeError:
            generate = first_visible(page, config["generate_selectors"], 2000)  # type: ignore[arg-type]
            if not generate.is_enabled(timeout=1000):
                raise RuntimeError("Tombol Generate ditemukan tetapi belum aktif")
            generate.click()
        if run_state is not None:
            run_state["generate_count"] = int(run_state.get("generate_count", 0)) + 1
            logging.info(
                "Generate ke-%s telah dikirim dengan model %s",
                run_state["generate_count"],
                run_state.get("active_model", "tidak diketahui"),
            )
        page.wait_for_timeout(int(config.get("after_generate_wait_ms", 1500)))
        success_dir = APP_DIR / "runtime" / "screenshots"
        success_dir.mkdir(parents=True, exist_ok=True)
        if config.get("debug_screenshots", False):
            safe_screenshot(page, success_dir / f"success-row-{row_number}.png")
        logging.info("Baris %s: Generate diklik", row_number)
        downloaded = download_new_outputs(page, before_ids, row_number, config, scene)  # type: ignore[arg-type]
        logging.info("Baris %s selesai: %s file diunduh", row_number, len(downloaded))
        return downloaded


def open_context(playwright: "Playwright", config: dict[str, Any]):
    profile_dir = resolve_path(config.get("profile_dir", "runtime/browser-profile"))
    profile_dir.mkdir(parents=True, exist_ok=True)
    return playwright.chromium.launch_persistent_context(
        user_data_dir=str(profile_dir),
        channel=config.get("browser_channel", "chrome"),
        headless=bool(config.get("headless", False)),
        slow_mo=int(config.get("slow_mo_ms", 0)),
        viewport={"width": 1440, "height": 960},
        args=[
            "--disable-features=Translate,TranslateUI",
            # Tetap merender walau jendela Chrome tertutup/diminimalkan.
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding",
            "--disable-background-timer-throttling",
        ],
    )


def configured_model_names(config: dict[str, Any]) -> list[str]:
    schedule = config.get("model_schedule", {})
    names = [
        str(schedule.get("primary_model", "Nano Banana 2")).strip(),
        str(schedule.get("fallback_model", "Nano Banana 2 Lite")).strip(),
        "Nano Banana Pro",
    ]
    return list(dict.fromkeys(name for name in names if name))


def model_name_in_text(text: str, config: dict[str, Any]) -> str | None:
    normalized = " ".join(str(text or "").casefold().split())
    for name in sorted(configured_model_names(config), key=len, reverse=True):
        if name.casefold() in normalized:
            return name
    return None


def visible_model_buttons(page: "Page", config: dict[str, Any]) -> list[tuple[float, float, Any, str]]:
    """Temukan kontrol model secara semantik, lalu urutkan menurut posisinya."""
    result: list[tuple[float, float, Any, str]] = []
    buttons = page.locator("button, [role='button']")
    for index in range(buttons.count()):
        button = buttons.nth(index)
        try:
            if not button.is_visible(timeout=200) or not button.is_enabled(timeout=200):
                continue
            identity = " ".join([
                button.inner_text(timeout=300) or "",
                button.get_attribute("aria-label", timeout=300) or "",
                button.get_attribute("title", timeout=300) or "",
                button.get_attribute("data-tooltip", timeout=300) or "",
            ])
            name = model_name_in_text(identity, config)
            box = button.bounding_box()
            if name and box is not None:
                result.append((float(box["y"]), float(box["x"]), button, name))
        except Exception:
            continue
    return sorted(result, key=lambda item: (item[0], item[1]))


def current_generation_model(page: "Page", config: dict[str, Any]) -> str | None:
    """Baca model aktif dari chip/dropdown yang terlihat di sekitar prompt."""
    candidates = visible_model_buttons(page, config)
    if not candidates:
        return None
    try:
        prompt = first_visible(page, config["prompt"]["selectors"], 1000)
        prompt_box = prompt.bounding_box()
    except Exception:
        prompt_box = None
    if prompt_box is None:
        return candidates[-1][3]
    prompt_center_x = prompt_box["x"] + prompt_box["width"] / 2
    prompt_center_y = prompt_box["y"] + prompt_box["height"] / 2
    ranked = []
    for y, x, button, name in candidates:
        box = button.bounding_box()
        if box is None:
            continue
        center_x = box["x"] + box["width"] / 2
        center_y = box["y"] + box["height"] / 2
        distance = (center_x - prompt_center_x) ** 2 + (center_y - prompt_center_y) ** 2
        ranked.append((distance, name))
    return min(ranked, key=lambda item: item[0])[1] if ranked else candidates[-1][3]


def visible_exact_model_option(page: "Page", model_name: str):
    pattern = re.compile(rf"^\s*(?:🍌\s*)?{re.escape(model_name)}\s*$", re.IGNORECASE)
    matches = page.get_by_text(pattern, exact=False)
    for index in range(matches.count()):
        candidate = matches.nth(index)
        try:
            if candidate.is_visible(timeout=200):
                return candidate
        except Exception:
            continue
    return None


def ensure_generation_model(page: "Page", config: dict[str, Any], model_name: str) -> None:
    """Pilih model lewat label UI dan verifikasi chip aktif berubah stabil."""
    schedule = config.get("model_schedule", {})
    timeout_ms = int(schedule.get("selection_timeout_ms", 30000))
    deadline = time.monotonic() + timeout_ms / 1000
    current = current_generation_model(page, config)
    if current and current.casefold() == model_name.casefold():
        logging.info("Model generate sudah sesuai: %s", model_name)
        return
    if not visible_model_buttons(page, config):
        if not _DIAG_DUMPED.get("no_model_picker"):
            expand_prompt(page, config)
        if not visible_model_buttons(page, config) and not visible_exact_model_option(page, model_name):
            if not _DIAG_DUMPED.get("no_model_picker"):
                _DIAG_DUMPED["no_model_picker"] = True
                logging.warning(
                    "CARA CADANGAN | pemilih model tidak ditemukan; generate memakai model yang aktif di Flow (bukan %s)",
                    model_name,
                )
            return

    last_seen = current or "tidak terdeteksi"
    while time.monotonic() < deadline:
        option = visible_exact_model_option(page, model_name)
        if option is not None:
            try:
                option.click(timeout=2500)
                page.wait_for_timeout(500)
            except Exception:
                page.wait_for_timeout(350)
        else:
            controls = visible_model_buttons(page, config)
            if not controls:
                expand_prompt(page, config)
                controls = visible_model_buttons(page, config)
            if controls:
                # Saat panel pengaturan sudah terbuka, dropdown model berada
                # di atas chip ringkasan prompt. Kandidat paling atas aman
                # untuk membuka daftar; saat panel tertutup hanya chip yang ada.
                try:
                    controls[0][2].click(timeout=2500)
                except Exception:
                    pass
            page.wait_for_timeout(500)

        stable = 0
        for _ in range(3):
            selected = current_generation_model(page, config)
            if selected:
                last_seen = selected
            if selected and selected.casefold() == model_name.casefold():
                stable += 1
            else:
                stable = 0
            page.wait_for_timeout(300)
        if stable >= 3:
            logging.info("Model generate dipilih dan terverifikasi: %s", model_name)
            close_settings_panel(page)
            return

    if last_seen == "tidak terdeteksi" and not visible_model_buttons(page, config):
        # Cara cadangan: tampilan Flow tanpa pemilih model -> pakai model yang aktif.
        logging.warning(
            "CARA CADANGAN | pemilih model tidak ditemukan; generate memakai model yang aktif di Flow (bukan %s)",
            model_name,
        )
        return
    raise RuntimeError(
        f"Model {model_name} tidak dapat dipilih setelah {timeout_ms} ms "
        f"(model terakhir terlihat: {last_seen})"
    )


def pace_row(page: "Page", config: dict[str, Any], overall_prompt: int) -> None:
    """Jeda acak antar-generate agar ritme tidak terlalu cepat untuk Flow."""
    import random

    if overall_prompt <= 1:
        return
    pace = config.get("delay_between_rows_ms", [15000, 30000])
    try:
        low, high = (int(pace[0]), int(pace[1])) if isinstance(pace, (list, tuple)) else (int(pace), int(pace))
    except Exception:
        low, high = 15000, 30000
    if high <= 0:
        return
    delay = random.randint(max(0, min(low, high)), max(low, high))
    logging.info("Jeda %s detik sebelum generate berikutnya", round(delay / 1000))
    page.wait_for_timeout(delay)


def close_settings_panel(page: "Page") -> None:
    """Tutup panel pengaturan model (UI Flow baru) agar tidak menghalangi tombol +."""
    for _ in range(2):
        try:
            family = page.locator(
                "button[aria-label='Select model family'], button[aria-label='Pilih rangkaian model']"
            )
            menus = page.locator("[role='menu']")
            if not any(family.nth(i).is_visible() for i in range(family.count())) and menus.count() == 0:
                return
            page.keyboard.press("Escape")
            page.wait_for_timeout(400)
        except Exception:
            return


def scheduled_generation_model(config: dict[str, Any], run_state: dict[str, Any]) -> str:
    schedule = config.get("model_schedule", {})
    primary = str(schedule.get("primary_model", "Nano Banana 2")).strip()
    fallback = str(schedule.get("fallback_model", "Nano Banana 2 Lite")).strip()
    primary_count = int(schedule.get("primary_generate_count", 0) or 0)
    if run_state.get("fallback_active"):
        return fallback
    if primary_count > 0 and int(run_state.get("generate_count", 0)) >= primary_count:
        return fallback
    return primary


MAIN_VIEW_NAV = re.compile(r"^\s*\S*\s*(All media|Semua media)\s*$", re.I)
OTHER_VIEW_NAV = re.compile(r"^\s*\S*\s*(Characters|Karakter|Scenes|Adegan)\s*$", re.I)


IMAGES_VIEW_NAV = re.compile(r"^\s*\S*\s*(Images|Gambar)\s*$", re.I)


def ensure_main_view(page: "Page") -> bool:
    """Flow mengingat menu terakhir (mis. Karakter setelah pindah karakter).
    Generate harus di 'All media/Semua media' (cadangan: 'Images/Gambar') agar
    gambar baru terlihat di grid."""
    try:
        active = page.locator("mat-list-item.mdc-list-item--activated, nav [aria-current='page']")
        texts = [active.nth(i).inner_text(timeout=500) or "" for i in range(min(active.count(), 5))]
        if not any(OTHER_VIEW_NAV.search(t) for t in texts):
            return False
        first_success("Pindah ke Semua media", [
            ("menu Semua media", lambda: _click_visible(page.locator("mat-list-item").filter(has_text=MAIN_VIEW_NAV))),
            ("tautan Semua media", lambda: _click_visible(
                page.locator("a, button, [role='tab'], [role='link'], li").filter(has_text=MAIN_VIEW_NAV))),
            ("menu Gambar", lambda: _click_visible(
                page.locator("mat-list-item, a, button, [role='tab'], li").filter(has_text=IMAGES_VIEW_NAV))),
        ])
        page.wait_for_timeout(1500)
        logging.info("Tampilan Flow dipindah dari menu Karakter ke Semua media/Gambar")
        return True
    except Exception as exc:
        if not _DIAG_DUMPED.get("main_view_warned"):
            _DIAG_DUMPED["main_view_warned"] = True
            logging.warning("Tampilan Semua media tidak bisa dibuka: %s", exc)
    return False


def expand_prompt(page: "Page", config: dict[str, Any]) -> None:
    """Prompt yang menyusut (hanya 'Apa yang ingin Anda buat?') perlu diklik agar
    pemilih model & tombol tampil."""
    try:
        first_success("Buka kotak prompt", [
            ("klik prompt", lambda: first_visible(page, config["prompt"]["selectors"], 1000).click(timeout=2000) or True),
            ("tombol Luaskan/Expand", lambda: _click_visible(
                page.locator("button[aria-label='Luaskan'], button[aria-label='Expand'], button[aria-label*='expand' i]"))),
        ])
        page.wait_for_timeout(700)
    except Exception:
        pass


_DIAG_DUMPED: dict[str, bool] = {}


def ensure_agent_mode_off(page: "Page") -> bool:
    """Akun Flow tertentu menyalakan mode 'Agen/Agent' di kotak prompt. Pada mode itu
    pemilih model (Nano Banana) tidak tampil dan prompt dikirim ke agen, bukan
    generate gambar biasa. Matikan mode agen bila aktif."""
    try:
        chips = page.locator(
            "button.agent-mode-chip[aria-pressed='true'], button.agent-mode-chip-checked"
        )
        if chips.count() == 0:
            chips = page.locator("button[aria-pressed='true']").filter(
                has_text=re.compile(r"^\s*(Agen|Agent)\s*$", re.I)
            )
        for index in range(chips.count()):
            chip = chips.nth(index)
            if chip.is_visible(timeout=300):
                chip.click(timeout=3000)
                page.wait_for_timeout(1200)
                logging.info("Mode Agen/Agent di kotak prompt dimatikan agar bisa generate gambar biasa")
                return True
    except Exception:
        pass
    return False


def wait_flow_ready(page: "Page", config: dict[str, Any], timeout_ms: int | None = None) -> None:
    """Tunggu loading hilang dan seluruh kontrol inti stabil/interaktif."""
    timeout_ms = int(timeout_ms or config.get("ready_timeout_ms", 120000))
    schedule = config.get("model_schedule", {})
    stable_required = max(1, int(schedule.get("ready_stable_checks", 3)))
    poll_ms = max(250, int(schedule.get("ready_poll_interval_ms", 1000)))
    started = time.monotonic()
    deadline = started + timeout_ms / 1000
    stable = 0
    reason = ""
    last_report = started
    dumped = _DIAG_DUMPED.get("flow_belum_siap", False)
    model_misses = 0
    model_missing_since: float | None = None
    model_fallback_logged = False
    while time.monotonic() < deadline:
        if project_missing(page):
            raise FlowProjectNotFound(
                f"Project Google Flow tidak ditemukan untuk akun profil ini: {config.get('url')}"
            )
        loading = page.get_by_text(
            re.compile(r"^Memuat(?:\.\.\.)?$|^Loading(?:\.\.\.)?$|^Please wait(?:\.\.\.)?$", re.IGNORECASE)
        )
        loading_visible = False
        for index in range(loading.count()):
            try:
                if loading.nth(index).is_visible():
                    loading_visible = True
                    break
            except Exception:
                loading_visible = True
                break
        if loading_visible:
            stable = 0
            reason = "halaman masih memuat"
        else:
            try:
                ready_state = page.evaluate("document.readyState")
                if ready_state not in {"interactive", "complete"}:
                    raise RuntimeError(f"document.readyState={ready_state}")
                if ensure_main_view(page):
                    raise RuntimeError("pindah ke tampilan Semua media")
                if ensure_agent_mode_off(page):
                    raise RuntimeError("mode Agen baru dimatikan")
                prompt = first_visible(page, config["prompt"]["selectors"], 1000)
                if not prompt.is_editable(timeout=1000) or not prompt.is_enabled(timeout=1000):
                    raise RuntimeError("kotak prompt belum interaktif")
                # Pastikan tombol + dan pemilih model sudah selesai dirender.
                prompt_action_button(page, config, "left")
                if schedule.get("enabled", False) and current_generation_model(page, config) is None:
                    model_misses += 1
                    if model_misses % 3 == 0:
                        expand_prompt(page, config)
                    # Cara cadangan: prompt & tombol + sudah siap tetapi pemilih model
                    # tidak tampil (tampilan Flow berbeda) -> lanjut dengan model aktif.
                    model_missing_since = model_missing_since or time.monotonic()
                    tolerance_s = min(
                        float(config.get("model_missing_tolerance_s", 40)), max(10.0, timeout_ms / 1000 / 3)
                    )
                    if time.monotonic() - model_missing_since < tolerance_s:
                        raise RuntimeError("pemilih model (Nano Banana) belum terlihat")
                    if not model_fallback_logged:
                        model_fallback_logged = True
                        logging.warning(
                            "CARA CADANGAN | pemilih model tidak terlihat; generate memakai model yang aktif di Flow"
                        )
                else:
                    model_missing_since = None
                stable += 1
                if stable >= stable_required:
                    logging.info(
                        "Google Flow siap dan stabil (%s/%s pemeriksaan)",
                        stable, stable_required,
                    )
                    return
            except Exception as exc:
                stable = 0
                reason = str(exc).splitlines()[0][:160]
        now = time.monotonic()
        if now - last_report >= 15:
            last_report = now
            logging.info("Menunggu Flow siap (%ss): %s", int(now - started), reason or "-")
        if not dumped and now - started >= 30:
            dumped = True
            _DIAG_DUMPED["flow_belum_siap"] = True  # cukup sekali per proses
            dump_page(page, "flow_belum_siap")
        page.wait_for_timeout(poll_ms)
    raise RuntimeError(f"Google Flow belum siap setelah menunggu {timeout_ms} ms ({reason})")


def refresh_for_next(
    page: "Page",
    config: dict[str, Any],
    *,
    timeout_ms: int | None = None,
    attempts: int = 2,
    reason: str = "pemulihan",
) -> None:
    """Segarkan Flow dan pastikan kotak prompt siap sebelum baris berikutnya."""
    last_error = None
    for attempt in range(max(1, attempts)):
        try:
            if attempt == 0:
                page.reload(wait_until="domcontentloaded")
            else:
                page.goto(config["url"], wait_until="domcontentloaded")
            wait_flow_ready(page, config, timeout_ms=timeout_ms)
            logging.info("Halaman direfresh, pulih, dan siap (%s)", reason)
            return
        except FlowProjectNotFound:
            raise
        except Exception as exc:
            last_error = exc
            logging.warning(
                "Refresh %s percobaan %s/%s belum siap: %s",
                reason, attempt + 1, max(1, attempts), exc,
            )
    raise RuntimeError(
        f"Halaman tidak siap setelah {max(1, attempts)} kali refresh ({reason})"
    ) from last_error


# ---------------------------------------------------------------------------
# Pindah karakter antar profil (akun Google) Flow
# ---------------------------------------------------------------------------

CHARACTER_ROOT = "downloads/_KARAKTER"
# Flow tampil dalam bahasa akun (Inggris / Indonesia).
SEARCH_INPUT = "input[aria-label='Search'], input[aria-label='Telusuri']"
NAME_INPUT = "input[aria-label='Character name'], input[aria-label='Nama karakter']"
PERSONALITY_INPUT = "textarea[aria-label='Character personality'], textarea[aria-label='Kepribadian karakter']"


def safe_file_name(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", str(value)).strip(" .") or "karakter"


def character_folder(profile_name: str, config: dict[str, Any]) -> Path:
    root = resolve_path(str(config.get("character_folder", CHARACTER_ROOT)))
    return root / safe_file_name(profile_name)


_METHOD_USED: dict[str, int] = {}


def first_success(label: str, methods: list[tuple[str, Any]]) -> Any:
    """Coba cara 1, bila gagal cara 2, dst. Tampilan Flow sering berubah; cara lama
    tetap disimpan sebagai cadangan agar bot tetap jalan di tampilan apa pun.
    Setiap fungsi cara harus mengembalikan nilai 'truthy' bila berhasil."""
    errors: list[str] = []
    for number, (method_name, func) in enumerate(methods, start=1):
        try:
            result = func()
            if not result:
                raise RuntimeError("tidak berhasil")
            if _METHOD_USED.get(label) != number:
                _METHOD_USED[label] = number
                if number > 1:
                    logging.info("CARA CADANGAN | %s memakai cara %s (%s)", label, number, method_name)
            return result
        except Exception as exc:
            errors.append(f"cara {number} ({method_name}): {str(exc).splitlines()[0][:140] if str(exc) else type(exc).__name__}")
    raise RuntimeError(f"{label} gagal dengan semua cara: " + " ; ".join(errors))


CHARACTER_NAV = re.compile(r"^\s*\S*\s*(Characters|Karakter)\s*$", re.I)


def _click_visible(locator) -> bool:
    for index in range(min(locator.count(), 20)):
        item = locator.nth(index)
        try:
            if item.is_visible(timeout=300):
                item.click(timeout=3000)
                return True
        except Exception:
            continue
    return False


def open_characters_view(page: "Page", config: dict[str, Any]) -> str:
    """Buka project lalu menu Characters/Karakter.

    Cara 1: item sidebar (mat-list-item). Cara 2: tautan/tombol/tab lain bertuliskan
    Characters/Karakter. Bila tidak ada menu karakter sama sekali (tampilan lama),
    kembalikan "panel" agar daftar karakter dibaca lewat panel aset (+ > Karakter).
    Tidak memakai wait_flow_ready: project kosong kadang belum menampilkan pemilih model."""
    page.goto(config["url"], wait_until="domcontentloaded")
    timeout_ms = int(config.get("character_nav_timeout_ms", 60000))
    deadline = time.monotonic() + timeout_ms / 1000
    methods = [
        ("menu samping", lambda: _click_visible(page.locator("mat-list-item").filter(has_text=CHARACTER_NAV))),
        ("tautan/tab lain", lambda: _click_visible(
            page.locator("a, button, [role='tab'], [role='link'], [role='menuitem'], li").filter(has_text=CHARACTER_NAV)
        )),
    ]
    clicked = False
    while time.monotonic() < deadline and not clicked:
        if project_missing(page):
            raise FlowProjectNotFound(
                f"Project Google Flow tidak ditemukan untuk akun profil ini: {config.get('url')}"
            )
        try:
            clicked = bool(first_success("Buka menu Karakter", methods))
        except Exception:
            page.wait_for_timeout(800)
    if not clicked:
        logging.warning("Menu Karakter/Characters tidak ada; memakai panel aset (+ > Karakter)")
        return "panel"
    # Tunggu grid karakter tampil: kartu karakter atau kartu "Karakter baru".
    # Project kosong bisa menampilkan tampilan lain; jangan gagal, cukup lanjut.
    try:
        page.wait_for_function(
            """() => document.querySelector('flow-character-tile, flow-custom-tile')
                || [...document.querySelectorAll('button')].some(b => /New character|Karakter baru|Create character|Buat karakter/i.test(b.innerText || ''))""",
            timeout=20000,
        )
    except Exception:
        logging.warning("Grid karakter belum terlihat (project mungkin masih kosong); lanjut")
        dump_page(page, "karakter_kosong")
    page.wait_for_timeout(1200)
    logging.info("Menu karakter terbuka")
    return "grid"


def project_base_url(page: "Page", config: dict[str, Any]) -> str:
    for value in (page.url, str(config.get("url", ""))):
        match = re.match(r"^(https?://[^?#]*?/project/[^/?#]+)", value)
        if match:
            return match.group(1)
    raise RuntimeError("URL project Flow tidak dikenali")


def dump_page(page: "Page", label: str) -> None:
    folder = APP_DIR / "runtime" / "logs" / "diag_karakter"
    try:
        folder.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%H%M%S")
        (folder / f"{label}_{stamp}.html").write_text(page.content(), encoding="utf-8")
        safe_screenshot(page, folder / f"{label}_{stamp}.png")
        logging.info("DIAG | %s_%s disimpan", label, stamp)
    except Exception as exc:
        logging.warning("DIAG | gagal simpan %s: %s", label, exc)


CHAR_SCROLLER_JS = """() => document.querySelector('.cdk-virtual-scrollable')
    || document.querySelector('cdk-virtual-scroll-viewport') || document.scrollingElement"""


def list_characters_grid(page: "Page") -> list[dict[str, str]]:
    """Cara 1: kartu karakter di grid menu Karakter (digulir sampai habis)."""
    found: dict[str, dict[str, str]] = {}
    idle = 0
    for _ in range(400):
        batch = page.evaluate(
            """() => [...document.querySelectorAll('flow-grid-tile-container')].map(tile => {
                const isChar = tile.querySelector('flow-character-tile')
                    || tile.querySelector('img[alt="Character thumbnail"], img[alt*="karakter" i], img[alt*="character" i]');
                const img = tile.querySelector('img');
                return isChar && img ? {name: tile.getAttribute('aria-label') || '', src: img.src || ''} : null;
            }).filter(Boolean)"""
        )
        before = len(found)
        for item in batch:
            name = str(item.get("name") or "").strip()
            if name and not UNTITLED_NAME.match(name) and name not in found:
                found[name] = {"name": name, "src": item.get("src") or ""}
        state = page.evaluate(
            "(() => { const el = (" + CHAR_SCROLLER_JS + ")();"
            " const top = el.scrollTop; el.scrollTop = top + Math.max(300, el.clientHeight * 0.8);"
            " return {moved: el.scrollTop !== top, height: el.scrollHeight}; })()"
        )
        if not state.get("moved"):
            # Cadangan: gulir dengan roda mouse di tengah halaman.
            try:
                page.mouse.move(800, 500)
                page.mouse.wheel(0, 700)
            except Exception:
                pass
        page.wait_for_timeout(800)
        if len(found) > before or state.get("moved"):
            idle = 0
        else:
            idle += 1
            # Di dasar grid: beri waktu Flow memuat halaman berikutnya.
            page.wait_for_timeout(1200)
            if idle >= 4:
                break
    try:
        page.evaluate("(() => { const el = (" + CHAR_SCROLLER_JS + ")(); el.scrollTop = 0; })()")
    except Exception:
        pass
    return list(found.values())


ASSET_ITEMS_JS = """() => [...document.querySelectorAll('[role=option], .asset-item')].map(e => {
    const title = e.querySelector('.asset-title, [class*=title]');
    const img = e.querySelector('img');
    const name = ((title && title.textContent) || e.innerText || '').trim().split('\\n')[0];
    return name && img ? {name, src: img.src || ''} : null;
}).filter(Boolean)"""


def list_characters_panel(page: "Page", config: dict[str, Any]) -> list[dict[str, str]]:
    """Cara 2: panel aset (+ > Karakter) seperti tampilan Flow lama."""
    open_asset_picker(page, config)
    click_character_category(page, 15000)
    page.wait_for_timeout(1500)
    found: dict[str, dict[str, str]] = {}
    idle = 0
    for _ in range(300):
        before = len(found)
        for item in page.evaluate(ASSET_ITEMS_JS):
            name = str(item.get("name") or "").strip()
            if name and not UNTITLED_NAME.match(name) and name not in found:
                found[name] = {"name": name, "src": item.get("src") or "", "panel": "1"}
        moved = page.evaluate(
            """() => { const el = document.querySelector('.asset-list-viewport, [role=listbox]');
                if (!el) return false; const top = el.scrollTop; el.scrollTop = top + 300; return el.scrollTop !== top; }"""
        )
        page.wait_for_timeout(600)
        idle = 0 if (moved or len(found) > before) else idle + 1
        if idle >= 3:
            break
    return list(found.values())


def close_asset_panel(page: "Page") -> None:
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
    except Exception:
        pass


def list_characters(page: "Page", config: dict[str, Any] | None = None, mode: str = "grid") -> list[dict[str, str]]:
    """Daftar karakter: cara 1 grid menu Karakter, cara 2 panel aset."""
    result: list[dict[str, str]] = []
    if mode == "grid":
        result = list_characters_grid(page)
    grid_present = page.evaluate(
        "() => !!document.querySelector('flow-character-tile, flow-custom-tile')"
    ) if mode == "grid" else False
    if not result and not grid_present and config is not None:
        try:
            result = list_characters_panel(page, config)
            if result:
                logging.info("CARA CADANGAN | daftar karakter dibaca lewat panel aset")
        except Exception as exc:
            logging.warning("Panel aset karakter tidak terbaca: %s", exc)
        finally:
            close_asset_panel(page)
    logging.info("KARAKTER | %s kartu karakter terbaca", len(result))
    return result


def dump_character_diagnostics(page: "Page", config: dict[str, Any], label: str) -> None:
    """Simpan HTML + screenshot halaman karakter untuk diperiksa bila tidak ada karakter terbaca."""
    folder = APP_DIR / "runtime" / "logs" / "diag_karakter"
    folder.mkdir(parents=True, exist_ok=True)
    tag = safe_file_name(label)

    def save(step: str) -> None:
        try:
            (folder / f"{tag}_{step}.html").write_text(page.content(), encoding="utf-8")
            safe_screenshot(page, folder / f"{tag}_{step}.png")
            logging.info("DIAG | %s_%s disimpan (url %s)", tag, step, page.url)
        except Exception as exc:
            logging.warning("DIAG | %s gagal: %s", step, exc)

    save("1_halaman")
    try:
        open_asset_picker(page, config)
        page.wait_for_timeout(1500)
        save("2_panel_aset")
        click_character_category(page, 10000)
        page.wait_for_timeout(2500)
        save("3_tab_karakter")
        page.keyboard.press("Escape")
    except Exception as exc:
        logging.warning("DIAG | panel aset: %s", exc)
        save("2_gagal")


def tile_selector(name: str) -> str:
    """Selector kartu karakter berdasarkan nama (aman untuk huruf non-ASCII & tanda kutip)."""
    escaped = str(name).replace("\\", "\\\\").replace('"', '\\"')
    return f'flow-grid-tile-container[aria-label="{escaped}"]'


def read_character_info(page: "Page", config: dict[str, Any], name: str) -> str:
    """Buka halaman karakter (klik dua kali kartunya) dan baca deskripsinya."""
    try:
        search = page.locator(SEARCH_INPUT).first
        search.fill(name, timeout=5000)
        page.wait_for_timeout(1500)
        tile = page.locator(tile_selector(name)).first
        tile.dblclick(timeout=8000)
        page.wait_for_url(re.compile(r"/character/[^/?#]+"), timeout=15000)
        page.wait_for_timeout(1500)
        text = page.locator(PERSONALITY_INPUT).first.input_value(timeout=5000)
        page.go_back(wait_until="domcontentloaded")
        page.wait_for_timeout(1500)
        return text or ""
    except Exception as exc:
        logging.warning("Deskripsi karakter %s tidak terbaca: %s", name, exc)
        try:
            open_characters_view(page, config)
        except Exception:
            pass
        return ""
    finally:
        try:
            page.locator(SEARCH_INPUT).first.fill("", timeout=3000)
        except Exception:
            pass


def _full_size_url(src: str) -> str:
    """Thumbnail lh3 (panel aset) -> ukuran asli (=s0)."""
    if "googleusercontent.com" in src:
        return re.sub(r"=[swh]\d+[^/]*$", "", src) + "=s0"
    return src


def download_character_image(page: "Page", item: dict[str, str]) -> tuple[bytes, str]:
    """Unduh gambar karakter. Cara 1 fetch di halaman, cara 2 request browser,
    cara 3 ukuran asli dari thumbnail, cara 4 screenshot kartu."""
    src = item.get("src") or ""

    def by_page(url: str):
        payload = page.evaluate(READ_BLOB_JS, url)
        data = base64.b64decode(payload["b64"])
        kind = str(payload.get("type") or "").casefold()
        if kind and not kind.startswith("image/"):
            raise RuntimeError(f"bukan gambar ({kind})")
        if len(data) < 500:
            raise RuntimeError("gambar terlalu kecil")
        return data, kind

    def by_request(url: str):
        response = page.context.request.get(url, timeout=30000)
        if not response.ok:
            raise RuntimeError(f"HTTP {response.status}")
        data = response.body()
        kind = str(response.headers.get("content-type", "")).casefold()
        if kind and not kind.startswith("image/"):
            raise RuntimeError(f"bukan gambar ({kind})")
        if len(data) < 500:
            raise RuntimeError("gambar terlalu kecil")
        return data, kind

    def by_screenshot():
        tile = page.locator(tile_selector(item["name"]) + " img")
        if tile.count() == 0:
            tile = page.locator("[role=option], .asset-item").filter(has_text=item["name"]).locator("img")
        return tile.first.screenshot(timeout=5000), "image/png"

    methods: list[tuple[str, Any]] = []
    if item.get("panel"):
        methods.append(("ukuran asli dari thumbnail", lambda: by_page(_full_size_url(src))))
    methods += [
        ("fetch di halaman", lambda: by_page(src)),
        ("request browser", lambda: by_request(src)),
        ("screenshot kartu", by_screenshot),
    ]
    return first_success("Unduh gambar karakter", methods)


def export_characters(
    page: "Page", config: dict[str, Any], profile_name: str, only: set[str] | None = None
) -> int:
    mode = open_characters_view(page, config)
    characters = list_characters(page, config, mode)
    if only is not None:
        characters = [c for c in characters if c["name"].casefold() in only]
    folder = character_folder(profile_name, config)
    folder.mkdir(parents=True, exist_ok=True)
    logging.info("KARAKTER | %s karakter ditemukan di profil %s", len(characters), profile_name)
    manifest = []
    for index, item in enumerate(characters, start=1):
        name = item["name"]
        try:
            data, content_type = download_character_image(page, item)
            suffix = ".png" if "png" in content_type else ".webp" if "webp" in content_type else ".jpg"
            file_name = safe_file_name(name) + suffix
            (folder / file_name).write_bytes(data)
        except Exception as exc:
            logging.error("KARAKTER | %s gagal diunduh: %s", name, exc)
            continue
        info = (
            read_character_info(page, config, name)
            if config.get("export_character_info", True) and mode == "grid" and not item.get("panel") else ""
        )
        manifest.append({"name": name, "file": file_name, "personality": info})
        logging.info("KARAKTER %s/%s | %s disimpan (%s KB)", index, len(characters), file_name, len(data) // 1024)
    manifest_path = folder / "karakter.json"
    if only is not None and manifest_path.is_file():
        # Gabungkan dengan hasil export sebelumnya.
        try:
            old = json.loads(manifest_path.read_text(encoding="utf-8")).get("characters", [])
            names = {m["name"].casefold() for m in manifest}
            manifest = [c for c in old if c.get("name", "").casefold() not in names] + manifest
        except Exception:
            pass
    manifest_path.write_text(
        json.dumps({"profile": profile_name, "characters": manifest}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    logging.info("KARAKTER SELESAI | %s karakter disimpan di %s", len(manifest), folder)
    return 0


def load_character_folder(folder: Path) -> list[dict[str, str]]:
    manifest = folder / "karakter.json"
    items: list[dict[str, str]] = []
    if manifest.is_file():
        data = json.loads(manifest.read_text(encoding="utf-8"))
        items = [c for c in data.get("characters", []) if (folder / c.get("file", "")).is_file()]
    known = {c["file"] for c in items}
    for path in sorted(folder.iterdir()) if folder.is_dir() else []:
        if path.suffix.casefold() in {".png", ".jpg", ".jpeg", ".webp"} and path.name not in known:
            items.append({"name": path.stem, "file": path.name, "personality": ""})
    return items


CHARACTER_STATE_JS = """() => {
    const onChar = location.pathname.includes('/character/');
    const img = [...document.querySelectorAll('img')].find(i => /character image|gambar karakter/i.test(i.alt || ''))
        || (onChar ? [...document.querySelectorAll('img')].find(i => i.naturalWidth > 300
            && i.getBoundingClientRect().width > 250 && !i.closest('flow-grid-tile-container, [role=dialog], mat-dialog-container')) : null);
    const leafs = [...document.querySelectorAll('*')].filter(e => e.children.length === 0);
    const busy = leafs.some(e => /^\\d{1,3}%$/.test((e.innerText || '').trim()));
    const uploadBtn = [...document.querySelectorAll('button')].some(b => /^\\s*\\S*\\s*(Upload|Unggah)\\s*$/i.test(b.innerText || ''));
    const toast = [...document.querySelectorAll('mat-snack-bar-container, [role=alert], [role=status], .mdc-snackbar__label')]
        .map(e => (e.innerText || '').trim()).filter(Boolean).join(' | ');
    return {ready: !!(img && img.naturalWidth > 100) && !busy, busy, empty: uploadBtn && !busy && !img, toast};
}"""


CHARACTER_URL = re.compile(r"/character/[^/?#]+")


def _norm_url(url: str | None) -> str:
    return re.sub(r"[?#].*$", "", url or "").rstrip("/")


# Judul karakter di atas halaman. Urutan = prioritas: selektor paling spesifik dulu
# (placeholder/aria), kelas umum paling akhir (bisa saja dipakai judul project).
HEADER_TITLE_SELECTORS = (
    "input[placeholder='Karakter tanpa judul']",
    "input[placeholder='Untitled character']",
    "input[aria-label='Teks yang dapat diedit']",
    "input[aria-label='Editable text']",
    "input.editable-text-input",
)
HEADER_TITLE_INPUT = ", ".join(HEADER_TITLE_SELECTORS)
UNTITLED_NAME = re.compile(r"^\s*(Karakter tanpa judul|Untitled character|Karakter tanpa nama)?\s*$", re.I)


CHARACTER_TITLE_PLACEHOLDER = re.compile(r"^\s*(Karakter tanpa judul|Untitled character|Karakter tanpa nama|Nama karakter|Character name)\s*$", re.I)


def header_title_box(page: "Page", writable: bool = False):
    """Kotak judul karakter di atas halaman. writable=True: hanya kotak yang placeholder-nya
    memang judul karakter (jangan sampai mengetik ke judul project)."""
    for selector in HEADER_TITLE_SELECTORS:
        try:
            locator = page.locator(selector)
            for index in range(min(locator.count(), 5)):
                box = locator.nth(index)
                if not writable:
                    return box
                if CHARACTER_TITLE_PLACEHOLDER.search(box.get_attribute("placeholder", timeout=2000) or ""):
                    return box
        except Exception:
            continue
    return None


def read_names_stable(page: "Page", tries: int = 4) -> tuple[str | None, str | None] | None:
    """Baca nama 2x berturut-turut (jeda 1,2 dtk) sampai hasilnya sama; judul wajib terbaca.
    None = tidak pasti (jangan dipakai untuk menghapus)."""
    previous = None
    for _ in range(tries):
        current = read_character_names(page)
        if current[0] is not None and current == previous:
            return current
        previous = current
        page.wait_for_timeout(1200)
    return None


def read_character_names(page: "Page") -> tuple[str | None, str | None]:
    """(judul di atas, kolom Nama karakter). None = tidak ada / tidak terbaca."""
    header: str | None = None
    typed: str | None = None
    box = header_title_box(page)
    if box is not None:
        try:
            header = box.input_value(timeout=2000).strip()
        except Exception:
            header = None
    try:
        names = page.locator(NAME_INPUT)
        if names.count():
            typed = names.first.input_value(timeout=2000).strip()
    except Exception:
        typed = None
    return header, typed


def wait_character_loaded(page: "Page", timeout_s: float = 25, need_image: bool = False) -> str | None:
    """Tunggu halaman satu karakter selesai dimuat.
    'ready' = gambar tampil, 'empty' = tanpa gambar (draf gagal upload), None = belum termuat.
    need_image=True: jangan berhenti di 'empty', tunggu gambar sampai waktu habis."""
    deadline = time.monotonic() + timeout_s
    empty_count = 0
    while time.monotonic() < deadline:
        try:
            state = page.evaluate(CHARACTER_STATE_JS)
        except Exception:
            state = {}
        if state.get("ready"):
            page.wait_for_timeout(600)
            return "ready"
        empty_count = empty_count + 1 if state.get("empty") else 0
        if empty_count >= 6 and not need_image:  # tetap kosong beberapa detik = memang tanpa gambar
            return "empty"
        page.wait_for_timeout(700)
    return None


def open_character_url(page: "Page", url: str, fresh: bool = True) -> bool:
    """Buka halaman karakter tertentu dan pastikan alamatnya benar (bukan halaman lain)."""
    target = _norm_url(url)
    if not target:
        return False
    if not fresh and _norm_url(page.url) == target:
        return True
    for wait_until in ("domcontentloaded", "load"):
        try:
            page.goto(target, wait_until=wait_until, timeout=45000)
        except Exception as exc:
            logging.warning("KARAKTER | gagal membuka %s: %s", target, str(exc).splitlines()[0][:150])
            continue
        page.wait_for_timeout(1500)
        if _norm_url(page.url) == target:
            return True
    return False


def delete_open_character(
    page: "Page", expected_url: str | None = None, own_name: str = "", strict: bool = False
) -> bool:
    """Hapus satu karakter TANPA NAMA. Kembalikan True hanya bila benar-benar terhapus.

    Pengaman:
    - hanya di halaman /character/<id>; bila expected_url diberikan, halaman harus karakter itu;
    - bila judul/kolom nama berisi nama lain (bukan kosong / "Karakter tanpa judul" / nama
      yang sedang dibuat bot), karakter TIDAK dihapus;
    - strict=True (pembersihan): nama wajib terbaca dan kosong."""
    if expected_url:
        if _norm_url(page.url) != _norm_url(expected_url) and not open_character_url(page, expected_url):
            logging.warning("KARAKTER | draf %s tidak bisa dibuka lagi; tidak ada yang dihapus", expected_url)
            return False
    if not CHARACTER_URL.search(page.url or ""):
        return False  # bukan halaman satu karakter -> jangan menghapus apa pun
    target = _norm_url(page.url)
    loaded = wait_character_loaded(page, 20)
    if loaded != "ready":
        # Gambar belum tampil: muat ulang & tunggu lebih lama, supaya karakter bernama yang
        # lambat termuat tidak sampai terbaca kosong.
        try:
            page.reload(wait_until="domcontentloaded", timeout=30000)
        except Exception:
            pass
        loaded = wait_character_loaded(page, 25)
        page.wait_for_timeout(5000 if loaded != "ready" else 1000)
        if _norm_url(page.url) != target:
            logging.warning("KARAKTER | %s berpindah halaman saat dicek; TIDAK dihapus", target)
            return False
    names = read_names_stable(page)
    if names is None:
        logging.warning("KARAKTER | nama %s tidak terbaca pasti; demi aman TIDAK dihapus", target)
        return False
    header, typed = names
    # Judul (nama tersimpan) WAJIB kosong. Kolom nama boleh berisi nama yang sedang diketik
    # bot (belum tersimpan) hanya saat menghapus draf bot sendiri (bukan pembersihan).
    typed_ok = typed is None or UNTITLED_NAME.match(typed) or (not strict and own_name and typed == own_name)
    if not UNTITLED_NAME.match(header or "") or not typed_ok:
        logging.warning("KARAKTER | %s punya nama '%s'; TIDAK dihapus", target, header or typed)
        return False

    delete_selector = (
        "button[aria-label='Delete'], button[aria-label='Hapus'], "
        "button[aria-label='Delete character'], button[aria-label='Hapus karakter']"
    )

    def click_delete():
        return _click_visible(page.locator(delete_selector))

    def click_delete_icon():
        return _click_visible(page.locator("button").filter(has_text=re.compile(r"^\s*delete\s*$", re.I)))

    def click_delete_text():
        return _click_visible(page.locator("button").filter(has_text=re.compile(r"^\s*\S*\s*(Delete|Hapus)\s*$", re.I)))

    def via_menu():
        if not _click_visible(page.locator(
            "button[aria-label='More options'], button[aria-label='Opsi lainnya'], button[aria-label*='More' i], "
            "button[aria-label*='opsi' i]"
        )):
            raise RuntimeError("menu tidak ada")
        page.wait_for_timeout(600)
        return _click_visible(page.locator("[role=menuitem], button").filter(has_text=re.compile(r"(Delete|Hapus)", re.I)))

    project_base = target.split("/character/")[0]

    def left_page() -> bool:
        current = _norm_url(page.url)
        return current != target and current.startswith(project_base) and not CHARACTER_URL.search(current)

    def confirm() -> None:
        page.wait_for_timeout(800)
        if left_page():
            return  # langsung terhapus tanpa jendela konfirmasi
        dialog_buttons = page.locator(
            "mat-dialog-container button, [role=dialog] button, [role=alertdialog] button, .cdk-overlay-pane button"
        )
        if _click_visible(dialog_buttons.filter(has_text=re.compile(r"^\s*\S*\s*(Delete|Hapus|Ya|Yes)\s*$", re.I))):
            return
        if _click_visible(dialog_buttons.filter(has_text=re.compile(r"(Delete|Hapus)", re.I))):
            return
        try:
            if page.locator("mat-dialog-container, [role=dialog], [role=alertdialog]").count():
                page.keyboard.press("Enter")
        except Exception:
            pass

    try:
        first_success("Tombol hapus karakter", [
            ("tombol Hapus/Delete", click_delete), ("ikon delete", click_delete_icon),
            ("tombol bertuliskan Hapus", click_delete_text), ("menu opsi", via_menu),
        ])
        confirm()
    except Exception as exc:
        logging.warning("KARAKTER | karakter kosong gagal dihapus: %s", exc)
        return False
    # Pastikan benar-benar terhapus: cara 1 halaman berpindah; cara 2 alamatnya dibuka lagi.
    for _ in range(20):
        if left_page():
            break
        page.wait_for_timeout(500)
    gone = left_page()
    if not gone:
        # Cara 2: buka lagi alamatnya. Terhapus bila diarahkan ke project / tertulis tidak ditemukan.
        try:
            page.goto(target, wait_until="domcontentloaded", timeout=30000)
            loaded = wait_character_loaded(page, 12)
            if left_page():
                gone = True
            elif loaded is None and header_title_box(page) is None:
                text = page.locator("body").inner_text(timeout=3000)
                gone = bool(re.search(r"tidak ditemukan|not found|doesn.t exist|does not exist|tidak tersedia|telah dihapus|was deleted|has been deleted|tidak ada karakter|karakter tidak ada", text, re.I))
        except Exception:
            gone = False
    if gone:
        logging.info("KARAKTER | draf tanpa nama dihapus (%s)", target.rsplit("/", 1)[-1])
    else:
        logging.warning("KARAKTER | %s belum terhapus", target)
    return gone


CONSENT_TEXT = re.compile(r"Hak untuk menggunakan gambar ini|Rights to use this image|right to use", re.I)
CONSENT_AGREE = re.compile(r"^\s*(\S+\s+)?(Saya setuju|Setuju|I agree|Agree|Accept|Terima)\s*$", re.I)


def wait_upload_consent(page: "Page", timeout_s: int = 300) -> None:
    """Akun baru: Flow meminta persetujuan 'Hak untuk menggunakan gambar ini'.
    Atas izin pengguna, bot mengklik 'Saya setuju' (cara 1). Bila tombolnya tidak
    ditemukan, bot menunggu pengguna mengklik sendiri (cara 2)."""
    dialog = page.locator(
        "mat-dialog-container, [role=dialog], [role=alertdialog], .cdk-overlay-pane"
    ).filter(has_text=CONSENT_TEXT)
    try:
        if dialog.count() == 0 or not dialog.first.is_visible(timeout=300):
            return
    except Exception:
        return
    agree = dialog.first.locator("button").filter(has_text=CONSENT_AGREE)
    if _click_visible(agree):
        page.wait_for_timeout(1500)
        logging.info("PERSETUJUAN FLOW | 'Saya setuju' diklik otomatis (hak penggunaan gambar)")
        return
    logging.warning(
        "PERSETUJUAN FLOW | Jendela 'Hak untuk menggunakan gambar ini' muncul tetapi tombol setuju tidak "
        "ditemukan. Klik 'Saya setuju' di jendela Chrome bot. Bot menunggu %s detik.", timeout_s
    )
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        page.wait_for_timeout(2000)
        try:
            if dialog.count() == 0 or not dialog.first.is_visible(timeout=300):
                logging.info("PERSETUJUAN FLOW | sudah disetujui, lanjut")
                return
        except Exception:
            return
    raise RuntimeError("Persetujuan upload gambar Flow belum diklik")


UPLOAD_BUTTON = re.compile(r"^\s*\S*\s*(Upload|Unggah)\s*$", re.I)
NEW_CHARACTER_BUTTON = re.compile(r"New character|Karakter baru|Create character|Buat karakter", re.I)


def open_new_character_page(page: "Page", config: dict[str, Any]) -> bool:
    def upload_ready() -> bool:
        page.wait_for_function(
            """() => [...document.querySelectorAll('button')].some(b => /^\\s*\\S*\\s*(Upload|Unggah)\\s*$/i.test(b.innerText || ''))
                || document.querySelector('input[type=file]')""",
            timeout=20000,
        )
        page.wait_for_timeout(800)
        if CHARACTER_URL.search(page.url or ""):
            # Halaman karakter yang sudah ada? Jangan upload ke sana (gambarnya bisa terganti).
            header, _typed = read_character_names(page)
            try:
                has_image = bool(page.evaluate(CHARACTER_STATE_JS).get("ready"))
            except Exception:
                has_image = True
            if (header is not None and not UNTITLED_NAME.match(header)) or has_image:
                raise RuntimeError("yang terbuka halaman karakter lain, bukan karakter baru")
        return True

    def by_url() -> bool:
        page.goto(project_base_url(page, config) + "/character", wait_until="domcontentloaded")
        return upload_ready()

    def by_button() -> bool:
        open_characters_view(page, config)
        if not _click_visible(page.locator("button, flow-custom-tile").filter(has_text=NEW_CHARACTER_BUTTON)):
            raise RuntimeError("tombol Karakter baru tidak ada")
        return upload_ready()

    return first_success("Buka halaman karakter baru", [("alamat /character", by_url), ("tombol Karakter baru", by_button)])


def upload_character_image(page: "Page", image: Path) -> bool:
    def by_chooser() -> bool:
        upload = page.locator("button").filter(has_text=UPLOAD_BUTTON)
        if upload.count() == 0:
            raise RuntimeError("tombol Upload tidak ada")
        with page.expect_file_chooser(timeout=15000) as chooser_info:
            upload.first.click(timeout=5000)
        chooser_info.value.set_files(str(image))
        return True

    def by_input() -> bool:
        inputs = page.locator("input[type=file]")
        if inputs.count() == 0:
            raise RuntimeError("input file tidak ada")
        inputs.first.set_input_files(str(image))
        return True

    return first_success("Upload gambar karakter", [("tombol Upload", by_chooser), ("input file langsung", by_input)])


def saved_name_matches(page: "Page", name: str) -> bool:
    """Nama dianggap tersimpan bila judul di atas halaman karakter ikut berubah.
    Bila judul tidak ada (tampilan lain), pakai isi kolom nama."""
    header, typed = read_character_names(page)
    if header is not None:
        return header == name
    return typed == name


def set_character_name(page: "Page", name: str) -> bool:
    """Isi nama karakter. Banyak cara cadangan; tiap cara diverifikasi dari judul halaman."""

    def verify() -> bool:
        # Flow menyimpan nama secara async; judul di atas harus ikut berubah dan kolom
        # nama tidak boleh berisi teks lain.
        for _ in range(10):
            header, typed = read_character_names(page)
            main = header if header is not None else typed
            if main == name and typed in (None, "", name):
                return True
            page.wait_for_timeout(500)
        raise RuntimeError("nama belum tersimpan (judul tidak berubah)")

    def blur() -> None:
        try:
            page.keyboard.press("Tab")
        except Exception:
            pass
        page.wait_for_timeout(700)

    def type_into(box, commit: str = "Enter") -> None:
        box.click(timeout=5000)
        page.keyboard.press("ControlOrMeta+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(name, delay=20)  # ketik sungguhan agar Flow mendeteksi perubahan
        page.keyboard.press(commit)
        page.wait_for_timeout(600)

    def name_box():
        box = page.locator(NAME_INPUT).first
        if not box.is_editable(timeout=3000):
            _click_visible(page.locator("button[aria-label='Edit name'], button[aria-label='Edit nama']"))
            page.wait_for_timeout(500)
        return box

    def need_header():
        box = header_title_box(page, writable=True)
        if box is None:
            raise RuntimeError("judul karakter tidak ada")
        return box

    def by_name_input() -> bool:
        type_into(name_box())
        blur()
        return verify()

    def by_header_title() -> bool:
        type_into(need_header())
        blur()
        return verify()

    def by_pencil_typing() -> bool:
        if not _click_visible(page.locator("button[aria-label='Edit name'], button[aria-label='Edit nama']")):
            raise RuntimeError("tombol edit nama tidak ada")
        page.wait_for_timeout(400)
        page.keyboard.press("ControlOrMeta+A")
        page.keyboard.type(name, delay=20)
        page.keyboard.press("Enter")
        blur()
        return verify()

    def by_fill_name() -> bool:
        box = name_box()
        box.fill(name, timeout=5000)
        box.press("Enter", timeout=5000)
        blur()
        return verify()

    def by_fill_header() -> bool:
        box = need_header()
        box.fill(name, timeout=5000)
        box.press("Enter", timeout=5000)
        blur()
        return verify()

    def by_triple_click_insert() -> bool:
        box = name_box()
        box.click(click_count=3, timeout=5000)
        page.keyboard.press("Delete")
        page.keyboard.insert_text(name)
        page.keyboard.press("Enter")
        blur()
        return verify()

    def by_script() -> bool:
        # Set nilai lewat setter asli + event input/change/blur (untuk Angular).
        count = page.evaluate(
            """([name, nameSel, headerSels]) => {
                let n = 0;
                const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
                const header = headerSels.flatMap(s => [...document.querySelectorAll(s)])
                    .find(e => /^\\s*(Karakter tanpa judul|Untitled character|Karakter tanpa nama|Nama karakter|Character name)\\s*$/i.test(e.placeholder || ''));
                for (const el of [document.querySelector(nameSel), header]) {
                    if (!el) continue;
                    el.focus(); setter.call(el, name);
                    for (const t of ['input', 'change']) el.dispatchEvent(new Event(t, {bubbles: true}));
                    el.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
                    el.dispatchEvent(new KeyboardEvent('keyup', {key: 'Enter', bubbles: true}));
                    el.blur(); el.dispatchEvent(new Event('blur', {bubbles: true}));
                    n++;
                }
                return n;
            }""",
            [name, NAME_INPUT, list(HEADER_TITLE_SELECTORS)],
        )
        if not count:
            raise RuntimeError("kolom nama tidak ada")
        page.wait_for_timeout(800)
        return verify()

    def reload_then(method):
        def run() -> bool:
            page.reload(wait_until="domcontentloaded", timeout=30000)
            wait_character_loaded(page, 20)
            return method()
        return run

    return first_success(
        "Isi nama karakter",
        [("kolom Nama karakter", by_name_input), ("judul di atas", by_header_title),
         ("ikon pensil + ketik", by_pencil_typing), ("isi langsung kolom nama", by_fill_name),
         ("isi langsung judul", by_fill_header), ("klik 3x + tempel teks", by_triple_click_insert),
         ("set nilai lewat script", by_script),
         ("muat ulang + kolom Nama karakter", reload_then(by_name_input)),
         ("muat ulang + judul di atas", reload_then(by_header_title)),
         ("muat ulang + set lewat script", reload_then(by_script))],
    )


def finish_character(page: "Page") -> bool:
    def by_aria() -> bool:
        return _click_visible(page.locator("button[aria-label='Done editing']"))

    def by_text() -> bool:
        return _click_visible(page.locator("button").filter(has_text=re.compile(r"^\s*(Done|Selesai)\s*$", re.I)))

    def by_back() -> bool:
        return _click_visible(page.locator("button[aria-label='Back'], button[aria-label='Kembali']"))

    ok = first_success("Simpan karakter", [("tombol Done editing", by_aria), ("tombol Selesai/Done", by_text), ("tombol kembali", by_back)])
    page.wait_for_timeout(2500)
    return ok


def create_character(
    page: "Page", config: dict[str, Any], image: Path, name: str, personality: str,
    track: dict[str, Any] | None = None,
) -> None:
    """Karakter baru > Upload gambar > tunggu selesai diproses > isi nama > Selesai.
    track["url"] berisi alamat halaman karakter yang dibuat (untuk ganti nama/hapus bila gagal)."""
    if track is None:
        track = {}
    open_new_character_page(page, config)
    before = _norm_url(page.url)  # alamat sebelum upload: karakter baru WAJIB beralamat lain
    upload_character_image(page, image)
    upload_started = time.monotonic()
    open_deadline = time.monotonic() + int(config.get("character_open_timeout_s", 120))
    while True:
        started_wait = time.monotonic()
        wait_upload_consent(page)
        open_deadline += max(0.0, time.monotonic() - started_wait - 1)  # waktu menunggu persetujuan tidak dihitung
        current = page.url or ""
        if CHARACTER_URL.search(current) and _norm_url(current) != before:
            break
        if CHARACTER_URL.search(before) and _norm_url(current) == before and time.monotonic() - upload_started > 5:
            break  # tampilan yang sudah memberi alamat /character/<id> sebelum upload
        if time.monotonic() > open_deadline:
            raise RuntimeError("Halaman karakter baru tidak terbuka setelah upload")
        page.wait_for_timeout(1000)
    track["url"] = _norm_url(page.url)
    # Tunggu gambar selesai diproses (indikator % hilang, gambar karakter muncul).
    limit = int(config.get("character_upload_timeout_s", 180))
    # Setelah URL berganti, halaman sempat tampil kosong (tombol Upload) beberapa
    # detik sebelum proses unggah (persen) muncul. Anggap gagal hanya bila tetap
    # kosong tanpa pernah memproses selama empty_grace detik.
    empty_grace = int(config.get("character_upload_empty_grace_s", 60))
    busy_grace = int(config.get("character_upload_after_busy_grace_s", 45))
    started = time.monotonic()
    deadline = started + limit
    seen_busy = False
    empty_since: float | None = None
    toasts: list[str] = []
    state: dict[str, Any] = {}
    failed_empty = False
    while time.monotonic() < deadline:
        wait_upload_consent(page)
        try:
            state = page.evaluate(CHARACTER_STATE_JS)
        except Exception:  # halaman sedang berpindah/dimuat -> cek lagi
            state = {}
            page.wait_for_timeout(1000)
            continue
        if state.get("toast") and state["toast"] not in toasts:
            toasts.append(state["toast"])
        if state.get("ready"):
            break
        if state.get("busy"):
            seen_busy = True
            empty_since = None
        elif state.get("empty"):
            empty_since = empty_since or time.monotonic()
            grace = busy_grace if seen_busy else empty_grace
            if time.monotonic() - empty_since >= grace:
                failed_empty = True
                break  # gambar hilang/ditolak: halaman tetap menampilkan tombol Upload
        page.wait_for_timeout(1000)
    if not state.get("ready") and track.get("url"):
        # Cara 2: Flow kadang baru menampilkan gambar belakangan (setelah persen upload
        # hilang, gambar masih diproses). Muat ulang draf dan tunggu gambarnya dulu
        # sebelum menganggap gagal -> draf yang sebenarnya berhasil tidak ikut dihapus.
        try:
            if open_character_url(page, track["url"]) and wait_character_loaded(
                page, int(config.get("character_upload_recheck_s", 45)), need_image=True
            ) == "ready":
                logging.info("KARAKTER | %s: gambar ternyata masuk (tampil terlambat); lanjut isi nama", name)
                state = {"ready": True}
        except Exception as exc:
            logging.warning("KARAKTER | cek ulang gambar %s gagal: %s", name, str(exc).splitlines()[0][:150])
    if not state.get("ready"):
        reason = "gambar ditolak/hilang setelah upload" if failed_empty else f"gambar belum selesai diproses setelah {limit} detik"
        if toasts:
            reason += f" (pesan Flow: {' | '.join(toasts)[:300]})"
        dump_page(page, "upload_gagal_" + safe_file_name(name))
        if delete_open_character(page, track.get("url"), own_name=name):
            track["deleted"] = True
        raise RuntimeError(reason)
    # Gambar siap: tunggu stabil beberapa detik agar penyimpanan otomatis Flow selesai
    # (kalau nama diisi terlalu cepat, Flow bisa menimpanya kembali jadi "Karakter tanpa judul").
    track["uploaded"] = True
    stable = 0
    for _ in range(20):
        try:
            now = page.evaluate(CHARACTER_STATE_JS)
        except Exception:
            now = {}
        stable = stable + 1 if now.get("ready") else 0
        if stable >= 3:
            break
        page.wait_for_timeout(1000)
    track["name_tried"] = True  # bila semua cara isi nama gagal, jangan diulang lagi di luar
    set_character_name(page, name)
    if personality:
        try:
            page.locator(PERSONALITY_INPUT).first.fill(personality)
        except Exception as exc:
            logging.warning("KARAKTER | deskripsi %s tidak terisi: %s", name, exc)
    page.wait_for_timeout(800)
    if not saved_name_matches(page, name):
        set_character_name(page, name)  # Flow menimpa nama -> isi ulang sebelum Selesai
    finish_character(page)


def rename_created_character(page: "Page", url: str, name: str, rounds: int = 2) -> bool:
    """Buka lagi karakter yang barusan dibuat, isi ulang namanya, lalu buka ulang
    halaman untuk memastikan nama benar-benar tersimpan. True = tersimpan."""
    for round_no in range(1, rounds + 1):
        if not open_character_url(page, url):
            logging.warning("KARAKTER | %s tidak bisa dibuka untuk isi ulang nama", url)
            return False
        wait_character_loaded(page, 25)
        if not saved_name_matches(page, name):
            try:
                set_character_name(page, name)
            except Exception as exc:
                logging.warning("KARAKTER | isi ulang nama %s (putaran %s) gagal: %s", name, round_no, str(exc)[:200])
                return False  # semua cara sudah dicoba (termasuk muat ulang) -> jangan diulang
        try:
            finish_character(page)
        except Exception:
            pass
        # Verifikasi: buka ulang halamannya, nama harus tetap ada.
        if open_character_url(page, url):
            wait_character_loaded(page, 25)
            if saved_name_matches(page, name):
                try:
                    finish_character(page)
                except Exception:
                    pass
                return True
    return False


def character_exists(page: "Page", config: dict[str, Any], name: str) -> bool:
    """Cara 1: kotak cari. Cara 2: baca seluruh daftar karakter."""
    target = name.casefold()

    def by_search() -> bool:
        mode = open_characters_view(page, config)
        if mode != "grid":
            raise RuntimeError("tidak ada grid karakter")
        search = page.locator(SEARCH_INPUT).first
        try:
            search.fill(name, timeout=5000)
            page.wait_for_timeout(2000)
            found = page.evaluate(
                """(name) => [...document.querySelectorAll('flow-grid-tile-container')].some(t =>
                    (t.getAttribute('aria-label') || '').trim().toLowerCase() === name.toLowerCase()
                    && t.querySelector('flow-character-tile, img'))""",
                name,
            )
        finally:
            try:
                search.fill("", timeout=3000)
            except Exception:
                pass
        if not found:
            raise RuntimeError("tidak ditemukan lewat pencarian")
        return True

    def by_full_list() -> bool:
        mode = open_characters_view(page, config)
        return any(c["name"].casefold() == target for c in list_characters(page, config, mode))

    try:
        return bool(first_success("Cek karakter tersimpan", [("kotak cari", by_search), ("daftar lengkap", by_full_list)]))
    except Exception:
        return False


def import_characters(
    page: "Page", config: dict[str, Any], folder: Path, names: list[str] | None, report: set[str] | None = None,
    deadline: float | None = None,
) -> int:
    items = [c for c in load_character_folder(folder) if not UNTITLED_NAME.match(c.get("name", ""))]
    if names:
        wanted = {n.casefold() for n in names}
        items = [c for c in items if c["name"].casefold() in wanted]
    if not items:
        logging.error("KARAKTER | tidak ada karakter di folder %s", folder)
        return 2
    mode = open_characters_view(page, config)
    existing = {c["name"].casefold() for c in list_characters(page, config, mode)}
    created = skipped = failed = 0
    retries = max(0, int(config.get("character_upload_retries", 1)))
    for index, item in enumerate(items, start=1):
        name = item["name"]
        if deadline is not None and time.monotonic() > deadline:
            logging.warning(
                "KARAKTER | batas waktu sinkron habis; %s karakter sisanya dilewati (generate tetap jalan)",
                len(items) - index + 1,
            )
            failed += len(items) - index + 1
            break
        if name.casefold() in existing:
            logging.info("KARAKTER %s/%s | %s sudah ada di profil tujuan, dilewati", index, len(items), name)
            skipped += 1
            if report is not None:
                report.add(norm_name(name))
            continue
        error = ""
        already = False
        for attempt in range(retries + 1):
            if attempt > 0 and deadline is not None and time.monotonic() > deadline:
                error = error or "batas waktu sinkron habis"
                break
            if attempt > 0 or config.get("character_precheck", True):
                # Cegah duplikat: daftar awal bisa belum lengkap, atau percobaan
                # sebelumnya ternyata sudah berhasil tersimpan.
                if character_exists(page, config, name):
                    error = ""
                    if attempt == 0:
                        logging.info("KARAKTER %s/%s | %s sudah ada di profil tujuan, dilewati", index, len(items), name)
                        already = True
                    break
            state: dict[str, Any] = {}
            try:
                create_character(page, config, folder / item["file"], name, item.get("personality", ""), state)
                if character_exists(page, config, name):
                    error = ""
                    break
                # Gambar sudah masuk tetapi namanya tidak tersimpan -> buka lagi & isi ulang nama.
                if state.get("url"):
                    logging.info("KARAKTER %s/%s | %s: nama belum tersimpan, diisi ulang", index, len(items), name)
                    if rename_created_character(page, state["url"], name) or character_exists(page, config, name):
                        error = ""
                        break
                error = "nama tidak tersimpan setelah dibuat"
            except Exception as exc:
                error = str(exc).splitlines()[0][:300]
                # Draf sudah terbentuk tetapi langkah nama error -> coba isi ulang nama dulu.
                if state.get("uploaded") and not state.get("name_tried") and state.get("url") and not state.get("deleted"):
                    try:
                        if rename_created_character(page, state["url"], name):
                            error = ""
                            break
                    except Exception as exc2:
                        logging.warning("KARAKTER | isi ulang nama %s gagal: %s", name, str(exc2)[:200])
            # Gagal: hapus HANYA draf yang barusan dibuat bot (tidak menyisakan "Karakter tanpa judul").
            if state.get("url") and not state.get("deleted"):
                named_ok = False
                try:
                    if open_character_url(page, state["url"]):
                        wait_character_loaded(page, 25)
                        named_ok = saved_name_matches(page, name)
                except Exception:
                    named_ok = False
                if named_ok:
                    # Namanya ternyata benar (daftar/pencarian saja yang terlambat) -> jangan dihapus.
                    logging.info("KARAKTER %s/%s | %s ternyata tersimpan dengan benar", index, len(items), name)
                    try:
                        finish_character(page)
                    except Exception:
                        pass
                    error = ""
                    break
                if delete_open_character(page, state["url"], own_name=name):
                    state["deleted"] = True
                else:
                    logging.warning(
                        "KARAKTER %s/%s | draf %s belum terhapus; bersihkan dengan tombol 'Hapus karakter tanpa nama'",
                        index, len(items), name,
                    )
            if attempt < retries:
                logging.warning("KARAKTER %s/%s | %s gagal (%s); coba lagi", index, len(items), name, error)
                page.wait_for_timeout(15000)
        if already:
            skipped += 1
            existing.add(name.casefold())
            if report is not None:
                report.add(norm_name(name))
            continue
        page.wait_for_timeout(int(config.get("character_pause_ms", 3000)))
        if not error:
            created += 1
            existing.add(name.casefold())
            if report is not None:
                report.add(norm_name(name))
            logging.info("KARAKTER %s/%s | %s berhasil dibuat", index, len(items), name)
        else:
            failed += 1
            logging.error("KARAKTER %s/%s | %s GAGAL: %s", index, len(items), name, error)
    logging.info(
        "KARAKTER SELESAI | dibuat %s, sudah ada %s, gagal %s", created, skipped, failed
    )
    return 0 if failed == 0 else 1


# ---------------------------------------------------------------------------
# Sinkron karakter otomatis sebelum generate
# ---------------------------------------------------------------------------

def norm_name(value: Any) -> str:
    return " ".join(str(value or "").split()).casefold()


def excel_character_names(input_paths: list[Path], config: dict[str, Any]) -> dict[str, str]:
    """Nama karakter yang dipakai baris Excel yang BELUM selesai (kunci = nama dinormalkan)."""
    names: dict[str, str] = {}
    for path in input_paths:
        file_config = config_for_input(config, path)
        for _, row in prompt_rows(path, config):
            for field, settings in file_config.get("fields", {}).items():
                value = row.get(str(settings.get("column", field)))
                if is_empty(value):
                    continue
                text = resolve_alias(value, file_config)
                if text:
                    names.setdefault(norm_name(text), text)
    return names


def load_profiles_info() -> list[dict[str, str]]:
    """Semua profil dari profiles.json + URL project dari profile_urls.json."""
    try:
        profiles = json.loads((APP_DIR / "profiles.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        profiles = {}
    try:
        urls = json.loads((APP_DIR / "profile_urls.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        urls = {}
    result = []
    for name, directory in profiles.items():
        path = Path(directory)
        path = path if path.is_absolute() else APP_DIR / path
        result.append({"name": name, "dir": str(path), "url": str(urls.get(name) or "")})
    return result


def current_profile_name(config: dict[str, Any], explicit: str | None) -> str:
    if explicit:
        return explicit
    own = resolve_path(str(config.get("profile_dir", ""))).resolve()
    for info in load_profiles_info():
        try:
            if Path(info["dir"]).resolve() == own:
                return info["name"]
        except OSError:
            continue
    return own.name


def catalog_path(config: dict[str, Any]) -> Path:
    return resolve_path(str(config.get("character_folder", CHARACTER_ROOT))) / "_katalog.json"


def load_catalog(config: dict[str, Any]) -> dict[str, Any]:
    try:
        return json.loads(catalog_path(config).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_catalog_entry(config: dict[str, Any], profile: str, names: list[str], confirmed_empty: bool = False) -> None:
    """Catat daftar karakter tiap profil (dipakai untuk mencari karakter antar profil)."""
    if not names and not confirmed_empty:
        return  # daftar kosong bisa berarti grid belum termuat; jangan dianggap pasti
    try:
        data = load_catalog(config)
        data[profile] = {"names": sorted(set(names), key=str.casefold), "checked_at": time.time()}
        path = catalog_path(config)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError as exc:
        logging.warning("Katalog karakter tidak tersimpan: %s", exc)


def find_in_saved_folders(config: dict[str, Any], wanted: dict[str, str], skip_profile: str) -> dict[str, tuple[Path, str]]:
    """Cara 1: cari di folder karakter yang sudah pernah diunduh (downloads/_KARAKTER/<profil>)."""
    root = resolve_path(str(config.get("character_folder", CHARACTER_ROOT)))
    found: dict[str, tuple[Path, str]] = {}
    if not root.is_dir():
        return found
    skip = safe_file_name(skip_profile).casefold()
    folders = sorted((f for f in root.iterdir() if f.is_dir()), key=lambda f: f.stat().st_mtime, reverse=True)
    for folder in folders:
        if folder.name.casefold() == skip or folder.name.startswith("_"):
            continue
        try:
            items = load_character_folder(folder)
        except Exception:
            continue
        for item in items:
            key = norm_name(item.get("name"))
            if key in wanted and key not in found:
                found[key] = (folder, item["name"])
    return found


def search_other_profiles_live(
    playwright: "Playwright", config: dict[str, Any], wanted: dict[str, str], current: str
) -> dict[str, tuple[Path, str]]:
    """Cara 2: buka profil lain satu per satu, cari karakter yang belum ketemu, unduh hanya itu."""
    sync = config.get("character_sync", {})
    max_age = float(sync.get("catalog_max_age_hours", 6)) * 3600
    catalog = load_catalog(config)
    found: dict[str, tuple[Path, str]] = {}
    remaining = dict(wanted)
    for info in load_profiles_info():
        if not remaining:
            break
        name = info["name"]
        try:
            same_dir = Path(info["dir"]).resolve() == resolve_path(str(config.get("profile_dir", ""))).resolve()
        except OSError:
            same_dir = False
        if name == current or same_dir or not info["url"] or not Path(info["dir"]).is_dir():
            continue
        entry = catalog.get(name) or {}
        fresh = time.time() - float(entry.get("checked_at", 0)) < max_age
        known = {norm_name(n) for n in entry.get("names", [])}
        if fresh and not (known & set(remaining)):
            continue  # katalog baru saja dicek dan profil ini tidak punya karakter yang dicari
        profile_config = dict(config)
        profile_config["profile_dir"] = info["dir"]
        profile_config["url"] = info["url"]
        logging.info("SINKRON KARAKTER | mencari %s karakter di profil %s", len(remaining), name)
        context = None
        try:
            context = open_context(playwright, profile_config)
            other = context.pages[0] if context.pages else context.new_page()
            other.set_default_timeout(int(config.get("timeout_ms", 30000)))
            mode = open_characters_view(other, profile_config)
            listed = list_characters(other, profile_config, mode)
            grid_loaded = mode == "grid" and bool(other.evaluate("() => !!document.querySelector('flow-custom-tile')"))
            save_catalog_entry(config, name, [c["name"] for c in listed], confirmed_empty=grid_loaded)
            hits = {norm_name(c["name"]): c["name"] for c in listed if norm_name(c["name"]) in remaining}
            if hits:
                export_characters(other, profile_config, name, {n.casefold() for n in hits.values()})
                folder = character_folder(name, config)
                saved = {norm_name(i["name"]): i["name"] for i in load_character_folder(folder)}
                for key in hits:
                    if key in saved:
                        found[key] = (folder, saved[key])
                        remaining.pop(key, None)
        except FlowProjectNotFound:
            logging.warning("SINKRON KARAKTER | project profil %s tidak bisa dibuka; dilewati", name)
        except Exception as exc:
            logging.warning(
                "SINKRON KARAKTER | profil %s tidak bisa dicek (mungkin Chrome profil itu sedang terbuka): %s",
                name, str(exc).splitlines()[0][:160],
            )
        finally:
            if context is not None:
                try:
                    context.close()
                except Exception:
                    pass
    return found


def sync_characters_before_generate(
    playwright: "Playwright", page: "Page", config: dict[str, Any], input_paths: list[Path], profile: str
) -> bool:
    """Sebelum generate: pastikan karakter di Excel ada di profil ini.
    Yang tidak ada dicari di profil lain (folder tersimpan -> buka profil lain), lalu
    dipindahkan dulu. Yang tidak ada di semua profil: dilewati, generate tetap jalan."""
    sync = config.get("character_sync", {})
    if not sync.get("enabled", True):
        return False
    wanted = excel_character_names(input_paths, config)
    if not wanted:
        logging.info("SINKRON KARAKTER | Excel tidak memakai karakter; lewati")
        return False
    logging.info("SINKRON KARAKTER | %s karakter dipakai Excel: %s", len(wanted), ", ".join(sorted(wanted.values())))
    mode = open_characters_view(page, config)
    listed = list_characters(page, config, mode)
    save_catalog_entry(config, profile, [c["name"] for c in listed])
    have = {norm_name(c["name"]) for c in listed}
    missing = {k: v for k, v in wanted.items() if k not in have}
    logging.info("SINKRON KARAKTER | sudah ada di %s: %s, belum ada: %s", profile, len(wanted) - len(missing), len(missing))
    if not missing:
        return True
    for name in sorted(missing.values()):
        logging.info("  BELUM ADA | %s", name)
    found = find_in_saved_folders(config, missing, profile)
    rest = {k: v for k, v in missing.items() if k not in found}
    if rest and sync.get("live_search", True):
        found.update(search_other_profiles_live(playwright, config, rest, profile))
    not_found = [v for k, v in missing.items() if k not in found]
    by_folder: dict[Path, list[str]] = {}
    for folder, source_name in found.values():
        by_folder.setdefault(folder, []).append(source_name)
    moved: set[str] = set()
    sync_deadline = time.monotonic() + float(sync.get("time_budget_s", 1800))
    for folder, names in by_folder.items():
        logging.info("SINKRON KARAKTER | pindahkan %s karakter dari %s: %s", len(names), folder.name, ", ".join(names))
        try:
            import_characters(page, config, folder, names, report=moved, deadline=sync_deadline)
        except Exception as exc:
            logging.warning("SINKRON KARAKTER | pindah dari %s gagal: %s", folder.name, exc)
    for name in not_found:
        logging.info("SINKRON KARAKTER | %s tidak ada di semua profil -> generate tetap jalan (karakter dari prompt)", name)
    still_missing = {k for k in missing if k not in moved}
    for key in sorted(still_missing - {norm_name(n) for n in not_found}):
        logging.warning("SINKRON KARAKTER | %s gagal dipindah -> generate tetap jalan tanpa referensi", missing[key])
    # Referensi yang PASTI tidak ada dilewati saat generate (lebih cepat). Dipastikan
    # dulu lewat kotak cari agar daftar yang terbaca sebagian tidak salah melewati.
    confirmed: list[str] = []
    for key in sorted(still_missing)[: int(sync.get("confirm_limit", 15))]:
        if time.monotonic() > sync_deadline + 300:
            confirmed.append(key)  # waktu habis: anggap tidak ada (referensi dilewati)
            continue
        if not character_exists(page, config, missing[key]):
            confirmed.append(key)
        else:
            logging.info("SINKRON KARAKTER | %s ternyata sudah ada di %s", missing[key], profile)
    config["_missing_characters"] = confirmed
    return True


# Alamat thumbnail Flow bertanda tangan (?Expires=...) dan berubah tiap dimuat; identitasnya bagian sebelum '?'.
TILE_INFO_JS = """(t) => ({label: t.getAttribute('aria-label') || '',
    src: ((t.querySelector('img') || {}).src || '').split('?')[0], isChar: !!t.querySelector('flow-character-tile')})"""


def _find_untitled_tile(page: "Page", tried: set[tuple[str, int, int]], skip_srcs: set[str]) -> dict[str, Any] | None:
    """Cari kartu karakter berlabel kosong/"Karakter tanpa judul" yang belum dicoba.
    Cara 1: cari "tanpa judul"; cara 2: "Untitled"; cara 3: tanpa kata kunci + gulir seluruh grid.
    Kartu dikenali dari gambar thumbnail-nya (src), jadi yang sudah dilewati tidak dibuka lagi."""
    for term in ("tanpa judul", "Untitled", ""):
        try:
            page.locator(SEARCH_INPUT).first.fill(term, timeout=5000)
            page.wait_for_timeout(2000)
        except Exception:
            if term:
                continue
        for step in range(80 if term == "" else 1):
            try:
                tiles = page.evaluate(
                    """() => [...document.querySelectorAll('flow-grid-tile-container')]
                        .filter(t => t.querySelector('flow-character-tile'))
                        .map(t => ({label: t.getAttribute('aria-label') || '',
                                    src: ((t.querySelector('img') || {}).src || '').split('?')[0]}))"""
                )
            except Exception:
                tiles = []
            counts: dict[str, int] = {}
            for tile in tiles:
                counts[tile.get("src") or ""] = counts.get(tile.get("src") or "", 0) + 1
            for index, tile in enumerate(tiles):
                src = tile.get("src") or ""
                if not UNTITLED_NAME.match(tile.get("label") or ""):
                    continue
                unique = bool(src) and counts.get(src) == 1  # gambar dipakai 1 kartu saja -> bisa jadi identitas
                if (unique and src in skip_srcs) or (term, step, index) in tried:
                    continue
                return {"term": term, "step": step, "index": index, "src": src, "unique": unique}
            if term != "":
                break
            try:
                moved = page.evaluate(
                    "(() => { const el = (" + CHAR_SCROLLER_JS + ")(); const top = el.scrollTop;"
                    " el.scrollTop = top + Math.max(300, el.clientHeight * 0.8); return el.scrollTop !== top; })()"
                )
            except Exception:
                moved = False
            page.wait_for_timeout(900)
            if not moved:
                break
    return None


def _open_tile(page: "Page", target: dict[str, Any]) -> bool:
    """Buka kartu yang dipilih. Sebelum diklik, label & gambar kartu dicek lagi supaya
    kartu yang bergeser (grid virtual) tidak salah dibuka."""
    tiles = page.locator("flow-grid-tile-container").filter(has=page.locator("flow-character-tile"))
    tile = tiles.nth(target["index"])
    tile.scroll_into_view_if_needed(timeout=5000)
    page.wait_for_timeout(400)

    def check_tile() -> None:
        info = tile.evaluate(TILE_INFO_JS)
        if not UNTITLED_NAME.match(info.get("label") or ""):
            raise RuntimeError(f"kartu bergeser (label '{info.get('label')}')")
        if target.get("src") and info.get("src") != target["src"]:
            raise RuntimeError("kartu bergeser (gambar berbeda)")

    def opened() -> bool:
        page.wait_for_url(CHARACTER_URL, timeout=10000)
        return True

    def by_dblclick() -> bool:
        check_tile()
        tile.dblclick(timeout=8000)
        return opened()

    def by_click() -> bool:
        check_tile()
        tile.click(timeout=8000)
        return opened()

    def by_inner() -> bool:
        check_tile()
        tile.locator("flow-character-tile, img").first.click(timeout=8000)
        return opened()

    def by_enter() -> bool:
        check_tile()
        tile.focus(timeout=5000)
        page.keyboard.press("Enter")
        return opened()

    return first_success("Buka kartu karakter", [
        ("klik dua kali", by_dblclick), ("klik sekali", by_click), ("klik gambar", by_inner), ("tombol Enter", by_enter),
    ])


def clean_untitled_characters(page: "Page", config: dict[str, Any], preview: bool = False) -> int:
    """Hapus karakter TANPA NAMA ("Karakter tanpa judul"/"Untitled character") sisa kegagalan.
    Karakter bernama tidak pernah disentuh: kartunya dicek, halaman karakternya dibuka,
    ditunggu sampai termuat, namanya dibaca 2x; hanya bila judul terbaca DAN kosong baru dihapus.
    preview=True: hanya menghitung (tidak ada yang dihapus)."""
    deleted = 0
    found_untitled = 0
    kept: set[str] = set()      # alamat karakter yang ternyata bernama
    unsure: set[str] = set()    # alamat yang namanya tidak terbaca pasti
    failed: set[str] = set()    # alamat yang gagal dihapus
    skip_srcs: set[str] = set() # thumbnail kartu yang sudah selesai dicek/dilewati
    tried: set[tuple[str, int, int]] = set()
    opens = 0
    repeats = 0  # kartu yang sudah dikenal terbuka lagi (sejak penghapusan terakhir)
    limit_repeat = False
    no_menu = False
    max_opens = int(config.get("clean_untitled_max", 300))
    label = "CEK" if preview else "BERSIH"
    while opens < max_opens:
        mode = ""
        for _attempt in range(2):
            try:
                mode = open_characters_view(page, config)
                break
            except FlowProjectNotFound:
                raise
            except Exception as exc:
                logging.warning("%s | menu Karakter gagal dibuka: %s", label, str(exc).splitlines()[0][:200])
                page.wait_for_timeout(3000)
        if mode != "grid":
            logging.error("%s | menu Karakter tidak ditemukan; berhenti", label)
            no_menu = True
            break
        target = _find_untitled_tile(page, tried, skip_srcs)
        if not target:
            break
        tried.add((target["term"], target["step"], target["index"]))
        opens += 1
        try:
            _open_tile(page, target)
        except Exception as exc:
            logging.warning("%s | kartu tanpa nama tidak bisa dibuka: %s", label, str(exc)[:200])
            continue
        url = _norm_url(page.url)

        def skip(bucket: set[str]) -> None:
            bucket.add(url)
            if target.get("unique"):
                skip_srcs.add(target["src"])

        if url in kept or url in failed or url in unsure:
            if target.get("unique"):
                skip_srcs.add(target["src"])
            repeats += 1
            if repeats > 2 * len(kept | failed | unsure) + 5:
                logging.warning("%s | kartu yang sama terus terbuka ulang; pemeriksaan dihentikan", label)
                limit_repeat = True
                break
            continue
        loaded = wait_character_loaded(page, 25)
        if loaded is None:
            logging.info("%s | %s belum termuat; demi aman dilewati", label, url)
            skip(unsure)
            continue
        names = read_names_stable(page)
        readable = [v for v in (names or ()) if v is not None]
        if names is None or not readable:
            logging.info("%s | nama %s tidak terbaca pasti; demi aman dilewati", label, url)
            skip(unsure)
            continue
        if not all(UNTITLED_NAME.match(v) for v in readable):
            name = next(v for v in readable if not UNTITLED_NAME.match(v))
            logging.info("%s | %s ternyata punya nama (%s); dilewati", label, url, name)
            skip(kept)
            continue
        found_untitled += 1
        if preview:
            logging.info("CEK | karakter tanpa nama ke-%s: %s (tidak dihapus, mode cek)", found_untitled, url)
            skip(kept)
            continue
        if delete_open_character(page, url, strict=True):
            deleted += 1
            repeats = 0
            tried.clear()  # urutan kartu bergeser setelah ada yang dihapus
            logging.info("BERSIH | karakter tanpa nama ke-%s dihapus", deleted)
        else:
            skip(failed)
    try:
        page.locator(SEARCH_INPUT).first.fill("", timeout=3000)
    except Exception:
        pass
    limit_hit = opens >= max_opens
    if limit_hit:
        logging.warning("%s | batas %s kali buka kartu tercapai; jalankan lagi bila masih ada", label, max_opens)
    if failed:
        logging.warning("BERSIH | %s karakter tanpa nama gagal dihapus; coba jalankan lagi", len(failed))
    if unsure:
        logging.warning("%s | %s karakter dilewati karena namanya tidak terbaca pasti", label, len(unsure))
    if preview:
        logging.info("CEK SELESAI | %s karakter tanpa nama ditemukan (tidak ada yang dihapus)", found_untitled)
    else:
        logging.info("BERSIH SELESAI | %s karakter tanpa nama dihapus (dilewati %s)", deleted, len(kept))
    return 0 if not (failed or unsure or no_menu or limit_hit or limit_repeat) else 1


def compare_characters(args: argparse.Namespace, config: dict[str, Any]) -> int:
    """Cek karakter profil sumber yang BELUM ada di profil tujuan, lalu unduh hanya itu."""
    from playwright.sync_api import sync_playwright

    target_config = dict(config)
    target_config["profile_dir"] = args.target_profile_dir
    target_config["url"] = args.target_url
    source_name = args.profile_name or "sumber"
    target_name = args.target_name or "tujuan"
    with sync_playwright() as playwright:
        logging.info("CEK KARAKTER | membaca daftar karakter profil tujuan: %s", target_name)
        context = open_context(playwright, target_config)
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(int(config.get("timeout_ms", 30000)))
        try:
            mode = open_characters_view(page, target_config)
            target_names = [c["name"] for c in list_characters(page, target_config, mode)]
            save_catalog_entry(config, target_name, target_names)
        finally:
            context.close()
        logging.info("CEK KARAKTER | %s punya %s karakter", target_name, len(target_names))

        logging.info("CEK KARAKTER | membaca daftar karakter profil sumber: %s", source_name)
        context = open_context(playwright, config)
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(int(config.get("timeout_ms", 30000)))
        try:
            mode = open_characters_view(page, config)
            source_names = [c["name"] for c in list_characters(page, config, mode)]
            save_catalog_entry(config, source_name, source_names)
            if not source_names:
                logging.warning("CEK KARAKTER | tidak ada karakter terbaca di %s; menyimpan diagnosa", source_name)
                dump_character_diagnostics(page, config, source_name)
            existing = {n.casefold() for n in target_names}
            missing = [n for n in source_names if n.casefold() not in existing]
            already = [n for n in source_names if n.casefold() in existing]
            logging.info(
                "CEK KARAKTER | %s karakter di %s; BELUM ADA di %s: %s; sudah ada: %s",
                len(source_names), source_name, target_name, len(missing), len(already),
            )
            for name in missing:
                logging.info("  BELUM ADA  | %s", name)
            for name in already:
                logging.info("  SUDAH ADA  | %s", name)
            if missing:
                export_characters(page, config, source_name, {n.casefold() for n in missing})
        finally:
            context.close()
    folder = character_folder(source_name, config)
    folder.mkdir(parents=True, exist_ok=True)
    result = {
        "source": source_name, "target": target_name,
        "missing": missing, "existing": already, "target_characters": target_names,
        "checked_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    (folder / f"_banding_{safe_file_name(target_name)}.json").write_text(
        json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    logging.info("CEK KARAKTER SELESAI | hasil disimpan di %s", folder)
    return 0


def character_main(args: argparse.Namespace, config: dict[str, Any]) -> int:
    from playwright.sync_api import sync_playwright

    if args.compare_characters:
        try:
            return compare_characters(args, config)
        except FlowProjectNotFound as exc:
            logging.critical("PROJECT TIDAK DITEMUKAN | %s", exc)
            return 6
        except Exception as exc:
            logging.exception("CEK KARAKTER GAGAL | %s", exc)
            return 1
    with sync_playwright() as playwright:
        context = open_context(playwright, config)
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(int(config.get("timeout_ms", 30000)))
        try:
            if args.clean_untitled:
                return clean_untitled_characters(page, config, preview=args.check_only)
            if args.export_characters:
                return export_characters(page, config, args.profile_name or "profil")
            folder = Path(args.character_folder)
            names = args.names
            if args.names_file:
                names = json.loads(Path(args.names_file).read_text(encoding="utf-8"))
            return import_characters(page, config, folder, names)
        except FlowProjectNotFound as exc:
            logging.critical("PROJECT TIDAK DITEMUKAN | %s", exc)
            return 6
        except Exception as exc:
            logging.exception("KARAKTER GAGAL | %s", exc)
            return 1
        finally:
            try:
                context.close()
            except Exception:
                pass


def _main() -> int:
    parser = argparse.ArgumentParser(description="Bot Google Chrome dengan profil terisolasi")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--input", help="Ganti file input dari konfigurasi")
    parser.add_argument("--inputs", nargs="+", help="Proses beberapa CSV/XLSX secara berurutan")
    parser.add_argument("--profile-dir", help="Folder profil Chrome khusus yang akan dipakai")
    parser.add_argument("--url", help="URL project Google Flow milik akun profil ini (mengganti url di config)")
    parser.add_argument("--login", action="store_true", help="Buka browser untuk login, tanpa memproses data")
    parser.add_argument("--dry-run", action="store_true", help="Validasi data dan logika skip tanpa klik situs")
    parser.add_argument("--max-rows", type=int, help="Batasi jumlah baris untuk pengujian")
    parser.add_argument("--start-row", type=int, help="Mulai dari nomor baris Excel tertentu")
    parser.add_argument(
        "--no-long-wait", action="store_true",
        help="Bila tetap diblokir 'unusual activity' setelah model cadangan, langsung berhenti (kode 10) "
             "tanpa jeda panjang, agar aplikasi bisa pindah ke profil berikutnya",
    )
    parser.add_argument("--refresh-every-row-test", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--export-characters", action="store_true", help="Unduh semua karakter project ke downloads/_KARAKTER/<profil>")
    parser.add_argument("--import-characters", action="store_true", help="Buat karakter dari folder ke project profil ini")
    parser.add_argument("--profile-name", help="Nama profil (untuk nama folder karakter)")
    parser.add_argument("--character-folder", help="Folder karakter hasil export")
    parser.add_argument("--names", nargs="*", help="Nama karakter yang dipindah (kosong = semua)")
    parser.add_argument("--names-file", help="File JSON berisi daftar nama karakter yang dipindah")
    parser.add_argument("--compare-characters", action="store_true", help="Bandingkan karakter sumber vs tujuan")
    parser.add_argument("--target-profile-dir", help="Folder profil Chrome tujuan")
    parser.add_argument("--target-url", help="URL project Flow profil tujuan")
    parser.add_argument("--target-name", help="Nama profil tujuan")
    parser.add_argument("--clean-untitled", action="store_true", help="Hapus karakter tanpa nama di profil ini")
    parser.add_argument("--check-only", action="store_true", help="Dengan --clean-untitled: hanya hitung, tidak menghapus")
    args = parser.parse_args()

    config_path = resolve_path(args.config)
    config = json.loads(config_path.read_text(encoding="utf-8"))
    if args.start_row:
        config["start_row"] = max(2, args.start_row)
    if args.no_long_wait:
        config["unusual_activity_max_waits"] = 0
    if args.profile_dir:
        config["profile_dir"] = args.profile_dir
    if args.url:
        config["url"] = args.url.strip()
    if args.export_characters or args.import_characters or args.compare_characters or args.clean_untitled:
        log_dir = APP_DIR / "runtime" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(levelname)s | %(message)s",
            handlers=[log_file_handler(log_dir), logging.StreamHandler()],
        )
        return character_main(args, config)
    raw_inputs = args.inputs or ([args.input] if args.input else [config["input_file"]])
    input_paths = []
    for raw_input in raw_inputs:
        input_path = resolve_path(raw_input)
        csv_fallback = input_path.with_suffix(".csv")
        if not input_path.is_file() and csv_fallback.is_file():
            print(f"INFO: {input_path.name} tidak ditemukan; memakai {csv_fallback.name}.")
            input_path = csv_fallback
        if not input_path.is_file():
            print(f"ERROR: File data tidak ditemukan: {input_path}", file=sys.stderr)
            return 2
        input_paths.append(input_path)
    log_dir = APP_DIR / "runtime" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[log_file_handler(log_dir), logging.StreamHandler()],
    )

    if args.dry_run:
        processed = 0
        totals = [(path, prompt_rows(path, config)) for path in input_paths]
        logging.info("Total: %s file, %s baris berisi prompt", len(totals), sum(len(rows) for _, rows in totals))
        for file_index, (input_path, rows) in enumerate(totals, start=1):
            file_config = config_for_input(config, input_path)
            logging.info("FILE %s/%s | %s | %s prompt", file_index, len(totals), input_path.name, len(rows))
            docx_path = find_narration_docx(input_path)
            logging.info("[SIMULASI] Narasi docx: %s", docx_path.name if docx_path else "tidak ditemukan")
            for prompt_index, (row_number, row) in enumerate(rows, start=1):
                logging.info("PROGRES | %s | prompt %s/%s | baris Excel %s", input_path.name, prompt_index, len(rows), row_number)
                process_row(None, row_number, row, file_config, True)
                processed += 1
                if args.max_rows and processed >= args.max_rows:
                    return 0
        return 0

    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = open_context(playwright, config)
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(int(config.get("timeout_ms", 30000)))
        page.goto(config["url"], wait_until="domcontentloaded")
        if args.login:
            input("Login di jendela Chrome khusus bot ini, lalu tekan Enter di sini...")
            context.close()
            return 0
        logging.info("Membuka project Flow: %s", config["url"])
        try:
            wait_flow_ready(page, config)
        except FlowProjectNotFound as exc:
            logging.critical(
                "PROJECT TIDAK DITEMUKAN | %s | Akun Google di profil ini tidak punya project tersebut. "
                "Di aplikasi klik 'URL project', lalu tempel URL project Flow milik akun ini "
                "(project yang berisi karakter referensinya).", exc,
            )
            context.close()
            return 6

        navigated = True
        try:
            navigated = sync_characters_before_generate(
                playwright, page, config, input_paths, current_profile_name(config, args.profile_name)
            )
        except FlowProjectNotFound:
            pass
        except Exception as exc:
            logging.warning("SINKRON KARAKTER dilewati karena error: %s", str(exc).splitlines()[0][:200])
        if navigated:
            # Kembali ke halaman project (Semua media) sebelum generate.
            try:
                page.goto(config["url"], wait_until="domcontentloaded")
                wait_flow_ready(page, config)
            except FlowProjectNotFound as exc:
                logging.critical("PROJECT TIDAK DITEMUKAN | %s", exc)
                context.close()
                return 6
            except Exception as exc:
                logging.warning("Flow belum siap setelah sinkron karakter (%s); muat ulang", str(exc)[:120])
                try:
                    refresh_for_next(page, config, attempts=3, reason="kembali setelah sinkron karakter")
                except Exception:
                    logging.exception("Flow tidak bisa dipulihkan setelah sinkron karakter")
                    try:
                        context.close()
                    except Exception:
                        pass
                    return 4

        screenshot_dir = APP_DIR / "runtime" / "screenshots"
        processed = 0
        run_state: dict[str, Any] = {
            "generate_count": 0,
            "fallback_active": False,
            "active_model": None,
        }
        jobs = [(path, prompt_rows(path, config)) for path in input_paths]
        total_prompts = sum(len(rows) for _, rows in jobs)
        logging.info("Total: %s file, %s baris berisi prompt", len(jobs), total_prompts)
        overall_prompt = 0
        for file_index, (input_path, rows) in enumerate(jobs, start=1):
            file_config = config_for_input(config, input_path)
            logging.info("FILE %s/%s | %s | %s prompt", file_index, len(jobs), input_path.name, len(rows))
            marker = open_marker(input_path, file_config)
            for prompt_index, (row_number, row) in enumerate(rows, start=1):
                overall_prompt += 1
                logging.info(
                    "PROGRES %s/%s | %s | prompt %s/%s | baris Excel %s",
                    overall_prompt, total_prompts, input_path.name, prompt_index, len(rows), row_number,
                )
                should_stop = False
                downloaded: list[str] = []
                transition_refresh_done = False
                retry_row = True
                pace_row(page, config, overall_prompt)
                while retry_row:
                  retry_row = False
                  try:
                    schedule = file_config.get("model_schedule", {})
                    if schedule.get("enabled", False):
                        desired_model = scheduled_generation_model(file_config, run_state)
                        ensure_generation_model(page, file_config, desired_model)
                        run_state["active_model"] = desired_model
                    downloaded = process_row(
                        page, row_number, row, file_config, False, run_state=run_state
                    )
                    safe_mark(marker, row_number, downloaded)
                    # Generate tidak kena kredit habis/pembatasan -> hitungan "berturut-turut" direset.
                    run_state["fail_streak"] = 0
                  except FlowPolicyRejected as exc:
                    # Ditolak kebijakan Flow: bukan kredit/pembatasan -> tandai DITOLAK, lanjut baris berikutnya.
                    logging.warning("%s baris %s: %s", input_path.name, row_number, exc)
                    safe_mark(marker, row_number, [], rejected=True)
                    run_state["fail_streak"] = 0
                  except (FlowBlockedError, FlowLimitError) as exc:
                    # Kredit habis / pembatasan ("unusual activity"):
                    # 1) masih di Nano Banana 2 -> langsung pindah ke Nano Banana 2 Lite, ulang baris ini;
                    # 2) di Lite: gagal 2x BERTURUT-TURUT (kredit habis atau pembatasan) -> berhenti,
                    #    aplikasi pindah ke profil berikutnya. Bila percobaan ke-2 berhasil, lanjut biasa.
                    blocked = isinstance(exc, FlowBlockedError)
                    kind = "PEMBATASAN (aktivitas tidak biasa)" if blocked else "KREDIT HABIS"
                    schedule = file_config.get("model_schedule", {})
                    fallback_model = str(schedule.get("fallback_model", "Nano Banana 2 Lite")).strip()
                    switch_allowed = schedule.get("switch_on_unusual_activity", True) if blocked \
                        else schedule.get("switch_on_limit", True)
                    can_switch = (
                        schedule.get("enabled", False)
                        and switch_allowed
                        and not run_state.get("fallback_active")
                        and fallback_model
                        and (run_state.get("active_model") or "").casefold() != fallback_model.casefold()
                    )
                    if can_switch:
                        logging.warning(
                            "%s di %s | langsung pindah ke %s lalu mengulang baris %s",
                            kind, run_state.get("active_model") or schedule.get("primary_model", "model utama"),
                            fallback_model, row_number,
                        )
                        try:
                            refresh_for_next(
                                page, file_config,
                                timeout_ms=int(schedule.get("reload_ready_timeout_ms", 600000)),
                                attempts=int(schedule.get("reload_attempts", 3)),
                                reason="pindah model karena " + ("pembatasan" if blocked else "kredit habis"),
                            )
                            ensure_generation_model(page, file_config, fallback_model)
                            run_state["fallback_active"] = True
                            run_state["active_model"] = fallback_model
                            run_state["fail_streak"] = 0
                            logging.info(
                                "TRANSISI MODEL SELESAI | generate berikutnya memakai %s; bila %s gagal %s kali "
                                "berturut-turut (kredit habis/pembatasan), pindah ke profil berikutnya",
                                fallback_model, fallback_model, int(config.get("fallback_fail_limit", 2)),
                            )
                            retry_row = True
                            continue
                        except Exception:
                            logging.exception("Pindah ke %s gagal; profil ini dihentikan", fallback_model)
                            run_state["fail_streak"] = 10 ** 6  # langsung ke keputusan berhenti di bawah
                    fail_limit = max(1, int(config.get("fallback_fail_limit", 2)))
                    streak = int(run_state.get("fail_streak", 0)) + 1
                    run_state["fail_streak"] = streak
                    if streak < fail_limit:
                        logging.warning(
                            "%s di %s | gagal %s/%s berturut-turut; baris %s diulang sekali lagi",
                            kind, run_state.get("active_model") or "model aktif", streak, fail_limit, row_number,
                        )
                        try:
                            page.wait_for_timeout(
                                int(config.get("unusual_activity_retry_pause_ms", 15000)) if blocked else 3000
                            )
                            refresh_for_next(page, file_config, reason="mengulang setelah " + kind.lower())
                        except Exception:
                            logging.exception("Halaman gagal dipulihkan; profil ini dihentikan")
                        else:
                            retry_row = True
                            continue
                    safe_mark(marker, row_number, [], ("pembatasan" if blocked else "kredit habis") + " - diulang di profil lain")
                    logging.critical(
                        "PROFIL BERHENTI | %s di %s (%s). Aplikasi pindah ke profil berikutnya bila ada; "
                        "Excel dilanjutkan dari baris %s.",
                        kind, run_state.get("active_model") or "model aktif", exc, row_number,
                    )
                    try:
                        context.close()
                    except Exception:
                        pass
                    return 10 if blocked else 9
                  except Exception as exc:
                    run_state["fail_streak"] = 0  # error lain (bukan kredit/pembatasan): lewati baris, lanjut
                    logging.exception("%s baris %s gagal", input_path.name, row_number)
                    safe_mark(marker, row_number, [], str(exc).splitlines()[0][:120] if str(exc) else "error")
                    if page.is_closed():
                        logging.critical(
                            "BROWSER TERTUTUP | proses dihentikan pada %s baris %s; baris berikutnya tidak ditandai selesai",
                            input_path.name, row_number,
                        )
                        try:
                            context.close()
                        except Exception:
                            pass
                        return 3
                    screenshot_dir.mkdir(parents=True, exist_ok=True)
                    try:
                        safe_screenshot(page, screenshot_dir / f"error-{input_path.stem}-row-{row_number}.png")
                    except Exception:
                        logging.exception("Screenshot kegagalan tidak dapat disimpan")
                    if config.get("stop_on_error", False):
                        should_stop = True

                processed += 1
                schedule = file_config.get("model_schedule", {})
                primary_count = int(schedule.get("primary_generate_count", 0) or 0)
                if (
                    schedule.get("enabled", False)
                    and primary_count > 0
                    and not run_state.get("fallback_active")
                    and int(run_state.get("generate_count", 0)) >= primary_count
                ):
                    fallback_model = str(
                        schedule.get("fallback_model", "Nano Banana 2 Lite")
                    ).strip()
                    logging.info(
                        "BATCH UTAMA SELESAI | %s generate memakai %s; "
                        "memulai reload wajib sebelum beralih ke %s",
                        primary_count,
                        schedule.get("primary_model", "Nano Banana 2"),
                        fallback_model,
                    )
                    try:
                        refresh_for_next(
                            page,
                            file_config,
                            timeout_ms=int(schedule.get("reload_ready_timeout_ms", 600000)),
                            attempts=int(schedule.get("reload_attempts", 3)),
                            reason=f"transisi model setelah generate ke-{primary_count}",
                        )
                        ensure_generation_model(page, file_config, fallback_model)
                        run_state["fallback_active"] = True
                        run_state["active_model"] = fallback_model
                        transition_refresh_done = True
                        logging.info(
                            "TRANSISI MODEL SELESAI | semua generate berikutnya memakai %s",
                            fallback_model,
                        )
                    except Exception:
                        logging.exception(
                            "Reload/transisi model gagal; bot dihentikan agar tidak "
                            "menghasilkan prompt berikutnya dengan model yang salah"
                        )
                        context.close()
                        return 5
                if should_stop:
                    context.close()
                    return 1
                max_rows = args.max_rows or int(config.get("max_rows_per_run") or 0)
                if max_rows and processed >= max_rows:
                    logging.info("BATAS UJI | %s baris diproses (max_rows_per_run); bot berhenti", processed)
                    context.close()
                    return 0

                expected = int(file_config.get("download", {}).get("outputs_per_row", 2))
                if (len(downloaded) < expected or args.refresh_every_row_test) and not transition_refresh_done:
                    try:
                        refresh_for_next(page, file_config)
                    except Exception:
                        logging.exception("Halaman gagal dipulihkan; bot dihentikan agar baris berikutnya tidak terlewat")
                        context.close()
                        return 4
                else:
                    logging.info("Semua hasil berhasil; lanjut tanpa refresh")
            logging.info("FILE SELESAI | %s", input_path.name)
        context.close()
    return 0


def main() -> int:
    """Pembungkus: error tak terduga dicatat ke log/panel Progres (kode 1),
    bukan jendela 'Unhandled exception' dari EXE."""
    try:
        return _main()
    except SystemExit:
        raise
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        if not logging.getLogger().handlers:
            logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
        logging.exception("BOT BERHENTI karena error: %s", exc)
        return 1


def _redirect_old_exe_launcher() -> None:
    """EXE lama (tampilan bawaan) memuat bot.py ini saat dibuka. Bila launcher.py
    terbaru ada di samping EXE, jalankan tampilan terbaru itu tanpa build ulang."""
    if (not getattr(sys, "frozen", False) or "--worker" in sys.argv
            or getattr(sys, "_flowbot_external", False)):
        return
    app_dir = Path(sys.executable).resolve().parent
    launcher = app_dir / "launcher.py"
    log = app_dir / "runtime" / "logs" / "launcher.log"

    def note(text: str) -> None:
        try:
            log.parent.mkdir(parents=True, exist_ok=True)
            with log.open("a", encoding="utf-8") as handle:
                handle.write(time.strftime("%Y-%m-%d %H:%M:%S ") + text + "\n")
        except OSError:
            pass

    if not launcher.is_file():
        return
    note("membuka tampilan terbaru dari launcher.py")
    try:
        code = compile(launcher.read_text(encoding="utf-8"), str(launcher), "exec")
        sys._flowbot_external = True  # type: ignore[attr-defined]
        namespace = {"__name__": "__main__", "__file__": str(launcher), "__FLOWBOT_EXTERNAL__": True}
        exec(code, namespace)
    except SystemExit:
        raise
    except BaseException:
        import traceback
        note("GAGAL:\n" + traceback.format_exc())
        sys._flowbot_external = False  # type: ignore[attr-defined]
        return
    raise SystemExit(0)


if __name__ != "__main__":
    _redirect_old_exe_launcher()


if __name__ == "__main__":
    raise SystemExit(main())
