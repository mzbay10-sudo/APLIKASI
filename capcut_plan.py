"""Buat sheet "Urutan CapCut" dari Excel FlowBot + narasi Word + SRT dubbing.

Alur:
  1. Kata-kata SRT dicocokkan ke kata-kata paragraf Word (TTS_SAFE), sehingga
     setiap paragraf (¶0, ¶1, ...) mendapat waktu mulai yang tepat.
  2. Setiap scene di Excel FlowBot punya Source Ref (¶). Scene dengan ¶ sama
     berbagi waktu paragraf tersebut, dipotong di awal kalimat SRT.
  3. Setiap scene memakai gambar hasil generate (FILE GAMBAR / folder downloads).
     Scene panjang memakai 2 gambar, scene pendek 1 gambar, scene tanpa gambar
     durasinya disambung ke gambar sebelumnya.
  4. Hasil ditulis ke sheet "Urutan CapCut" (Urutan | Nama file | Mulai |
     Selesai | Durasi | Catatan) di Excel yang sama, siap dibaca
     CapCut Studio Sync.
"""

from __future__ import annotations

import argparse
import difflib
import logging
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import bot

SHEET_NAME = "Urutan CapCut"
MIN_CLIP = 0.8            # detik, batas minimal CapCut Studio Sync
TWO_IMAGE_MIN = 6.0       # scene >= 6 detik memakai 2 gambar
LONG_CLIP = 12.0          # peringatan bila satu gambar tampil lebih lama dari ini
IMAGE_EXT = (".jpeg", ".jpg", ".png", ".webp")


# ---------------------------------------------------------------------------
# SRT
# ---------------------------------------------------------------------------

@dataclass
class Cue:
    start: float
    end: float
    text: str


def _srt_time(value: str) -> float:
    hours, minutes, rest = value.strip().replace(".", ",").split(":")
    seconds, millis = rest.split(",")
    return int(hours) * 3600 + int(minutes) * 60 + int(seconds) + int(millis) / 1000


def read_srt(path: Path) -> list[Cue]:
    text = path.read_text(encoding="utf-8-sig", errors="replace").replace("\r\n", "\n")
    cues = []
    for block in re.split(r"\n\s*\n", text.strip()):
        lines = [line for line in block.split("\n") if line.strip()]
        for index, line in enumerate(lines):
            if "-->" in line:
                start, end = line.split("-->")
                cues.append(Cue(_srt_time(start), _srt_time(end.split()[0]), " ".join(lines[index + 1:]).strip()))
                break
    cues.sort(key=lambda cue: cue.start)
    return cues


def fmt(seconds: float) -> str:
    millis = int(round(max(0.0, seconds) * 1000))
    hours, millis = divmod(millis, 3_600_000)
    minutes, millis = divmod(millis, 60_000)
    secs, millis = divmod(millis, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}.{millis:03d}"


# ---------------------------------------------------------------------------
# Cocokkan SRT ke paragraf Word
# ---------------------------------------------------------------------------

def _words(text: str) -> list[str]:
    return [word for word in re.findall(r"\w+", text.casefold()) if word]


def paragraph_starts(paragraphs: list[str], cues: list[Cue]) -> tuple[list[float | None], float]:
    """Waktu mulai (detik) setiap paragraf Word berdasarkan SRT, plus rasio kata cocok."""
    doc_words: list[str] = []
    doc_para: list[int] = []
    for index, paragraph in enumerate(paragraphs):
        for word in _words(paragraph):
            doc_words.append(word)
            doc_para.append(index)

    srt_words: list[str] = []
    srt_time: list[float] = []
    for cue in cues:
        words = _words(cue.text)
        if not words:
            continue
        # Perkirakan waktu tiap kata di dalam cue berdasarkan panjang huruf.
        lengths = [len(word) + 1 for word in words]
        total = sum(lengths)
        offset = 0
        for word, length in zip(words, lengths):
            srt_words.append(word)
            srt_time.append(cue.start + (cue.end - cue.start) * offset / total)
            offset += length

    matcher = difflib.SequenceMatcher(None, doc_words, srt_words, autojunk=False)
    word_time: list[float | None] = [None] * len(doc_words)
    matched = 0
    for block in matcher.get_matching_blocks():
        for k in range(block.size):
            word_time[block.a + k] = srt_time[block.b + k]
        matched += block.size

    starts: list[float | None] = [None] * len(paragraphs)
    for index, para in enumerate(doc_para):
        if starts[para] is None and word_time[index] is not None:
            # Kata pertama yang cocok; koreksi bila kata sebelumnya tidak cocok.
            starts[para] = word_time[index]
    # Paragraf yang tidak cocok sama sekali: interpolasi dari tetangganya.
    known = [(i, t) for i, t in enumerate(starts) if t is not None]
    for i, value in enumerate(starts):
        if value is None and known:
            before = [(j, t) for j, t in known if j < i]
            after = [(j, t) for j, t in known if j > i]
            if before and after:
                (j0, t0), (j1, t1) = before[-1], after[0]
                starts[i] = t0 + (t1 - t0) * (i - j0) / (j1 - j0)
            elif before:
                starts[i] = before[-1][1]
            else:
                starts[i] = after[0][1]
    # Jaga agar naik terus.
    for i in range(1, len(starts)):
        if starts[i] is not None and starts[i - 1] is not None and starts[i] < starts[i - 1]:
            starts[i] = starts[i - 1]
    ratio = matched / max(1, len(doc_words))
    return starts, ratio


def cue_start_penalties(cues: list[Cue]) -> dict[float, float]:
    """Awal cue -> penalti (detik). 0 = awal kalimat, 1 = setelah koma, 2,5 = tengah frasa."""
    result = {}
    for index, cue in enumerate(cues):
        previous = cues[index - 1].text.strip() if index else "."
        if re.search(r"[.!?…][\"”’']*$", previous):
            penalty = 0.0
        elif re.search(r"[,;:—–-][\"”’']*$", previous):
            penalty = 1.0
        else:
            penalty = 2.5
        result[cue.start] = min(penalty, result.get(cue.start, penalty))
    return result


PENALTY_WEIGHT = 2.0   # detik²: harga memotong setelah koma (1x) / di tengah frasa (2,5x)


def split_points(
    t0: float,
    t1: float,
    parts: int,
    cue_starts: dict[float, float],
    weights: list[float] | None = None,
) -> list[float]:
    """Bagi [t0, t1) menjadi `parts` bagian di awal baris SRT.

    Durasi ideal tiap bagian mengikuti `weights` (panjang cue cerita tiap
    scene). Batas di awal kalimat penuh tidak dikenai biaya; setelah koma atau
    di tengah frasa dikenai biaya tambahan, sehingga awal kalimat diutamakan
    tanpa membuat satu gambar terlalu pendek/panjang.
    """
    if parts <= 1:
        return [t0, t1]
    weights = weights if weights and len(weights) == parts else [1.0] * parts
    total_weight = sum(weights) or parts
    ideals = [(t1 - t0) * w / total_weight for w in weights]
    candidates = sorted(
        (start, penalty) for start, penalty in cue_starts.items()
        if t0 + MIN_CLIP <= start <= t1 - MIN_CLIP
    )
    points = [(t0, 0.0)] + candidates
    n = len(points)
    k = parts - 1
    inf = float("inf")
    cost = [[inf] * n for _ in range(k + 1)]
    prev = [[-1] * n for _ in range(k + 1)]
    cost[0][0] = 0.0
    for j in range(1, k + 1):
        for i in range(1, n):
            extra = points[i][1] * PENALTY_WEIGHT
            for h in range(0, i):
                if cost[j - 1][h] == inf:
                    continue
                length = points[i][0] - points[h][0]
                if length < MIN_CLIP:
                    continue
                value = cost[j - 1][h] + (length - ideals[j - 1]) ** 2 + extra
                if value < cost[j][i]:
                    cost[j][i], prev[j][i] = value, h
    best, best_i = inf, -1
    for i in range(1, n):
        tail = t1 - points[i][0]
        if cost[k][i] == inf or tail < MIN_CLIP:
            continue
        value = cost[k][i] + (tail - ideals[-1]) ** 2
        if value < best:
            best, best_i = value, i
    if best_i < 0:
        bounds = [t0]
        for ideal in ideals[:-1]:
            bounds.append(bounds[-1] + ideal)
        return bounds + [t1]
    chosen = []
    i, j = best_i, k
    while j > 0:
        chosen.append(points[i][0])
        i, j = prev[j][i], j - 1
    return [t0] + sorted(chosen) + [t1]


# ---------------------------------------------------------------------------
# Scene dari Excel FlowBot
# ---------------------------------------------------------------------------

@dataclass
class Scene:
    row: int
    label: str
    ref: str
    span: tuple[int, int] | None
    images: list[str] = field(default_factory=list)
    weight: float = 1.0
    start: float = 0.0
    end: float = 0.0


def parse_span(ref: Any) -> tuple[int, int] | None:
    match = re.search(r"¶\s*(\d+)(?:\s*-\s*(\d+))?", str(ref or ""))
    if not match:
        return None
    a = int(match.group(1))
    b = int(match.group(2) or a)
    return (min(a, b), max(a, b))


def image_sort_key(name: str) -> tuple[int, str]:
    match = re.search(r"-(\d+)\.[a-z]+$", name, re.IGNORECASE)
    return (int(match.group(1)) if match else 0, name)


def load_scenes(excel: Path, config: dict[str, Any], image_dir: Path) -> list[Scene]:
    file_config = bot.config_for_input(config, excel)
    headers = bot.read_headers(excel, config.get("sheet_name"))
    ref_header = bot.find_header(headers, bot.SOURCE_REF_HEADERS)
    prompt_column = str(file_config.get("prompt", {}).get("column", "1"))
    base = file_config["download"]["base_name"]
    on_disk = {}
    if image_dir.is_dir():
        for path in image_dir.iterdir():
            if path.suffix.casefold() in IMAGE_EXT:
                on_disk.setdefault(path.stem.casefold(), path.name)
    disk_names = set(on_disk.values())

    scenes = []
    for number, row in bot.read_rows(excel, config.get("sheet_name")):
        if bot.is_empty(row.get(prompt_column)):
            continue
        label = bot.scene_label(number, row, file_config)
        ref = str(row.get(ref_header) or "") if ref_header else ""
        listed = []
        for key, value in row.items():
            if str(key).casefold() == bot.MARK_FILES.casefold() and value:
                listed = [line.strip() for line in str(value).splitlines() if line.strip()]
        # Hanya file yang benar-benar ada (jika foldernya ada); tambah file di
        # folder yang cocok dengan scene walaupun belum tercatat di Excel.
        images = [name for name in listed if not disk_names or name in disk_names]
        prefix = f"{base} {label}-".casefold()
        for stem, name in on_disk.items():
            if stem.startswith(prefix) and name not in images:
                images.append(name)
        images.sort(key=image_sort_key)
        # Bobot durasi = panjang cue cerita (“...”) di prompt, bila ada.
        cue = re.search(r"[“\"]([^”\"]{8,})[”\"]", str(row.get(prompt_column) or ""))
        weight = float(len(cue.group(1))) if cue else 0.0
        scenes.append(Scene(number, label, ref, parse_span(ref), images, weight))
    known = [scene.weight for scene in scenes if scene.weight > 0]
    average = sum(known) / len(known) if known else 1.0
    for scene in scenes:
        if scene.weight <= 0:
            scene.weight = average
    return scenes


# ---------------------------------------------------------------------------
# Rencana timeline
# ---------------------------------------------------------------------------

@dataclass
class Clip:
    name: str
    start: float
    end: float
    scene: Scene
    note: str = ""


def assign_scene_times(scenes: list[Scene], starts: list[float | None], cues: list[Cue]) -> list[str]:
    """Isi scene.start/end. Kembalikan daftar peringatan."""
    warnings = []
    total_end = cues[-1].end if cues else 0.0
    penalties = cue_start_penalties(cues)

    # Kelompokkan scene beruntun yang memakai ¶ yang sama. Scene tanpa ¶
    # (mis. ¶soulsearcha) ikut kelompok sebelumnya.
    # Scene tanpa nomor ¶ (mis. ¶soulsearcha) ikut kelompok sebelumnya:
    # waktunya dibagi dari sisa paragraf itu, dipotong di awal kalimat.
    groups: list[tuple[tuple[int, int] | None, list[Scene]]] = []
    for scene in scenes:
        if scene.span is None or scene.span[0] >= len(starts):
            if scene.span is not None:
                warnings.append(f"{scene.label}: {scene.ref} melebihi jumlah paragraf Word")
            elif scene.ref:
                warnings.append(f"{scene.label}: Source Ref '{scene.ref}' tanpa nomor ¶, digabung ke paragraf sebelumnya")
            if groups:
                groups[-1][1].append(scene)
            else:
                groups.append((None, [scene]))
            continue
        if groups and groups[-1][0] == scene.span:
            groups[-1][1].append(scene)
        else:
            groups.append((scene.span, [scene]))

    anchors = []
    for span, _ in groups:
        anchors.append(starts[span[0]] if span is not None else None)
    # Anchor pertama = 0; anchor mundur (urutan ¶ tidak naik) digabung.
    cleaned = []
    last = 0.0
    for index, anchor in enumerate(anchors):
        if index == 0:
            cleaned.append(0.0)
            continue
        if anchor is None or anchor < last + MIN_CLIP:
            cleaned.append(None)
            continue
        cleaned.append(anchor)
        last = anchor

    # Kelompok tanpa anchor valid digabung ke kelompok sebelumnya.
    merged: list[tuple[float, list[Scene]]] = []
    for anchor, (_, members) in zip(cleaned, groups):
        if anchor is None and merged:
            merged[-1][1].extend(members)
        else:
            merged.append((anchor if anchor is not None else 0.0, list(members)))

    for index, (t0, members) in enumerate(merged):
        t1 = merged[index + 1][0] if index + 1 < len(merged) else total_end
        points = split_points(t0, t1, len(members), penalties, [m.weight for m in members])
        for k, scene in enumerate(members):
            scene.start, scene.end = points[k], points[k + 1]
    return warnings


def build_clips(scenes: list[Scene], cues: list[Cue]) -> tuple[list[Clip], list[str]]:
    warnings = []
    sentence_breaks = {t: p for t, p in cue_start_penalties(cues).items() if p <= 1.0}
    clips: list[Clip] = []
    pending_start: float | None = None   # waktu scene tanpa gambar di awal
    for scene in scenes:
        duration = scene.end - scene.start
        if duration <= 0:
            continue
        if not scene.images:
            warnings.append(f"{scene.label} ({scene.ref}): belum ada gambar, durasi disambung ke gambar sebelumnya")
            if clips:
                clips[-1].end = scene.end
            elif pending_start is None:
                pending_start = scene.start
            continue
        start = pending_start if pending_start is not None else scene.start
        pending_start = None
        # Gambar ke-2 hanya dipakai bila scene cukup panjang dan ada awal
        # kalimat / setelah koma untuk tempat berganti gambar.
        breaks = {
            t: p for t, p in sentence_breaks.items()
            if start + MIN_CLIP * 2 <= t <= scene.end - MIN_CLIP * 2
        }
        if len(scene.images) >= 2 and scene.end - start >= TWO_IMAGE_MIN and breaks:
            points = split_points(start, scene.end, 2, breaks)
            clips.append(Clip(scene.images[0], points[0], points[1], scene))
            clips.append(Clip(scene.images[1], points[1], points[2], scene))
        else:
            clips.append(Clip(scene.images[0], start, scene.end, scene))

    # Gabungkan klip yang terlalu pendek ke klip sebelumnya.
    result: list[Clip] = []
    for clip in clips:
        if result and clip.end - clip.start < MIN_CLIP:
            result[-1].end = clip.end
            continue
        result.append(clip)
    if len(result) >= 2 and result[0].end - result[0].start < MIN_CLIP:
        result[1].start = result[0].start
        result.pop(0)
    # Pastikan kontinu tanpa celah.
    for i in range(1, len(result)):
        result[i].start = result[i - 1].end
    if result:
        result[0].start = 0.0
    for clip in result:
        spoken = [cue.text for cue in cues if cue.start < clip.end - 0.05 and cue.end > clip.start + 0.05]
        clip.note = " ".join(spoken)
        if clip.end - clip.start > LONG_CLIP:
            warnings.append(
                f"{clip.name} tampil {clip.end - clip.start:.1f} detik — narasi {clip.scene.ref} panjang, "
                "pertimbangkan menambah scene di Excel prompt"
            )
    return result, warnings


def capcut_file_name(docx: Path | None, excel: Path) -> str:
    """Nama file mengikuti Word narasi: BTS_MODE_FINAL_Episode_041_TTS_SAFE.docx
    -> BTS_MODE_FINAL_Episode_041_CAPCUT.xlsx."""
    base = docx.stem if docx is not None else excel.stem
    base = re.sub(r"[_ -]*TTS[_ -]*SAFE$", "", base, flags=re.IGNORECASE).strip(" _-") or excel.stem
    return f"{base}_CAPCUT.xlsx"


def write_sync_file(path: Path, clips: list[Clip]) -> Path:
    """File terpisah format "Sinkron Awal Kalimat": 1 sheet Urutan CapCut,
    kolom Urutan | Nama file | Durasi (nilai waktu Excel [h]:mm:ss.000)."""
    from datetime import timedelta
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    sheet.append(["Urutan", "Nama file", "Durasi"])
    for order, clip in enumerate(clips, start=1):
        millis = int(round((clip.end - clip.start) * 1000))
        sheet.append([order, clip.name, timedelta(milliseconds=millis)])
        sheet.cell(row=order + 1, column=3).number_format = "[h]:mm:ss.000"
    sheet.column_dimensions["A"].width = 10
    sheet.column_dimensions["B"].width = 82
    sheet.column_dimensions["C"].width = 15
    sheet.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(path)
        return path
    except PermissionError:
        alternative = path.with_name(f"{path.stem}_baru{path.suffix}")
        workbook.save(alternative)
        return alternative


def write_sheet(excel: Path, clips: list[Clip], sheet_name: str | None) -> Path:
    workbook = load_workbook(excel, keep_vba=excel.suffix.casefold() == ".xlsm")
    active_title = (workbook[sheet_name] if sheet_name else workbook.active).title
    if SHEET_NAME in workbook.sheetnames:
        del workbook[SHEET_NAME]
    sheet = workbook.create_sheet(SHEET_NAME)
    headers = ["Urutan", "Nama file", "Mulai", "Selesai", "Durasi", "Catatan", "Scene", "Source Ref"]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
    for order, clip in enumerate(clips, start=1):
        sheet.append([
            order, clip.name, fmt(clip.start), fmt(clip.end), fmt(clip.end - clip.start),
            clip.note[:500], clip.scene.label, clip.scene.ref,
        ])
    for letter, width in zip("ABCDEFGH", (8, 30, 14, 14, 14, 80, 9, 14)):
        sheet.column_dimensions[letter].width = width
    for row in sheet.iter_rows(min_row=2):
        row[5].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "A2"
    # Sheet prompt tetap menjadi sheet aktif agar FlowBot membaca sheet yang benar.
    workbook.active = workbook.sheetnames.index(active_title)
    temp = excel.with_name(excel.name + ".tmp")
    try:
        workbook.save(temp)
        os.replace(temp, excel)
        return excel
    except PermissionError:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass
        fallback = excel.with_name(f"{excel.stem}_CAPCUT{excel.suffix}")
        workbook.save(fallback)
        return fallback


def find_srt(excel: Path) -> Path | None:
    folder = excel.parent
    srts = [path for path in folder.glob("*.srt")]
    stem = excel.stem.casefold()
    matches = [path for path in srts if stem in path.stem.casefold()]
    if not matches:
        episode = re.search(r"(?:episode|eps?|ep)[ _-]*0*(\d+)", stem)
        if episode:
            number = episode.group(1)
            matches = [
                path for path in srts
                if re.search(rf"(?:episode|eps?|ep)[ _-]*0*{number}(?!\d)", path.stem.casefold())
            ]
    if not matches:
        return None
    return max(matches, key=lambda path: path.stat().st_mtime)


def make_plan(
    excel: Path,
    config: dict[str, Any],
    srt: Path | None = None,
    docx: Path | None = None,
    image_dir: Path | None = None,
    write: bool = True,
) -> dict[str, Any]:
    srt = srt or find_srt(excel)
    if srt is None:
        raise RuntimeError(f"File SRT untuk {excel.name} tidak ditemukan di folder {excel.parent}")
    docx = docx or bot.find_narration_docx(excel)
    if docx is None:
        raise RuntimeError(f"File Word narasi untuk {excel.name} tidak ditemukan di folder {excel.parent}")
    if image_dir is None:
        image_dir = bot.resolve_path(bot.config_for_input(config, excel)["download"]["folder"])

    cues = read_srt(srt)
    if not cues:
        raise RuntimeError(f"SRT kosong atau tidak terbaca: {srt.name}")
    paragraphs = bot.docx_paragraphs(docx)
    starts, ratio = paragraph_starts(paragraphs, cues)
    scenes = load_scenes(excel, config, image_dir)
    warnings = assign_scene_times(scenes, starts, cues)
    clips, clip_warnings = build_clips(scenes, cues)
    warnings += clip_warnings
    if ratio < 0.9:
        warnings.insert(0, f"Hanya {ratio:.0%} kata Word yang cocok dengan SRT — pastikan SRT dari Word yang sama")
    output = None
    if write and clips:
        output = write_sheet(excel, clips, config.get("sheet_name"))
        # File siap pakai untuk CapCut Studio Sync, disimpan di folder gambar.
        sync_file = write_sync_file(image_dir / capcut_file_name(docx, excel), clips)
    used = {clip.name for clip in clips}
    all_images = {name for scene in scenes for name in scene.images}
    return {
        "srt": srt, "docx": docx, "image_dir": image_dir, "match_ratio": ratio,
        "scenes": scenes, "clips": clips, "warnings": warnings, "output": output,
        "sync_file": sync_file if write and clips else None,
        "total": cues[-1].end, "unused": sorted(all_images - used, key=image_sort_key),
    }


def summary_text(result: dict[str, Any]) -> str:
    clips = result["clips"]
    with_images = sum(1 for scene in result["scenes"] if scene.images)
    lines = [
        f"SRT      : {result['srt'].name}",
        f"Word     : {result['docx'].name} (kata cocok {result['match_ratio']:.1%})",
        f"Gambar   : {result['image_dir']}",
        f"Scene    : {len(result['scenes'])} (ada gambar {with_images})",
        f"Klip     : {len(clips)} gambar, total {fmt(result['total'])}",
    ]
    if result["output"]:
        lines.append(f"Disimpan : sheet '{SHEET_NAME}' di {result['output'].name}")
    if result.get("sync_file"):
        lines.append(f"CapCut   : {result['sync_file']}")
    elif not clips:
        lines.append("Belum ada gambar hasil generate, sheet tidak dibuat.")
    if result["unused"]:
        lines.append(f"Tidak dipakai: {len(result['unused'])} gambar (scene pendek cukup 1 gambar)")
    if result["warnings"]:
        lines.append(f"Peringatan ({len(result['warnings'])}):")
        lines += [f"  - {w}" for w in result["warnings"][:40]]
        if len(result["warnings"]) > 40:
            lines.append(f"  ... dan {len(result['warnings']) - 40} lainnya")
    return "\n".join(lines)


def main() -> int:
    import json

    parser = argparse.ArgumentParser(description="Buat sheet Urutan CapCut dari Excel FlowBot + Word + SRT")
    parser.add_argument("excel")
    parser.add_argument("--srt")
    parser.add_argument("--docx")
    parser.add_argument("--images", help="Folder gambar (default downloads/<nama Excel>)")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--check", action="store_true", help="Hanya periksa, tidak menulis sheet")
    args = parser.parse_args()
    config = json.loads(bot.resolve_path(args.config).read_text(encoding="utf-8"))
    result = make_plan(
        Path(args.excel).resolve(), config,
        Path(args.srt) if args.srt else None,
        Path(args.docx) if args.docx else None,
        Path(args.images) if args.images else None,
        write=not args.check,
    )
    print(summary_text(result))
    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    raise SystemExit(main())
