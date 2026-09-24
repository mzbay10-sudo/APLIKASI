from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def _run_external_launcher() -> None:
    """EXE menjalankan launcher.py di samping EXE (bila ada) agar tampilan
    aplikasi bisa diperbarui tanpa build ulang. Gagal -> pakai versi bawaan."""
    if (not getattr(sys, "frozen", False) or globals().get("__FLOWBOT_EXTERNAL__")
            or getattr(sys, "_flowbot_external", False)):
        return
    path = APP_DIR / "launcher.py"
    if not path.is_file():
        return
    try:
        code = compile(path.read_text(encoding="utf-8"), str(path), "exec")
    except Exception as exc:
        print(f"INFO: launcher.py eksternal tidak valid ({exc}); memakai versi bawaan EXE.")
        return
    sys._flowbot_external = True  # type: ignore[attr-defined]
    namespace = {"__name__": "__main__", "__file__": str(path), "__FLOWBOT_EXTERNAL__": True}
    exec(code, namespace)
    raise SystemExit(0)


_run_external_launcher()


def _load_external_module(name: str) -> None:
    """Pada EXE, pakai bot.py/capcut_plan.py di samping EXE bila ada.

    Dengan begitu perbaikan kode cukup mengganti file .py tanpa build ulang.
    Jika file eksternal gagal dimuat, modul bawaan EXE tetap dipakai.
    """
    if not getattr(sys, "frozen", False):
        return
    path = APP_DIR / f"{name}.py"
    if not path.is_file():
        return
    import importlib.util

    try:
        spec = importlib.util.spec_from_file_location(name, path)
        if spec is None or spec.loader is None:
            return
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
    except Exception as exc:  # pragma: no cover - fallback ke modul bawaan
        sys.modules.pop(name, None)
        print(f"INFO: {path.name} eksternal gagal dimuat ({exc}); memakai versi bawaan EXE.")


_load_external_module("bot")
_load_external_module("capcut_plan")

import bot  # noqa: E402
import capcut_plan  # noqa: E402


def reload_modules() -> None:
    """Muat ulang bot.py/capcut_plan.py terbaru (tanpa menutup aplikasi)."""
    global bot, capcut_plan
    import importlib

    for name in ("bot", "capcut_plan"):
        if getattr(sys, "frozen", False):
            _load_external_module(name)
        else:
            try:
                importlib.reload(sys.modules[name])
            except Exception as exc:
                print(f"INFO: {name}.py gagal dimuat ulang ({exc})")
    bot = sys.modules["bot"]
    capcut_plan = sys.modules["capcut_plan"]

PROFILES_FILE = APP_DIR / "profiles.json"
CONFIG_FILE = APP_DIR / "config.json"
PROJECT_URLS_FILE = APP_DIR / "profile_urls.json"
ROTATION_FILE = APP_DIR / "profile_rotation.json"
PROJECT_URL_PATTERN = re.compile(r"^https://[^\s/]+/(?:[^\s?#]*/)?project/[A-Za-z0-9-]{8,}", re.IGNORECASE)


def load_project_urls(profiles: dict[str, str]) -> dict[str, str]:
    """URL project Flow per profil. Project Flow hanya bisa dibuka akun pemiliknya."""
    urls: dict[str, str] = {}
    if PROJECT_URLS_FILE.exists():
        try:
            urls = json.loads(PROJECT_URLS_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            urls = {}
    if not urls:
        # Migrasi: URL lama di config.json milik profil yang memakai profile_dir config.
        try:
            config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            owner_dir = str(config.get("profile_dir", "runtime/browser-profile")).replace("\\", "/")
            for name, path in profiles.items():
                if str(path).replace("\\", "/") == owner_dir and config.get("url"):
                    urls[name] = config["url"]
            if urls:
                PROJECT_URLS_FILE.write_text(json.dumps(urls, indent=2), encoding="utf-8")
        except (OSError, ValueError):
            pass
    return urls


def bot_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def load_profiles() -> dict[str, str]:
    if PROFILES_FILE.exists():
        return json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
    profiles = {"Renegade Immortal": "runtime/browser-profile"}
    PROFILES_FILE.write_text(json.dumps(profiles, indent=2), encoding="utf-8")
    return profiles


def worker_command(extra: list[str]) -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--worker", *extra]
    return [sys.executable, str(Path(__file__).resolve()), "--worker", *extra]



class FlowBotApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Flow Bot v8 - Generate Gambar")
        self.geometry("980x720")
        self.minsize(860, 620)
        self.profiles = load_profiles()
        self.project_urls = load_project_urls(self.profiles)
        self.project_var = tk.StringVar()
        self.files: list[Path] = []
        self.process: subprocess.Popen[str] | None = None
        self.job = ""
        self.job_info: dict[str, str] = {}
        if not self.profiles:
            self.profiles = {"Profil 1": "runtime/browser-profile"}
            PROFILES_FILE.write_text(json.dumps(self.profiles, indent=2), encoding="utf-8")
        names = list(self.profiles)
        self.profile_var = tk.StringVar(value=names[0])
        self.summary_var = tk.StringVar(value="Belum ada Excel dipilih")
        self.status_var = tk.StringVar(value="Siap")
        self.rotate_var = tk.BooleanVar(value=False)
        self.rotation: list[str] = []
        self.rotation_index = 0
        self.pending_launch: str | None = None
        self.retry_done = False
        self.progress_start: tuple[float, int] | None = None
        self.output: queue.Queue = queue.Queue()
        self.after(100, self._poll_output)
        self._build()
        self.load_data_files()

    # ------------------------------------------------------------------ UI
    def _build(self) -> None:
        style = ttk.Style(self)
        try:
            style.configure("Big.TButton", padding=(10, 6))
            style.configure("Head.TLabel", font=("Segoe UI", 10, "bold"))
        except tk.TclError:
            pass
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        profile_box = ttk.LabelFrame(outer, text="Profil Google / Chrome", padding=10)
        profile_box.pack(fill="x")
        ttk.Label(profile_box, text="Profil:").pack(side="left")
        self.profile_combo = ttk.Combobox(profile_box, textvariable=self.profile_var, state="readonly", width=28)
        self.profile_combo.pack(side="left", padx=(4, 8))
        ttk.Button(profile_box, text="Tambah profil", command=self.add_profile).pack(side="left", padx=3)
        ttk.Button(profile_box, text="Login profil", command=self.login_profile).pack(side="left", padx=3)
        ttk.Button(profile_box, text="URL project", command=self.set_project_url).pack(side="left", padx=3)
        ttk.Label(profile_box, textvariable=self.project_var, foreground="#555").pack(side="left", padx=8)
        self.profile_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_project_label())

        gen = ttk.LabelFrame(outer, text="Generate Gambar", padding=10)
        gen.pack(fill="x", pady=(10, 0))
        self._build_generate(gen)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(8, 0))
        self.stop_button = ttk.Button(bottom, text="Hentikan", command=self.stop, state="disabled")
        self.stop_button.pack(side="left")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left", padx=10)

        log_box = ttk.LabelFrame(outer, text="Progres", padding=8)
        log_box.pack(fill="both", expand=True, pady=(6, 0))
        self.log = tk.Text(log_box, wrap="word", state="disabled", font=("Consolas", 10), height=10)
        scroll = ttk.Scrollbar(log_box, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.refresh_profile_lists()

    def _build_generate(self, parent: ttk.Frame) -> None:
        ttk.Label(
            parent, foreground="#555", wraplength=900, justify="left",
            text="Generate memakai profil yang dipilih di atas. Karakter di Excel yang belum ada di profil ini "
                 "otomatis dicari di profil lain lalu dipindahkan sebelum generate (tutup Chrome profil lain).",
        ).pack(anchor="w")
        buttons = ttk.Frame(parent)
        buttons.pack(fill="x", pady=(6, 0))
        ttk.Button(buttons, text="Tambah Excel", command=self.add_files).pack(side="left", padx=(0, 5))
        ttk.Button(buttons, text="Hapus pilihan", command=self.remove_files).pack(side="left", padx=5)
        ttk.Button(buttons, text="Naik", command=lambda: self.move_file(-1)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Turun", command=lambda: self.move_file(1)).pack(side="left", padx=5)
        self.file_list = tk.Listbox(parent, height=7, selectmode="extended")
        self.file_list.pack(fill="x", pady=(8, 4))
        ttk.Label(parent, textvariable=self.summary_var).pack(anchor="w")
        rotate_box = ttk.LabelFrame(parent, text="Pindah profil otomatis", padding=6)
        rotate_box.pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(
            rotate_box, variable=self.rotate_var, command=self.save_rotation,
            text="Aktif. Nano Banana 2 → (kredit habis/pembatasan) → Nano Banana 2 Lite → (gagal 2x berturut-turut) "
                 "→ profil berikutnya. Semua profil gagal → bot berhenti",
        ).pack(anchor="w")
        row = ttk.Frame(rotate_box)
        row.pack(fill="x", pady=(4, 0))
        ttk.Label(row, text="Profil bergantian (klik untuk pilih/lepas).\nMulai dari profil di atas,\nlalu sesuai urutan daftar:",
                  foreground="#555").pack(side="left", anchor="n")
        self.rotation_list = tk.Listbox(row, height=4, selectmode="multiple", exportselection=False)
        self.rotation_list.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.rotation_list.bind("<<ListboxSelect>>", lambda _e: self.save_rotation())
        run_box = ttk.Frame(parent)
        run_box.pack(fill="x", pady=(8, 0))
        self.start_button = ttk.Button(run_box, text="▶  Mulai generate", style="Big.TButton", command=self.start)
        self.start_button.pack(side="left")
        self.capcut_button = ttk.Button(run_box, text="Buat Urutan CapCut", style="Big.TButton", command=self.make_capcut)
        self.capcut_button.pack(side="left", padx=8)
        ttk.Button(run_box, text="Buka folder hasil", command=lambda: self.open_folder(APP_DIR / "downloads")).pack(side="left")
        self.clean_button = ttk.Button(run_box, text="Hapus karakter tanpa nama", command=self.clean_untitled)
        self.clean_button.pack(side="right")

    def refresh_profile_lists(self) -> None:
        names = list(self.profiles)
        self.profile_combo["values"] = names
        self.refresh_project_label()
        self.load_rotation()

    def load_rotation(self) -> None:
        saved: dict = {}
        if ROTATION_FILE.exists():
            try:
                saved = json.loads(ROTATION_FILE.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                saved = {}
        chosen = saved.get("profiles") if "profiles" in saved else None  # belum pernah disimpan -> semua
        self.rotate_var.set(bool(saved.get("enabled", False)))
        last = saved.get("last_profile")
        if last in self.profiles and not getattr(self, "_last_restored", False):
            # Mulai dari profil terakhir yang dipakai (berkelanjutan, tidak mengulang dari awal).
            self._last_restored = True
            self.profile_var.set(last)
            self.refresh_project_label()
        self.rotation_list.delete(0, "end")
        for index, name in enumerate(self.profiles):
            self.rotation_list.insert("end", name)
            if chosen is None or name in chosen:
                self.rotation_list.selection_set(index)
        if (last not in self.profiles and chosen and self.rotate_var.get()
                and not getattr(self, "_last_restored", False)):
            # Belum ada riwayat: mulai dari profil pertama yang dipilih untuk rotasi.
            first = next((n for n in self.profiles if n in chosen), None)
            if first:
                self._last_restored = True
                self.profile_var.set(first)
                self.refresh_project_label()

    def save_rotation(self, last_profile: str | None = None) -> None:
        previous: dict = {}
        if ROTATION_FILE.exists():
            try:
                previous = json.loads(ROTATION_FILE.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                previous = {}
        data = {
            "enabled": bool(self.rotate_var.get()),
            "profiles": [self.rotation_list.get(i) for i in self.rotation_list.curselection()],
            "last_profile": last_profile or previous.get("last_profile"),
        }
        try:
            ROTATION_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass

    def set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for button in (self.start_button, self.capcut_button, self.clean_button):
            button.configure(state=state)
        self.stop_button.configure(state="normal" if busy else "disabled")

    @staticmethod
    def open_folder(path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        if hasattr(os, "startfile"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(path)])

    # ------------------------------------------------------------ profiles
    def add_profile(self) -> None:
        name = simpledialog.askstring("Tambah profil", "Nama profil (contoh: Google Kedua):", parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        slug = re.sub(r"[^a-z0-9]+", "-", name.casefold()).strip("-") or "profile"
        path = f"runtime/profiles/{slug}"
        self.profiles[name] = path
        PROFILES_FILE.write_text(json.dumps(self.profiles, indent=2), encoding="utf-8")
        self.profile_var.set(name)
        chosen = [self.rotation_list.get(i) for i in self.rotation_list.curselection()]
        self.refresh_profile_lists()
        for index, item in enumerate(self.profiles):
            if item in chosen or item == name:
                self.rotation_list.selection_set(index)
        self.save_rotation()
        messagebox.showinfo(
            "Profil ditambahkan",
            "Klik 'Login profil' untuk login Google pada profil baru, lalu klik 'URL project' "
            "dan tempel URL project Flow milik akun tersebut.",
        )

    def refresh_project_label(self) -> None:
        url = self.project_urls.get(self.profile_var.get())
        if url:
            project_id = url.rstrip("/").rsplit("/", 1)[-1]
            self.project_var.set(f"Project: …{project_id[-12:]}")
        else:
            self.project_var.set("Project: BELUM DIATUR")

    def set_project_url(self, name: str | None = None) -> bool:
        """Minta URL project Flow milik akun profil lalu simpan."""
        name = name or self.profile_var.get()
        value = simpledialog.askstring(
            "URL project Google Flow",
            f"Profil: {name}\n\n"
            "1. Klik 'Login profil', buka Google Flow dengan akun profil ini.\n"
            "2. Buka project yang berisi karakter referensinya.\n"
            "3. Salin URL dari address bar lalu tempel di sini.\n\n"
            "Contoh: https://labs.google/fx/id/tools/flow/project/xxxxxxxx-xxxx-...",
            initialvalue=self.project_urls.get(name, ""),
            parent=self,
        )
        if value is None:
            return False
        match = PROJECT_URL_PATTERN.match(value.strip())
        if not match:
            messagebox.showerror(
                "URL tidak valid",
                "URL harus berupa alamat project Flow, contoh:\n"
                "https://labs.google/fx/id/tools/flow/project/xxxxxxxx-...",
            )
            return False
        self.project_urls[name] = match.group(0)
        PROJECT_URLS_FILE.write_text(json.dumps(self.project_urls, indent=2), encoding="utf-8")
        self.refresh_project_label()
        return True

    def ensure_url(self, name: str) -> str | None:
        url = self.project_urls.get(name)
        if url:
            return url
        messagebox.showinfo(
            "URL project belum diatur",
            f"Profil '{name}' belum punya URL project Flow. "
            "Setiap akun Google punya project sendiri, jadi URL-nya harus diatur per profil.",
        )
        if not self.set_project_url(name):
            return None
        return self.project_urls.get(name)

    def profile_path(self, name: str | None = None) -> Path:
        path = Path(self.profiles[name or self.profile_var.get()])
        return path if path.is_absolute() else APP_DIR / path

    def login_profile(self) -> None:
        candidates = [
            Path(os.environ.get("PROGRAMFILES", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
        ]
        chrome = next((path for path in candidates if path.is_file()), None)
        if chrome is None:
            messagebox.showerror("Chrome tidak ditemukan", "Instal Google Chrome biasa terlebih dahulu.")
            return
        profile = self.profile_path()
        profile.mkdir(parents=True, exist_ok=True)
        subprocess.Popen([str(chrome), f"--user-data-dir={profile}", "--no-first-run", "https://labs.google/fx/tools/flow"])
        messagebox.showinfo("Login profil", "Login di Chrome yang terbuka, lalu tutup seluruh jendela profil tersebut sebelum menjalankan bot.")

    # ---------------------------------------------------------- generate
    def load_data_files(self) -> None:
        """Muat otomatis semua Excel/CSV dari folder data saat aplikasi dibuka."""
        data_dir = APP_DIR / "data"
        if not data_dir.is_dir():
            return
        supported = {".xlsx", ".xlsm", ".csv"}
        self.files = sorted(
            (path for path in data_dir.iterdir()
             if path.is_file() and path.suffix.casefold() in supported and not path.name.startswith("~$")),
            key=lambda path: path.name.casefold(),
        )
        self.refresh_files()

    def add_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Pilih satu atau beberapa Excel/CSV",
            filetypes=[("Excel dan CSV", "*.xlsx *.xlsm *.csv"), ("Semua file", "*.*")],
        )
        for value in selected:
            path = Path(value)
            if path not in self.files:
                self.files.append(path)
        self.refresh_files()

    def remove_files(self) -> None:
        selected = set(self.file_list.curselection())
        self.files = [path for index, path in enumerate(self.files) if index not in selected]
        self.refresh_files()

    def move_file(self, direction: int) -> None:
        selected = self.file_list.curselection()
        if len(selected) != 1:
            return
        old = selected[0]
        new = max(0, min(len(self.files) - 1, old + direction))
        if old == new:
            return
        self.files[old], self.files[new] = self.files[new], self.files[old]
        self.refresh_files()
        self.file_list.selection_set(new)

    def refresh_files(self) -> None:
        self.file_list.delete(0, "end")
        total = 0
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        for index, path in enumerate(self.files, start=1):
            try:
                count = len(bot.prompt_rows(path, config))
                total += count
                self.file_list.insert("end", f"{index}. {path.name} — {count} prompt belum selesai")
            except Exception as exc:
                self.file_list.insert("end", f"{index}. {path.name} — ERROR: {exc}")
        self.summary_var.set(f"{len(self.files)} file • total {total} prompt belum selesai (baris bertanda SELESAI/SEBAGIAN/DITOLAK dilewati)")

    def append_log(self, value: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", value)
        # Batasi panel Progres agar aplikasi tetap ringan (log lengkap ada di runtime/logs/bot.log).
        lines = int(self.log.index("end-1c").split(".")[0])
        if lines > 4000:
            self.log.delete("1.0", f"{lines - 3000}.0")
        self.log.see("end")
        self.log.configure(state="disabled")
        if any(key in value for key in ("PROGRES ", "FILE ", "KARAKTER ", "CEK KARAKTER")):
            text = value.split("| INFO |")[-1].strip()[:140]
            eta = self.progress_eta(value)
            self.status_var.set(text + (f"  •  sisa ±{eta}" if eta else ""))

    def progress_eta(self, line: str) -> str:
        """Perkiraan sisa waktu dari baris 'PROGRES n/total' (rata-rata waktu per prompt)."""
        found = re.search(r"PROGRES (\d+)/(\d+)", line)
        if not found:
            return ""
        done, total = int(found.group(1)) - 1, int(found.group(2))
        now = time.monotonic()
        if self.progress_start is None or done < self.progress_start[1]:
            self.progress_start = (now, done)
            return ""
        started, first = self.progress_start
        if done - first < 2:
            return ""
        seconds = (now - started) / (done - first) * (total - done)
        hours, minutes = int(seconds // 3600), int(seconds % 3600 // 60)
        return f"{hours} j {minutes} m" if hours else f"{max(1, minutes)} menit"

    def run_worker(self, job: str, args: list[str], status: str) -> None:
        reload_modules()
        self.job = job
        self.set_busy(True)
        self.progress_start = None
        self.status_var.set(status)
        self.append_log(f"\n=== {status} ===\n")
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        try:
            self.process = subprocess.Popen(
                worker_command(args), cwd=APP_DIR, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", bufsize=1, env=env,
            )
        except Exception as exc:
            self.process = None
            self.job = ""
            self.set_busy(False)
            self.status_var.set("Gagal menjalankan bot")
            self.append_log(f"ERROR: bot tidak bisa dijalankan: {exc}\n")
            messagebox.showerror("Gagal", f"Bot tidak bisa dijalankan:\n{exc}")
            return
        threading.Thread(target=self._read_process, args=(self.process,), daemon=True).start()

    def start(self) -> None:
        if self.process is not None or self.pending_launch is not None or self.job:
            return
        if not self.files:
            messagebox.showwarning("Belum ada file", "Pilih minimal satu Excel atau CSV.")
            return
        reload_modules()
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        try:
            total_prompts = sum(len(bot.prompt_rows(path, config)) for path in self.files)
        except Exception as exc:
            messagebox.showerror("Excel tidak dapat dibaca", str(exc))
            return
        if total_prompts == 0:
            messagebox.showwarning(
                "Tidak ada prompt",
                "Tidak ada prompt yang perlu dikerjakan. Semua baris sudah bertanda SELESAI/SEBAGIAN/DITOLAK di kolom STATUS, "
                "atau kolom Prompt kosong. Kosongkan sel STATUS untuk mengulang scene tertentu.",
            )
            return
        if not self.confirm_excel_closed():
            return
        name = self.profile_var.get()
        url = self.ensure_url(name)
        if not url:
            return
        self.retry_done = False
        self.save_rotation()
        self.rotation = [name]
        if self.rotate_var.get():
            picked = {self.rotation_list.get(i) for i in self.rotation_list.curselection()}
            order = []
            for other in self.profiles:  # urutan daftar profil
                if other != name and other not in picked:
                    continue
                if other != name and not self.project_urls.get(other):
                    self.append_log(f"INFO: profil '{other}' dilewati dari rotasi (URL project belum diatur)\n")
                    continue
                order.append(other)
            # Mulai dari profil sekarang, lanjut ke bawah, lalu memutar ke atas: tiap profil sekali.
            start_at = order.index(name)
            self.rotation = order[start_at:] + order[:start_at]
            if len(self.rotation) > 1:
                self.append_log("ROTASI PROFIL: " + " → ".join(self.rotation) + "\n")
            else:
                self.append_log("INFO: Pindah profil otomatis aktif, tetapi belum ada profil lain yang dipilih/punya URL.\n")
        self.rotation_index = 0
        self.launch_generate()

    @staticmethod
    def excel_is_open(path: Path) -> bool:
        """Excel sedang membuka file ini? (file kunci ~$... atau file tidak bisa ditulis)."""
        if path.suffix.casefold() not in {".xlsx", ".xlsm"}:
            return False
        locks = {path.with_name("~$" + path.name), path.with_name("~$" + path.name[2:])}
        if any(lock.exists() for lock in locks):
            return True
        try:
            with path.open("r+b"):
                return False
        except PermissionError:
            return True
        except OSError:
            return False

    def confirm_excel_closed(self) -> bool:
        opened = [path.name for path in self.files if self.excel_is_open(path)]
        if not opened:
            return True
        return messagebox.askyesno(
            "Excel sedang dibuka",
            "File ini sedang dibuka di Excel:\n\n" + "\n".join(f"• {n}" for n in opened)
            + "\n\nSebaiknya tutup dulu (simpan perubahan Anda) supaya tanda STATUS langsung masuk ke file asli. "
            "Kalau tetap lanjut, tanda disimpan sementara di <nama>_TANDA.xlsx dan otomatis digabung "
            "ke file asli pada run berikutnya.\n\nLanjut tanpa menutup Excel?",
        )

    def summary_lines(self) -> tuple[str, int, int]:
        """Ringkasan status per Excel, jumlah scene yang masih bisa diulang (GAGAL/belum),
        dan jumlah scene yang sudah punya gambar (SELESAI/SEBAGIAN)."""
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        lines, retry, good = [], 0, 0
        for path in self.files:
            try:
                c = bot.status_summary(path, config)
            except Exception as exc:
                lines.append(f"{path.name}: tidak terbaca ({exc})")
                continue
            retry += c["gagal"] + c["belum"]
            good += c["selesai"] + c["sebagian"]
            lines.append(
                f"{path.name}: SELESAI {c['selesai']} • SEBAGIAN {c['sebagian']} • GAGAL {c['gagal']}"
                f" • DITOLAK {c['ditolak']} • BELUM {c['belum']}"
            )
        return "\n".join(lines), retry, good

    def launch_generate(self) -> None:
        self.pending_launch = None
        if self.rotation_index >= len(self.rotation):
            # Dibatalkan lewat tombol Hentikan saat jeda pindah profil.
            self.job = ""
            self.set_busy(False)
            self.status_var.set("Dihentikan")
            return
        name = self.rotation[self.rotation_index]
        url = self.project_urls[name]
        self.job_info = {"profile": name}
        self.profile_var.set(name)
        self.refresh_project_label()
        self.save_rotation(last_profile=name)
        args = ["--profile-dir", str(self.profile_path(name)), "--url", url, "--profile-name", name,
                "--inputs", *[str(path) for path in self.files]]
        step = f" ({self.rotation_index + 1}/{len(self.rotation)})" if len(self.rotation) > 1 else ""
        self.run_worker("generate", args, f"Generate gambar | profil {name}{step}")

    def _read_process(self, process: subprocess.Popen) -> None:
        """Thread pembaca: hanya menaruh ke antrean; UI diperbarui di thread Tk."""
        assert process.stdout is not None
        for line in process.stdout:
            self.output.put(("line", line))
        process.stdout.close()
        self.output.put(("exit", process.wait()))

    def _poll_output(self) -> None:
        try:
            for _ in range(400):
                kind, value = self.output.get_nowait()
                if kind == "line":
                    self.append_log(value)
                    if self.job == "clean":
                        found = re.search(
                            r"(?:BERSIH SELESAI \||CEK SELESAI \||karakter tanpa nama ke-)\s*(\d+)", value
                        )
                        if found:
                            self.job_info["cleaned"] = found.group(1)
                else:
                    self._finished(int(value))
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_output)

    def _finished(self, code: int) -> None:
        self.process = None
        job, self.job = self.job, ""
        if job == "clean":
            self.set_busy(False)
            count = self.job_info.get("cleaned", "0")
            profile = self.job_info.get("profile", "")
            if self.job_info.get("preview") and code != 6:
                self.status_var.set(f"Cek: {count} karakter tanpa nama di {profile}")
                messagebox.showinfo(
                    "Cek selesai",
                    f"Ditemukan {count} karakter tanpa nama di profil '{profile}'. Tidak ada yang dihapus.\n"
                    "Untuk menghapusnya, klik tombol ini lagi lalu pilih YES."
                    + ("" if code == 0 else "\nSebagian kartu belum bisa dicek; lihat panel Progres."),
                )
            elif code == 0:
                self.status_var.set(f"{count} karakter tanpa nama dihapus di {profile}")
                messagebox.showinfo("Selesai", f"{count} karakter tanpa nama dihapus di profil '{profile}'.")
            elif code == 6:
                self._project_not_found(profile)
            else:
                self.status_var.set(f"{count} karakter tanpa nama dihapus di {profile} (belum semua)")
                messagebox.showwarning(
                    "Belum selesai",
                    f"{count} karakter tanpa nama dihapus di profil '{profile}', tetapi sebagian belum terhapus "
                    "atau menu Karakter tidak terbuka. Klik tombol ini sekali lagi; detail ada di panel Progres.",
                )
            return
        self.set_busy(False)
        try:
            self.refresh_files()
        except Exception as exc:
            self.append_log(f"INFO: daftar Excel belum bisa dimuat ulang: {exc}\n")
        profile = self.job_info.get("profile", self.profile_var.get())
        if code in (6, 9, 10) and self.rotation_index + 1 < len(self.rotation):
            self.rotation_index += 1
            nxt = self.rotation[self.rotation_index]
            reason = {9: "berhenti karena kredit habis", 10: "berhenti karena pembatasan ('aktivitas tidak biasa')",
                      6: "project Flow tidak ditemukan"}[code]
            self.append_log(f"\n=== PINDAH PROFIL OTOMATIS | {profile} {reason} → lanjut dengan profil {nxt} ===\n")
            self.status_var.set(f"Pindah ke profil {nxt}...")
            self.set_busy(True)
            self.job = "generate"  # tahan tombol sampai profil berikutnya berjalan
            self.pending_launch = self.after(3000, self.launch_generate)
            return
        summary, retry, good = ("", 0, 0)
        try:
            summary, retry, good = self.summary_lines()
        except Exception as exc:
            self.append_log(f"INFO: ringkasan belum bisa dibuat: {exc}\n")
        if (code == 0 and retry and good and not self.retry_done and self.rotation_index < len(self.rotation)
                and bot_config().get("auto_retry_failed", True)):
            # Ulang otomatis SEKALI untuk scene GAGAL (error sementara), dengan profil yang sama.
            # Tidak dilakukan bila tidak ada satu pun scene berhasil (masalahnya bukan sementara).
            self.retry_done = True
            self.append_log(f"\n=== ULANG OTOMATIS | {retry} scene GAGAL/belum jadi dicoba sekali lagi ===\n")
            self.set_busy(True)
            self.job = "generate"
            self.pending_launch = self.after(3000, self.launch_generate)
            return
        if summary:
            self.append_log("\n=== RINGKASAN ===\n" + summary + "\n")
        self.status_var.set("Semua selesai" if code == 0 else f"Berhenti dengan kode {code}")
        if code == 0:
            messagebox.showinfo(
                "Selesai", "Semua file yang dipilih sudah diproses.\n\n" + summary
                + ("\n\nScene GAGAL bisa diulang dengan klik Mulai lagi." if retry else ""),
            )
        elif code == 5:
            messagebox.showerror(
                "Pergantian model gagal",
                "Bot berhenti karena reload/pergantian model (Nano Banana 2 -> Lite) gagal. "
                "Lihat panel Progres, lalu klik Mulai lagi untuk melanjutkan.",
            )
        elif code == 6:
            self._project_not_found(profile)
        elif code in (9, 10):
            many = len(self.rotation) > 1
            reason = ("kredit habis/kena pembatasan" if many
                      else "kredit habis" if code == 9 else "kena pembatasan ('aktivitas tidak biasa')")
            messagebox.showwarning(
                "Bot dihentikan",
                (f"Semua profil ({', '.join(self.rotation)}) {reason} di Nano Banana 2 dan Lite. " if many
                 else f"Profil '{profile}' {reason} di Nano Banana 2 dan Lite. ")
                + "Bot dihentikan. Klik Mulai lagi nanti: bot mulai dari profil terakhir "
                f"('{profile}') dan melanjutkan dari scene terakhir.",
            )
        else:
            messagebox.showerror("Bot berhenti", "Lihat panel Progres untuk detail kesalahan.")

    def _project_not_found(self, profile: str) -> None:
        messagebox.showerror(
            "Project Flow tidak ditemukan",
            f"Akun Google di profil '{profile}' tidak dapat membuka project:\n"
            f"{self.project_urls.get(profile, '-')}\n\n"
            "Pilih profil itu di atas, klik 'URL project' lalu tempel URL project Flow milik akun tersebut.",
        )

    def make_capcut(self) -> None:
        """Buat sheet "Urutan CapCut" dari Excel terpilih + Word narasi + SRT."""
        if self.process is not None:
            messagebox.showwarning("Bot sedang berjalan", "Tunggu bot selesai atau hentikan dulu.")
            return
        reload_modules()
        selected = self.file_list.curselection()
        if len(selected) == 1:
            excel = self.files[selected[0]]
        elif len(self.files) == 1:
            excel = self.files[0]
        else:
            messagebox.showinfo("Pilih Excel", "Klik satu Excel di daftar, lalu tekan 'Buat Urutan CapCut'.")
            return
        if excel.suffix.casefold() not in {".xlsx", ".xlsm"}:
            messagebox.showwarning("Format tidak didukung", "Urutan CapCut hanya dapat dibuat dari .xlsx.")
            return
        srt = capcut_plan.find_srt(excel)
        if srt is None:
            value = filedialog.askopenfilename(
                title=f"Pilih SRT dubbing untuk {excel.name}",
                initialdir=str(excel.parent), filetypes=[("Subtitle SRT", "*.srt")],
            )
            if not value:
                return
            srt = Path(value)
        docx = bot.find_narration_docx(excel)
        if docx is None:
            value = filedialog.askopenfilename(
                title=f"Pilih Word narasi untuk {excel.name}",
                initialdir=str(excel.parent), filetypes=[("Word", "*.docx")],
            )
            if not value:
                return
            docx = Path(value)
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        try:
            result = capcut_plan.make_plan(excel, config, srt, docx)
        except Exception as exc:
            messagebox.showerror("Urutan CapCut gagal", str(exc))
            return
        summary = capcut_plan.summary_text(result)
        self.append_log("\n=== URUTAN CAPCUT | " + excel.name + " ===\n" + summary + "\n")
        if result["output"] is None:
            messagebox.showwarning("Urutan CapCut", summary)
        else:
            messagebox.showinfo(
                "Urutan CapCut selesai",
                f"Sheet '{capcut_plan.SHEET_NAME}' dibuat di {result['output'].name}.\n"
                f"{len(result['clips'])} gambar, total {capcut_plan.fmt(result['total'])}.\n\n"
                "File CapCut ada di folder gambar episode ini. Detail ada di panel Progres.",
            )

    def stop(self) -> None:
        self.rotation_index = len(self.rotation)  # batalkan pindah profil berikutnya
        if self.pending_launch is not None:
            self.after_cancel(self.pending_launch)
            self.pending_launch = None
        if self.process is None and self.job == "generate":
            self.launch_generate()  # indeks sudah di luar rotasi -> hanya membereskan tombol
            return
        if self.process is not None:
            self.process.terminate()
            self.status_var.set("Menghentikan...")

    # -------------------------------------------------------- characters
    def clean_untitled(self) -> None:
        """Hapus karakter 'Karakter tanpa judul' (tanpa nama) di profil yang dipilih di atas."""
        if self.process is not None or self.job or self.pending_launch is not None:
            return
        name = self.profile_var.get()
        url = self.ensure_url(name)
        if not url:
            return
        answer = messagebox.askyesnocancel(
            "Hapus karakter tanpa nama",
            f"Profil '{name}': bot membuka Flow lalu memeriksa SEMUA karakter tanpa nama "
            "('Karakter tanpa judul'), sisa kegagalan bot.\n\n"
            "Karakter yang punya nama TIDAK disentuh (dicek satu per satu sebelum dihapus).\n\n"
            "YES = hapus karakter tanpa nama\n"
            "NO = cek saja (hanya dihitung, tidak ada yang dihapus)\n"
            "CANCEL = batal\n\n"
            f"Tutup dulu Chrome profil '{name}'.",
        )
        if answer is None:
            return
        preview = answer is False
        self.job_info = {"profile": name, "cleaned": "0", "preview": preview}
        args = ["--profile-dir", str(self.profile_path(name)), "--url", url, "--profile-name", name, "--clean-untitled"]
        if preview:
            args.append("--check-only")
        title = "Cek karakter tanpa nama" if preview else "Hapus karakter tanpa nama"
        self.run_worker("clean", args, f"{title} | profil {name}")


if __name__ == "__main__":
    if "--worker" in sys.argv:
        sys.argv.remove("--worker")
        raise SystemExit(bot.main())
    FlowBotApp().mainloop()
