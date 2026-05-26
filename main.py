"""
main.py — XBRL 체크리스트 실행기 (드래그앤드롭 GUI)

사용법
  1) 실행하면 창이 뜬다.
  2) 창 안 영역에 .zip 파일을 드래그해서 놓거나, '파일 선택' 버튼을 누른다.
  3) 완료되면 입력 zip 과 같은 폴더에 XBRL_CoE_Checklist_Result_<회사명>.xlsx 가 저장된다.
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:
    _HAS_DND = False
    import tkinter as TkinterDnD  # type: ignore
    DND_FILES = None               # type: ignore

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from xbrl_zip_parser import parse_xbrl_zip
from checklist_engine import run_all_checks
from standard_taxonomy import StandardTaxonomy, enrich_axis_domain_check, enrich_dart_negate_check
from checklist_export import export_checklist


APP_TITLE = "XBRL CoE 체크리스트"
DROP_HINT = (
    "여기에 .zip 파일을 드래그하세요\n"
    "(또는 아래 '파일 선택' 버튼을 누르세요)"
)

_std: StandardTaxonomy | None = None


def _resource_path(rel: str) -> str:
    """PyInstaller 로 묶었을 때도 동작하는 리소스 경로."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def _get_std() -> StandardTaxonomy:
    """StandardTaxonomy 싱글턴 — 최초 1회만 로드."""
    global _std
    if _std is None:
        _std = StandardTaxonomy()
        axis_path  = _resource_path(os.path.join("data", "Axis_Domain_Check.xlsx"))
        negate_path = _resource_path(os.path.join("data", "DART_Negate_Check.xlsx"))
        if os.path.exists(axis_path):
            enrich_axis_domain_check(_std, axis_path)
        if os.path.exists(negate_path):
            enrich_dart_negate_check(_std, negate_path)
    return _std


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        root.geometry("520x380")
        root.minsize(440, 320)

        # ── 헤더 ─────────────────────────────────────────
        ttk.Label(
            root, text=APP_TITLE,
            font=("맑은 고딕", 14, "bold"),
            padding=(12, 10, 12, 4),
        ).pack(fill="x")

        ttk.Label(
            root,
            text="DART 원문 XBRL ZIP → XBRL_CoE_Checklist_Result.xlsx",
            foreground="#555",
            padding=(12, 0, 12, 8),
        ).pack(fill="x")

        # ── 드롭 영역 ──────────────────────────────────────
        self.drop_var = tk.StringVar(value=DROP_HINT)
        self.drop = tk.Label(
            root,
            textvariable=self.drop_var,
            bg="#F4F6FB",
            fg="#334",
            relief="ridge",
            bd=2,
            font=("맑은 고딕", 11),
            wraplength=460,
            justify="center",
        )
        self.drop.pack(fill="both", expand=True, padx=14, pady=8)

        if _HAS_DND:
            self.drop.drop_target_register(DND_FILES)       # type: ignore[attr-defined]
            self.drop.dnd_bind("<<Drop>>", self.on_drop)    # type: ignore[attr-defined]

        # ── 버튼 줄 ────────────────────────────────────────
        btn_row = ttk.Frame(root, padding=(14, 4, 14, 0))
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="파일 선택…", command=self.choose_file).pack(side="left")
        ttk.Button(btn_row, text="종료", command=root.destroy).pack(side="right")

        # ── 상태 / 진행 ────────────────────────────────────
        self.status = tk.StringVar(value="대기 중")
        ttk.Label(root, textvariable=self.status, padding=(14, 8, 14, 0)).pack(fill="x")

        self.pb = ttk.Progressbar(root, mode="indeterminate")
        self.pb.pack(fill="x", padx=14, pady=(4, 12))

        self._busy = False

        # 앱 시작 시 표준 택사노미를 백그라운드로 미리 로드
        threading.Thread(target=_get_std, daemon=True).start()

    # ── 이벤트 핸들러 ───────────────────────────────────────
    def on_drop(self, event) -> None:
        if self._busy:
            return
        paths = self._parse_dnd_paths(event.data)
        zips = [p for p in paths if p.lower().endswith(".zip")]
        if not zips:
            messagebox.showwarning(APP_TITLE, ".zip 파일을 드롭해 주세요.")
            return
        threading.Thread(target=self._process_many, args=(zips,), daemon=True).start()

    def choose_file(self) -> None:
        if self._busy:
            return
        paths = filedialog.askopenfilenames(
            title="XBRL ZIP 파일 선택",
            filetypes=[("ZIP 파일", "*.zip"), ("모든 파일", "*.*")],
        )
        if not paths:
            return
        threading.Thread(target=self._process_many, args=(list(paths),), daemon=True).start()

    # ── 내부 로직 ────────────────────────────────────────────
    @staticmethod
    def _parse_dnd_paths(raw: str) -> list[str]:
        out: list[str] = []
        buf = ""
        in_brace = False
        for ch in raw:
            if ch == "{":
                in_brace = True; continue
            if ch == "}":
                in_brace = False; out.append(buf); buf = ""; continue
            if ch == " " and not in_brace:
                if buf:
                    out.append(buf); buf = ""
                continue
            buf += ch
        if buf:
            out.append(buf)
        return [p for p in (s.strip() for s in out) if p]

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if busy:
            self.pb.start(12)
        else:
            self.pb.stop()

    def _process_many(self, zip_paths: list[str]) -> None:
        self.root.after(0, self._set_busy, True)
        ok, ng = [], []
        for i, zp in enumerate(zip_paths, 1):
            name = os.path.basename(zp)
            self.root.after(0, self.status.set, f"[{i}/{len(zip_paths)}] 체크 중… {name}")
            try:
                out = self._run_one(zp)
                ok.append(out)
            except Exception as e:
                ng.append((zp, e))
                traceback.print_exc()

        self.root.after(0, self._set_busy, False)
        self.root.after(0, self._show_result, ok, ng)

    def _run_one(self, zip_path: str) -> str:
        # 1. ZIP 파싱
        with open(zip_path, "rb") as f:
            file_bytes = f.read()
        data = parse_xbrl_zip(file_bytes)
        if data.errors:
            raise RuntimeError("\n".join(data.errors))

        # 2. 체크리스트 실행
        std = _get_std()
        results = run_all_checks(data, std)

        # 3. 결과 저장 (입력 ZIP 과 같은 폴더)
        company = data.company_name or (data.entity_id or "XBRL").split("_")[0]
        in_path = Path(zip_path)
        out_path = in_path.parent / f"XBRL_CoE_Checklist_Result_{company}.xlsx"

        template_path = _resource_path(
            os.path.join("template", "XBRL_CoE_Checklist_Result.xlsx")
        )
        return export_checklist(results, template_path, str(out_path))

    def _show_result(self, ok: list[str], ng: list[tuple[str, Exception]]) -> None:
        if not ng and ok:
            self.status.set(f"완료 — {len(ok)}개 파일 저장됨")
            self.drop_var.set(
                "체크리스트 완료!\n\n"
                + "\n".join(f"• {Path(p).name}" for p in ok)
                + f"\n\n저장 위치: {Path(ok[0]).parent}"
            )
            messagebox.showinfo(
                APP_TITLE,
                "체크리스트 결과가 저장되었습니다.\n" + "\n".join(ok),
            )
        elif ng and not ok:
            self.status.set(f"실패 — {len(ng)}개 오류")
            err = "\n\n".join(f"{Path(p).name}\n  → {e}" for p, e in ng)
            messagebox.showerror(APP_TITLE, "오류가 발생했습니다.\n\n" + err)
        else:
            self.status.set(f"부분 성공 — 성공 {len(ok)}, 실패 {len(ng)}")
            messagebox.showwarning(
                APP_TITLE,
                "성공:\n" + "\n".join(ok)
                + "\n\n실패:\n" + "\n".join(f"{p}: {e}" for p, e in ng),
            )


def main() -> None:
    if _HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    App(root)
    if not _HAS_DND:
        messagebox.showinfo(
            APP_TITLE,
            "tkinterdnd2 가 설치되어 있지 않아 드래그앤드롭이 비활성화되었습니다.\n"
            "'파일 선택…' 버튼으로 zip 을 선택해 주세요.\n\n"
            "(설치: pip install tkinterdnd2)",
        )
    root.mainloop()


if __name__ == "__main__":
    main()
