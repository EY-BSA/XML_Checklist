"""
main.py — XBRL ZIP → XLSX 변환기 (드래그앤드롭 GUI)

사용법
  1) 실행하면 작은 창이 뜬다.
  2) 창 안 영역에 .zip 파일을 드래그해서 놓거나, '파일 선택' 버튼을 누른다.
  3) 변환이 끝나면 결과 xlsx 가 같은 폴더에 저장된다.
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path

# tkinterdnd2 가 있으면 사용, 없으면 일반 tkinter 만으로 동작
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:
    _HAS_DND = False
    import tkinter as TkinterDnD  # type: ignore  # 별칭만 맞춤
    DND_FILES = None  # type: ignore

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from xbrl_to_xlsx import convert_zip_to_xlsx


APP_TITLE = "XBRL ZIP → XLSX 변환기"
DROP_HINT = (
    "여기에 .zip 파일을 드래그하세요\n"
    "(또는 아래 '파일 선택' 버튼을 누르세요)"
)


def _resource_path(rel: str) -> str:
    """PyInstaller 로 묶었을 때도 동작하는 리소스 경로."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        root.geometry("520x360")
        root.minsize(440, 300)

        # ── 헤더 ───────────────────────────────────────
        header = ttk.Label(
            root, text=APP_TITLE,
            font=("맑은 고딕", 14, "bold"),
            padding=(12, 10, 12, 4),
        )
        header.pack(fill="x")

        sub = ttk.Label(
            root,
            text="DART 원문 XBRL ZIP → 다중 시트 xlsx (DataType · Period · Decimal 포함)",
            foreground="#555",
            padding=(12, 0, 12, 8),
        )
        sub.pack(fill="x")

        # ── 드롭 영역 ─────────────────────────────────
        self.drop_var = tk.StringVar(value=DROP_HINT)
        self.drop = tk.Label(
            root,
            textvariable=self.drop_var,
            bg="#F4F6FB",
            fg="#334",
            relief="ridge",
            bd=2,
            font=("맑은 고딕", 11),
            wraplength=440,
            justify="center",
        )
        self.drop.pack(fill="both", expand=True, padx=14, pady=8)

        if _HAS_DND:
            self.drop.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            self.drop.dnd_bind("<<Drop>>", self.on_drop)  # type: ignore[attr-defined]

        # ── 버튼 줄 ───────────────────────────────────
        btn_row = ttk.Frame(root, padding=(14, 4, 14, 0))
        btn_row.pack(fill="x")

        ttk.Button(btn_row, text="파일 선택…", command=self.choose_file).pack(side="left")
        ttk.Button(btn_row, text="종료", command=root.destroy).pack(side="right")

        # ── 상태 / 진행 ───────────────────────────────
        self.status = tk.StringVar(value="대기 중")
        ttk.Label(root, textvariable=self.status, padding=(14, 8, 14, 0)).pack(fill="x")

        self.pb = ttk.Progressbar(root, mode="indeterminate")
        self.pb.pack(fill="x", padx=14, pady=(4, 12))

        self._busy = False

    # ── 이벤트 핸들러 ─────────────────────────────────────
    def on_drop(self, event) -> None:
        if self._busy:
            return
        paths = self._parse_dnd_paths(event.data)
        zips = [p for p in paths if p.lower().endswith(".zip")]
        if not zips:
            messagebox.showwarning(APP_TITLE, ".zip 파일을 드롭해 주세요.")
            return
        # 여러 개 드롭해도 순차 처리
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

    # ── 내부 로직 ────────────────────────────────────────
    @staticmethod
    def _parse_dnd_paths(raw: str) -> list[str]:
        """
        tkinterdnd2 의 event.data 는 공백 구분, 경로에 공백이 있으면 {} 로 감싸짐.
        """
        out: list[str] = []
        buf = ""
        in_brace = False
        for ch in raw:
            if ch == "{":
                in_brace = True
                continue
            if ch == "}":
                in_brace = False
                out.append(buf)
                buf = ""
                continue
            if ch == " " and not in_brace:
                if buf:
                    out.append(buf)
                    buf = ""
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
            self.root.after(0, self.status.set, f"[{i}/{len(zip_paths)}] 변환 중… {name}")
            try:
                out = self._convert_one(zp)
                ok.append(out)
            except Exception as e:
                ng.append((zp, e))
                traceback.print_exc()

        self.root.after(0, self._set_busy, False)
        self.root.after(0, self._show_result, ok, ng)

    def _convert_one(self, zip_path: str) -> str:
        in_path = Path(zip_path)
        out_path = in_path.with_suffix(".xlsx")
        # 이미 같은 이름이 있으면 _converted 추가
        if out_path.exists():
            out_path = in_path.with_name(f"{in_path.stem}_converted.xlsx")
        return convert_zip_to_xlsx(str(in_path), str(out_path))

    def _show_result(self, ok: list[str], ng: list[tuple[str, Exception]]) -> None:
        if not ng and ok:
            self.status.set(f"완료 — {len(ok)}개 파일 저장됨")
            self.drop_var.set(
                "변환 완료!\n\n"
                + "\n".join(f"• {Path(p).name}" for p in ok)
                + f"\n\n저장 위치: {Path(ok[0]).parent}"
            )
            messagebox.showinfo(
                APP_TITLE,
                "변환이 완료되었습니다.\n" + "\n".join(ok),
            )
        elif ng and not ok:
            self.status.set(f"실패 — {len(ng)}개 오류")
            err = "\n\n".join(f"{Path(p).name}\n  → {e}" for p, e in ng)
            messagebox.showerror(APP_TITLE, "변환 중 오류가 발생했습니다.\n\n" + err)
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
        # Windows 에서 깔끔한 테마
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    App(root)
    if not _HAS_DND:
        # 드래그앤드롭이 없는 환경에서도 사용 가능하도록 안내
        messagebox.showinfo(
            APP_TITLE,
            "tkinterdnd2 가 설치되어 있지 않아 드래그앤드롭이 비활성화되었습니다.\n"
            "'파일 선택…' 버튼으로 zip 을 선택해 주세요.\n\n"
            "(설치: pip install tkinterdnd2)",
        )
    root.mainloop()


if __name__ == "__main__":
    main()
