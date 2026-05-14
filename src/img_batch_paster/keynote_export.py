"""Convert a .pptx to .key via Keynote (macOS only).

The implementation follows the approach proven in the step-plot project:
open the pptx via `open -a Keynote` (more reliable than AppleScript's own open
command), wait for Keynote to finish importing, then save as .key.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

# step-plot 的經驗：非 ASCII（如中文）路徑會讓 osascript 不穩，
# 因此先把 pptx 複製成 ASCII 暫存檔再轉檔，最後 rename 回原檔名。
_TMP_PPTX_NAME = "_kn_intermediate.pptx"
_TMP_KEY_NAME = "_kn_output.key"

APPLESCRIPT = r'''
on run argv
    set pptxPath to item 1 of argv
    set keyPath to item 2 of argv

    do shell script "rm -rf " & quoted form of keyPath

    -- 確保 Keynote 在背景啟動 (不搶焦點)
    do shell script "open -ga Keynote"
    delay 1.5

    -- 等到 Keynote 可回應 AppleScript 指令
    set tries to 0
    repeat
        try
            tell application "Keynote" to get name
            exit repeat
        on error
            delay 0.5
            set tries to tries + 1
            if tries > 30 then error "Keynote 無回應"
        end try
    end repeat

    tell application "Keynote"
        activate
        try
            close every document saving no
        end try
    end tell
    delay 1.0

    -- open -a 比 AppleScript open 穩
    do shell script "open -a Keynote " & quoted form of pptxPath

    -- 等 import 完成
    set waited to 0
    repeat
        delay 1
        set waited to waited + 1
        try
            tell application "Keynote"
                if (count of documents) > 0 then exit repeat
            end tell
        end try
        if waited > 90 then error "Keynote 90s 內無法開啟 pptx: " & pptxPath
    end repeat
    delay 1.0

    tell application "Keynote"
        set doc to front document
        save doc in (POSIX file keyPath)
        delay 1.0
        close doc saving no
    end tell

    return "done"
end run
'''


def _remove(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def convert_pptx_to_key(pptx_path: Path, key_path: Path, timeout: int = 180) -> Path:
    if sys.platform != "darwin":
        raise RuntimeError(".key 匯出僅支援 macOS")
    if not shutil.which("osascript"):
        raise RuntimeError("找不到 osascript")

    pptx_path = pptx_path.resolve()
    key_path = key_path.resolve()
    key_path.parent.mkdir(parents=True, exist_ok=True)

    # 用 ASCII-only 暫存檔包裝，避免中文路徑導致 osascript 失敗
    work_dir = key_path.parent
    tmp_pptx = work_dir / _TMP_PPTX_NAME
    tmp_key = work_dir / _TMP_KEY_NAME

    _remove(tmp_pptx)
    _remove(tmp_key)
    shutil.copy2(pptx_path, tmp_pptx)

    try:
        # -609 (連線錯誤) 偶發；重試一次
        last_err = ""
        for attempt in range(2):
            try:
                result = subprocess.run(
                    ["osascript", "-e", APPLESCRIPT, str(tmp_pptx), str(tmp_key)],
                    capture_output=True, text=True, timeout=timeout,
                )
            except subprocess.TimeoutExpired as e:
                last_err = f"osascript timeout after {timeout}s"
                print(f"[keynote_export] attempt {attempt+1}: {last_err}", file=sys.stderr)
                _remove(tmp_key)
                time.sleep(2)
                continue
            if result.returncode == 0 and tmp_key.exists():
                break
            last_err = (result.stderr or result.stdout).strip()
            print(f"[keynote_export] attempt {attempt+1} failed:\nstderr: {result.stderr}\nstdout: {result.stdout}",
                  file=sys.stderr)
            _remove(tmp_key)
            time.sleep(2)
        else:
            raise RuntimeError(f"Keynote 轉檔失敗 (重試後): {last_err}")

        _remove(key_path)
        tmp_key.rename(key_path)
    finally:
        _remove(tmp_pptx)
        _remove(tmp_key)

    return key_path
