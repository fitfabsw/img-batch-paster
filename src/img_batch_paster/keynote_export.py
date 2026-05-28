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
_TMP_KEY_IN_NAME = "_kn_input.key"
_TMP_PPTX_OUT_NAME = "_kn_converted.pptx"
_TMP_PPTX_RENDER_NAME = "_kn_render.pptx"

KEY_TO_PPTX_SCRIPT = r'''
on run argv
    set keyPath to item 1 of argv
    set pptxPath to item 2 of argv
    -- item 3 (optional)：PNG 輸出目錄；給了的話會在同一個 Keynote session 內順便匯出 slide images
    set pngDir to ""
    try
        set pngDir to item 3 of argv
    end try
    do shell script "rm -rf " & quoted form of pptxPath
    if pngDir is not "" then
        do shell script "rm -rf " & quoted form of pngDir
        do shell script "mkdir -p " & quoted form of pngDir
    end if

    do shell script "open -gja Keynote"
    delay 1.5
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
        try
            close every document saving no
        end try
    end tell
    delay 1.0

    do shell script "open -gja Keynote " & quoted form of keyPath
    set waited to 0
    repeat
        delay 1
        set waited to waited + 1
        try
            tell application "Keynote"
                if (count of documents) > 0 then exit repeat
            end tell
        end try
        if waited > 90 then error "Keynote 90s 內無法開啟 key: " & keyPath
    end repeat
    delay 0.5

    -- 開檔會把 Keynote 推到前景；立即用 System Events 隱藏起來
    try
        tell application "System Events" to set visible of process "Keynote" to false
    end try
    delay 0.3

    tell application "Keynote"
        set doc to front document
        export doc to (POSIX file pptxPath) as Microsoft PowerPoint
        delay 0.3
        if pngDir is not "" then
            export doc to (POSIX file pngDir) as slide images with properties {image format: PNG}
            delay 0.3
        end if
        close doc saving no
    end tell

    return "done"
end run
'''

APPLESCRIPT = r'''
on run argv
    set pptxPath to item 1 of argv
    set keyPath to item 2 of argv

    do shell script "rm -rf " & quoted form of keyPath

    -- 確保 Keynote 在背景啟動 (不搶焦點)
    do shell script "open -gja Keynote"
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
        try
            close every document saving no
        end try
    end tell
    delay 1.0

    -- open -a 比 AppleScript open 穩
    do shell script "open -gja Keynote " & quoted form of pptxPath

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


def convert_key_to_pptx(key_path: Path, pptx_path: Path, timeout: int = 180) -> Path:
    """Open a .key in Keynote, export as .pptx. macOS-only."""
    if sys.platform != "darwin":
        raise RuntimeError("Keynote → pptx 轉檔僅支援 macOS")
    if not shutil.which("osascript"):
        raise RuntimeError("找不到 osascript")

    key_path = key_path.resolve()
    pptx_path = pptx_path.resolve()
    pptx_path.parent.mkdir(parents=True, exist_ok=True)

    work_dir = pptx_path.parent
    tmp_key = work_dir / _TMP_KEY_IN_NAME
    tmp_pptx = work_dir / _TMP_PPTX_OUT_NAME
    png_dir = work_dir / "_kn_png_out"

    # .key 可能是 bundle (資料夾) 也可能是單檔；用 cp -R 保險
    _remove(tmp_key)
    _remove(tmp_pptx)
    shutil.rmtree(png_dir, ignore_errors=True)
    if key_path.is_dir():
        shutil.copytree(key_path, tmp_key)
    else:
        shutil.copy2(key_path, tmp_key)

    try:
        last_err = ""
        for attempt in range(2):
            try:
                result = subprocess.run(
                    ["osascript", "-e", KEY_TO_PPTX_SCRIPT, str(tmp_key), str(tmp_pptx), str(png_dir)],
                    capture_output=True, text=True, timeout=timeout,
                )
            except subprocess.TimeoutExpired:
                last_err = f"osascript timeout after {timeout}s"
                print(f"[keynote_export key→pptx] attempt {attempt+1}: {last_err}", file=sys.stderr)
                _remove(tmp_pptx)
                time.sleep(2)
                continue
            if result.returncode == 0 and tmp_pptx.exists():
                break
            last_err = (result.stderr or result.stdout).strip()
            print(f"[keynote_export key→pptx] attempt {attempt+1} failed:\nstderr: {result.stderr}\nstdout: {result.stdout}",
                  file=sys.stderr)
            _remove(tmp_pptx)
            time.sleep(2)
        else:
            raise RuntimeError(f"Keynote → pptx 轉檔失敗 (重試後): {last_err}")

        _remove(pptx_path)
        tmp_pptx.rename(pptx_path)

        # 順便把同 session 匯出的 PNG 放到 render cache、之後 render_first_slide 可直接 hit
        try:
            pngs = sorted(png_dir.glob("*.png")) if png_dir.is_dir() else []
            if pngs:
                from .web.template_render import CACHE_DIR, _cache_key
                cache_key = _cache_key(pptx_path)
                cache_png = CACHE_DIR / f"{cache_key}.png"
                CACHE_DIR.mkdir(parents=True, exist_ok=True)
                shutil.move(str(pngs[0]), str(cache_png))
        except Exception as _e:
            print(f"[keynote_export] cache PNG failed (non-fatal): {_e}", file=sys.stderr)
    finally:
        _remove(tmp_key)
        _remove(tmp_pptx)
        shutil.rmtree(png_dir, ignore_errors=True)

    return pptx_path


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


# === 用 Keynote 把 .pptx 第一張 slide 渲染成 PNG ===
PPTX_TO_PNG_SCRIPT = r'''
on run argv
    set pptxPath to item 1 of argv
    set outDir to item 2 of argv
    do shell script "rm -rf " & quoted form of outDir
    do shell script "mkdir -p " & quoted form of outDir

    do shell script "open -gja Keynote"
    delay 1.5
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

    do shell script "open -gja Keynote " & quoted form of pptxPath
    set waited to 0
    repeat
        delay 0.5
        set waited to waited + 0.5
        try
            tell application "Keynote"
                if (count of documents) > 0 then exit repeat
            end tell
        end try
        if waited > 60 then error "Keynote 60s 內無法開 pptx: " & pptxPath
    end repeat
    delay 1.0

    with timeout of 120 seconds
        tell application "Keynote"
            set doc to front document
            export doc to (POSIX file outDir) as slide images with properties {image format: PNG}
            delay 0.3
            close doc saving no
        end tell
    end timeout

    return "done"
end run
'''


def render_pptx_via_keynote(pptx_path: Path, timeout: int = 180) -> Path | None:
    """用 Keynote 把 .pptx 第一張 slide 渲染成 PNG；回傳 PNG path 或 None (失敗 / 非 macOS)。"""
    if sys.platform != "darwin":
        return None
    if not shutil.which("osascript"):
        return None

    pptx_path = Path(pptx_path).resolve()
    if not pptx_path.is_file():
        return None

    import tempfile
    # 用 ASCII 暫存檔包裝避免中文路徑導致 osascript 不穩 (沿用 step-plot 經驗)
    tmp_root = Path(tempfile.gettempdir()) / "img-batch-paster-kn-render"
    tmp_root.mkdir(parents=True, exist_ok=True)
    tmp_pptx = tmp_root / _TMP_PPTX_RENDER_NAME
    out_dir = tmp_root / "out"
    _remove(tmp_pptx)
    shutil.copy2(pptx_path, tmp_pptx)

    try:
        try:
            result = subprocess.run(
                ["osascript", "-e", PPTX_TO_PNG_SCRIPT, str(tmp_pptx), str(out_dir)],
                capture_output=True, text=True, timeout=timeout,
            )
        except subprocess.TimeoutExpired:
            print(f"[keynote_render] timeout after {timeout}s", file=sys.stderr)
            return None
        if result.returncode != 0:
            print(f"[keynote_render] failed:\nstderr: {result.stderr}\nstdout: {result.stdout}",
                  file=sys.stderr)
            return None

        pngs = sorted(out_dir.glob("*.png")) if out_dir.is_dir() else []
        if not pngs:
            return None
        return pngs[0]
    finally:
        _remove(tmp_pptx)


def prewarm_keynote() -> None:
    """背景啟動 Keynote (open -ga) 讓它一直 alive，後續渲染快很多。
    若使用者上次 Cmd+Q 退出 Keynote、會留 Saved Application State 導致 CLI 啟動立即又退。
    偵測到後清掉狀態並重試。
    """
    if sys.platform != "darwin":
        return
    saved_state = Path.home() / "Library/Containers/com.apple.iWork.Keynote/Data/Library/Saved Application State"
    try:
        subprocess.run(["open", "-ga", "Keynote"], check=False, timeout=10)
        time.sleep(2)
        if subprocess.run(["pgrep", "-x", "Keynote"], capture_output=True).returncode != 0:
            # 沒起來 → 清掉 Cmd+Q 殘留狀態、重試
            if saved_state.exists():
                shutil.rmtree(saved_state, ignore_errors=True)
            subprocess.run(["open", "-ga", "Keynote"], check=False, timeout=10)
            time.sleep(2)
    except Exception:
        pass
