"""PyInstaller entry point — Flask server + native pywebview window."""
from __future__ import annotations

import socket
import sys
import threading
import time


def _find_free_port(preferred: int = 5050) -> int:
    """Try preferred port; if busy fall back to OS-assigned."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", preferred))
        s.close()
        return preferred
    except OSError:
        s.close()
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
        s.close()
        return port


def main() -> None:
    from img_batch_paster.web.app import app

    port = _find_free_port(5050)
    url = f"http://127.0.0.1:{port}/"

    # Flask in background; daemon so app exit kills it
    def _serve() -> None:
        # Disable Flask's banner; use_reloader must be False outside main thread
        app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)

    threading.Thread(target=_serve, daemon=True).start()

    # 等 server 起來
    for _ in range(40):
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.2):
                break
        except OSError:
            time.sleep(0.1)

    try:
        import webview  # type: ignore
        webview.create_window("img-batch-paster", url, width=1600, height=1000, resizable=True)
        webview.start()
    except ImportError:
        # 沒裝 pywebview 就 fallback 開瀏覽器
        import webbrowser
        webbrowser.open(url)
        # 主執行緒守著別讓 daemon thread 跟著死
        try:
            while True:
                time.sleep(60)
        except KeyboardInterrupt:
            sys.exit(0)


if __name__ == "__main__":
    main()
