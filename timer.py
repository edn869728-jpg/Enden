"""計時器 (Timer / Stopwatch)"""

import time


def countdown(seconds: int, tick_callback=None) -> None:
    """
    倒數計時器。每秒呼叫一次 tick_callback(remaining)。
    預設將剩餘時間印出到終端機。使用牆上時鐘校正，避免累積誤差。
    """
    if tick_callback is None:
        tick_callback = _default_tick

    start = time.perf_counter()
    for remaining in range(seconds, -1, -1):
        tick_callback(remaining)
        if remaining == 0:
            break
        elapsed = time.perf_counter() - start
        target = seconds - remaining + 1
        time.sleep(max(0.0, target - elapsed))


def stopwatch() -> float:
    """
    碼錶：按 Enter 開始，再按 Enter 停止，回傳經過的秒數。
    """
    input("按 Enter 開始計時...")
    start = time.perf_counter()
    input("按 Enter 停止計時...")
    return time.perf_counter() - start


def _fmt_time(total_seconds) -> str:
    """將秒數格式化為 MM:SS 或 HH:MM:SS 字串。"""
    total_seconds = int(total_seconds)
    hours, rem = divmod(total_seconds, 3600)
    mins, secs = divmod(rem, 60)
    if hours:
        return f"{hours:02d}:{mins:02d}:{secs:02d}"
    return f"{mins:02d}:{secs:02d}"


def _default_tick(remaining: int) -> None:
    print(f"\r  ⏱  {_fmt_time(remaining)}", end="", flush=True)


def run_timer_cli() -> None:
    """互動式計時器 CLI。"""
    while True:
        print("\n=== 計時器 ===")
        print("  1. 倒數計時")
        print("  2. 碼錶")
        print("  0. 返回主選單")
        choice = input("請選擇: ").strip()

        if choice == "1":
            raw = input("倒數秒數: ").strip()
            try:
                total = int(raw)
                if total <= 0:
                    raise ValueError
            except ValueError:
                print("  ✘ 請輸入正整數")
                continue
            print()
            try:
                countdown(total)
                print("\n  ⏰ 時間到！")
            except KeyboardInterrupt:
                print("\n  ⏹ 計時已中止")

        elif choice == "2":
            try:
                elapsed = stopwatch()
            except KeyboardInterrupt:
                print("\n  ⏹ 碼錶已中止")
                continue
            hours, rem = divmod(elapsed, 3600)
            mins, secs = divmod(rem, 60)
            if hours:
                print(f"  經過時間: {int(hours):02d}:{int(mins):02d}:{secs:05.2f}")
            else:
                print(f"  經過時間: {int(mins):02d}:{secs:05.2f}")

        elif choice == "0":
            break

        else:
            print("  ✘ 無效選項，請輸入 0、1 或 2")
