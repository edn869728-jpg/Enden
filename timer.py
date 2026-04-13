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
        sleep_time = max(0.0, target - elapsed)
        time.sleep(sleep_time)


def stopwatch() -> float:
    """
    碼錶：按 Enter 開始，再按 Enter 停止，回傳經過的秒數。
    """
    input("按 Enter 開始計時...")
    start = time.perf_counter()
    input("按 Enter 停止計時...")
    elapsed = time.perf_counter() - start
    return elapsed


def _default_tick(remaining: int) -> None:
    hours, rem = divmod(remaining, 3600)
    mins, secs = divmod(rem, 60)
    if hours:
        print(f"\r  ⏱  {hours:02d}:{mins:02d}:{secs:02d}", end="", flush=True)
    else:
        print(f"\r  ⏱  {mins:02d}:{secs:02d}", end="", flush=True)


def run_timer_cli() -> None:
    """互動式計時器 CLI。"""
    commands = {
        "1": "倒數計時",
        "2": "碼錶",
        "0": "返回主選單",
    }
    while True:
        print("\n=== 計時器 ===")
        for key, desc in commands.items():
            print(f"  {key}. {desc}")
        choice = input("請選擇: ").strip()

        if choice == "1":
            try:
                total = int(input("倒數秒數: "))
                if total <= 0:
                    raise ValueError
                print()
                try:
                    countdown(total)
                    print("\n  ⏰ 時間到！")
                except KeyboardInterrupt:
                    print("\n  ⏹ 計時已中止")
            except ValueError:
                print("請輸入正整數")
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
            print("無效選項，請重新輸入")
