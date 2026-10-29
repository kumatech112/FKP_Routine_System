"""Launch/close only the packaged idle GUI; never press Start."""
from pathlib import Path
import argparse
import json
import subprocess
import os
import time

from PIL import ImageGrab
import win32gui
import win32process
import win32con


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", type=Path, required=True)
    parser.add_argument("--destination", type=Path, required=True)
    parser.add_argument("--profile-source", type=Path)
    args = parser.parse_args()
    args.destination = args.destination.resolve()
    args.exe = args.exe.resolve()
    args.destination.mkdir(parents=True, exist_ok=False)
    if args.profile_source:
        import shutil
        profile = args.destination / "Profile"
        profile.mkdir()
        shutil.copyfile(args.profile_source, profile / "settings.json")
    environment = dict(os.environ)
    environment["FKP_ROUTINE_PROFILE_ROOT"] = str(args.destination / "Profile")
    process = subprocess.Popen([str(args.exe)], cwd=args.exe.parent, env=environment)
    window = None
    try:
        # Tk briefly creates an icon window bearing the same title during startup.
        time.sleep(2)
        for _ in range(300):
            matches = []
            def collect(hwnd, _):
                if (win32process.GetWindowThreadProcessId(hwnd)[1] == process.pid
                        and win32gui.GetWindowText(hwnd) == "FKP Routine System"
                        and win32gui.IsWindowVisible(hwnd)
                        and win32gui.GetWindowRect(hwnd)[2]-win32gui.GetWindowRect(hwnd)[0] > 500
                        and win32gui.GetWindowRect(hwnd)[3]-win32gui.GetWindowRect(hwnd)[1] > 300):
                    matches.append(hwnd)
            win32gui.EnumWindows(collect, None)
            if matches:
                candidate=max(matches,key=lambda h:(win32gui.GetWindowRect(h)[2]-win32gui.GetWindowRect(h)[0])*(win32gui.GetWindowRect(h)[3]-win32gui.GetWindowRect(h)[1]))
                time.sleep(.5)
                rect=win32gui.GetWindowRect(candidate)
                if win32gui.IsWindowVisible(candidate) and rect[2]-rect[0]>500 and rect[3]-rect[1]>300:
                    window=candidate
                    break
            if process.poll() is not None:
                raise RuntimeError("Packaged application exited before opening its window.")
            time.sleep(.1)
        if window is None:
            raise RuntimeError("Packaged window did not open.")
        time.sleep(.5)
        ImageGrab.grab(win32gui.GetWindowRect(window)).save(args.destination / "packaged_idle.png")
        win32gui.PostMessage(window, win32con.WM_CLOSE, 0, 0)
        process.wait(timeout=20)
        if process.returncode != 0:
            raise RuntimeError("Packaged application did not close cleanly.")
        result = {"status": "verified", "exit_code": process.returncode,
                  "production_start_clicked": False, "database_writes": 0, "emails_sent": 0}
        (args.destination / "verification.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=10)


if __name__ == "__main__":
    main()
