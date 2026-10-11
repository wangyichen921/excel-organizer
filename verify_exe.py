"""运行打包后的 EXE，并核对输出和原始示例文件。"""

from hashlib import sha256
from pathlib import Path
import ctypes
import os
import subprocess
import time

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
EXECUTABLE = ROOT / "dist" / "Excel自动整理工具.exe"
SOURCES = [ROOT / "examples" / "订单_一.xlsx", ROOT / "examples" / "订单_二.xlsx"]
OUTPUT = ROOT / "build" / "exe_verification" / "整理结果.xlsx"


def digest(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def main() -> None:
    if not EXECUTABLE.is_file():
        raise SystemExit(f"未找到 EXE：{EXECUTABLE}")
    if not all(path.is_file() for path in SOURCES):
        raise SystemExit("缺少示例文件；请先运行 generate_examples.py。")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    before = [digest(path) for path in SOURCES]
    env = os.environ.copy()
    env.pop("PYTHONPATH", None)
    env.pop("PYTHONHOME", None)
    # 排除本机 Python 目录，检查 EXE 不依赖 PATH 中的 Python。
    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    env["PATH"] = os.pathsep.join([str(system_root / "System32"), str(system_root)])

    command = [str(EXECUTABLE), "--smoke-test", *(str(path) for path in SOURCES), str(OUTPUT), "金额"]
    completed = subprocess.run(command, env=env, timeout=120, check=False)
    if completed.returncode != 0:
        raise SystemExit(f"EXE 诊断模式退出码：{completed.returncode}")
    if [digest(path) for path in SOURCES] != before:
        raise SystemExit("原始示例文件发生变化。")

    book = load_workbook(OUTPUT, read_only=True, data_only=True)
    try:
        rows = list(book.active.values)
    finally:
        book.close()
    expected_ids = ["A002", "A001", "A003"]
    if [row[0] for row in rows[1:]] != expected_ids:
        raise SystemExit(f"输出数据或排序不正确：{rows}")
    if len(rows) != 4:
        raise SystemExit(f"输出行数不正确：{len(rows)}")
    # 再以正常模式启动，确认实际窗口可见。
    process = subprocess.Popen([str(EXECUTABLE)], env=env)
    try:
        user32 = ctypes.windll.user32
        user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
        user32.FindWindowW.restype = ctypes.c_void_p
        user32.IsWindowVisible.argtypes = [ctypes.c_void_p]
        deadline = time.monotonic() + 30
        window = 0
        while time.monotonic() < deadline:
            window = user32.FindWindowW(None, "Excel 自动整理工具")
            if window and user32.IsWindowVisible(window):
                break
            if process.poll() is not None:
                raise SystemExit(f"GUI 启动后立即退出，退出码：{process.returncode}")
            time.sleep(0.25)
        else:
            raise SystemExit("30 秒内未检测到可见的主窗口。")
    finally:
        # PyInstaller onefile 使用父子进程；结束本次测试启动的整棵进程树。
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
        )

    print(f"EXE 诊断通过：{len(rows)-1} 行数据；排序正确；原文件 SHA-256 未变化。")
    print("GUI 验证通过：正常模式下主窗口可见。")
    print(f"验证文件：{OUTPUT}")


if __name__ == "__main__":
    main()
