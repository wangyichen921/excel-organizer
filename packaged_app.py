"""PyInstaller 入口；保留 app.py 作为普通 Python 启动方式。"""

import sys

from app import OrganizerApp
from organizer import merge_workbooks


def main() -> None:
    # 诊断模式让无控制台的 EXE 也能接受自动化端到端检查。
    if len(sys.argv) == 6 and sys.argv[1] == "--smoke-test":
        window = OrganizerApp()
        window.withdraw()
        window.update()
        window.destroy()
        merge_workbooks([sys.argv[2], sys.argv[3]], sys.argv[4], sys.argv[5])
        return
    if len(sys.argv) != 1:
        raise SystemExit("用法：EXE [--smoke-test 输入1.xlsx 输入2.xlsx 输出.xlsx 排序列]")
    OrganizerApp().mainloop()


if __name__ == "__main__":
    main()
