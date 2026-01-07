import sys

from PyQt6.QtWidgets import QApplication

from main_window import MainWindow


def main() -> None:
    # QApplication のインスタンス生成
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    # イベントループ開始
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
