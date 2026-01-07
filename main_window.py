from __future__ import annotations

from pathlib import Path
from typing import Optional
import hashlib
import tempfile
import sys

from version import __version__, __app_name__

from PyQt6.QtWebEngineCore import QWebEnginePage
from PyQt6.QtCore import Qt, QUrl, QSettings
from PyQt6.QtGui import (
    QAction,
    QKeySequence,
    QDesktopServices,
    QIcon,
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QFileDialog,
    QMessageBox,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QToolBar,
    QDialog,
    QTextEdit,
    QDialogButtonBox,
)

from transformer import XmlToStyledHtmlTransformer


class HelpReadmeDialog(QDialog):
    """README（使い方）をツール内で表示するためのシンプルなダイアログ"""

    def __init__(self, title: str, text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(800, 600)

        layout = QVBoxLayout(self)

        self.text_edit = QTextEdit(self)
        self.text_edit.setReadOnly(True)
        self.text_edit.setPlainText(text)
        layout.addWidget(self.text_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=self)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


class DroppableWebEngineView(QWebEngineView):
    """
    HTML表示エリア（QWebEngineView）でドラッグ＆ドロップを確実に受けるためのサブクラス。
    - 親Widgetに被せる方式は、WebEngine側にイベントが吸われて効かないことがあるため採用しない
    - ここで dropEvent を拾って MainWindow.open_xml_via_drop() を呼ぶ
    """

    def __init__(self, owner: "MainWindow", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._owner = owner
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        urls = [u for u in event.mimeData().urls() if u.isLocalFile()]
        if len(urls) != 1:
            QMessageBox.information(
                self,
                "ドロップエラー",
                "XMLファイルは「1つだけ」ドロップしてください。",
            )
            event.ignore()
            return

        xml_path = Path(urls[0].toLocalFile())
        self._owner.open_xml_via_drop(xml_path)
        event.acceptProposedAction()


class MainWindow(QMainWindow):
    """
    メインウィンドウ
    - 左: ファイルツリー
    - 右: Web ビュー (XML + スタイルシートを HTML に変換して表示)
    """

    ORG_NAME = "SurugaLab"
    APP_NAME = "XmlStyledViewer"

    def __init__(self, initial_xml: Optional[Path] = None) -> None:
        super().__init__()

        # ------------------------------------------------------------------
        # 基本設定
        # ------------------------------------------------------------------
        self.setWindowTitle("XML スタイルビューア")
        self.resize(1200, 800)

        self.settings = QSettings(self.ORG_NAME, self.APP_NAME)

        # 現在開いている XML ファイルパス
        self.current_xml: Optional[Path] = None

        # 現在のツリーのルートフォルダ
        self.current_folder: Optional[Path] = None

        # フォルダ遷移履歴（戻る/進む）
        self._folder_history: list[Path] = []
        self._folder_history_index: int = -1  # -1 は未設定

        # Transformer (XML + XSL → HTML)
        self.transformer = XmlToStyledHtmlTransformer()

        # UI 構築
        self._setup_ui()
        self._create_actions()
        self._create_menus()
        self._create_search_toolbar()

        # 検索用テキスト
        self._search_text: str = ""

        # ステータスバー初期表示
        self.statusBar().showMessage("準備完了")
        self._show_empty_message()


        # ★ 追加：アプリ実行中アイコン設定（ウィンドウ左上・タスクバー）
        self._apply_app_icon()

        # 初期 XML が指定されていれば読み込み
        if initial_xml is not None:
            self.open_xml(initial_xml)

    # ------------------------------------------------------------------
    # アイコン関連
    # ------------------------------------------------------------------
    def _app_icon_path(self) -> Optional[Path]:
        """
        ico の探索：
        - 開発時: main_window.py と同じフォルダ
        - exe運用時: exe と同じフォルダ
        """
        candidates: list[Path] = [
            Path(__file__).resolve().parent / "ico_xml_viewer.ico",
            Path(sys.executable).resolve().parent / "ico_xml_viewer.ico",
        ]
        for p in candidates:
            if p.exists():
                return p
        return None

    def _apply_app_icon(self) -> None:
        icon_path = self._app_icon_path()
        if icon_path is None:
            # アイコンが無くても動作はするので、警告は出さない（必要なら出してOK）
            return
        self.setWindowIcon(QIcon(str(icon_path)))

    # ------------------------------------------------------------------
    # UI 構築
    # ------------------------------------------------------------------
    def _setup_ui(self) -> None:
        """
        左にファイルツリー、右に QWebEngineView を配置する。
        ※右側（HTML表示エリア）のみ D&D を受ける
        """
        central = QWidget(self)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ------------------------------
        # 左: ファイルツリー
        # ------------------------------
        left_panel = QWidget(central)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(4, 4, 4, 4)
        left_layout.setSpacing(4)

        # フォルダ遷移ボタンバー
        nav_bar = QWidget(left_panel)
        nav_layout = QHBoxLayout(nav_bar)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(4)

        self.btn_back = QPushButton("←", nav_bar)
        self.btn_back.setToolTip("戻る")
        self.btn_back.setFixedWidth(32)
        self.btn_back.clicked.connect(self._on_folder_back)

        self.btn_forward = QPushButton("→", nav_bar)
        self.btn_forward.setToolTip("進む")
        self.btn_forward.setFixedWidth(32)
        self.btn_forward.clicked.connect(self._on_folder_forward)

        self.btn_up = QPushButton("↑", nav_bar)
        self.btn_up.setToolTip("一つ上の階層へ")
        self.btn_up.setFixedWidth(32)
        self.btn_up.clicked.connect(self._on_folder_up)

        nav_layout.addWidget(self.btn_back)
        nav_layout.addWidget(self.btn_forward)
        nav_layout.addWidget(self.btn_up)
        nav_layout.addStretch(1)

        left_layout.addWidget(nav_bar)

        # 上部に簡易ヘッダー
        header_label = QLabel("XML ファイル一覧", left_panel)
        header_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        left_layout.addWidget(header_label)

        # ファイルツリー本体
        self.tree = QTreeWidget(left_panel)
        self.tree.setHeaderHidden(True)
        self.tree.itemDoubleClicked.connect(self._on_tree_item_double_clicked)
        left_layout.addWidget(self.tree, 1)

        # 「フォルダを開く」ボタンを左下に配置
        self.open_folder_button = QPushButton("フォルダを開く...", left_panel)
        self.open_folder_button.clicked.connect(self._on_open_folder_clicked)
        left_layout.addWidget(self.open_folder_button)

        # ------------------------------
        # 右: Web ビュー（D&D 受付はこの領域だけ）
        # ------------------------------
        self.view = DroppableWebEngineView(owner=self, parent=central)

        root_layout.addWidget(left_panel, 1)
        root_layout.addWidget(self.view, 3)

        self.setCentralWidget(central)

        # 初期状態でナビボタンを無効化
        self._update_folder_nav_buttons()

    # ------------------------------------------------------------------
    # 検索バー（ページ内検索）
    # ------------------------------------------------------------------
    def _create_search_toolbar(self) -> None:
        """文書内検索用のツールバーを作成します。"""
        toolbar = QToolBar("検索", self)
        toolbar.setObjectName("SearchToolBar")

        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("文書内を検索 (Enter で次を検索)")
        self.search_bar.returnPressed.connect(self._on_search_triggered)

        toolbar.addWidget(QLabel("検索: "))
        toolbar.addWidget(self.search_bar)

        # 「次を検索」アクション
        self.search_next_action = QAction("次を検索", self)
        self.search_next_action.setShortcut(QKeySequence.StandardKey.FindNext)
        self.search_next_action.triggered.connect(self._on_search_triggered)
        toolbar.addAction(self.search_next_action)

        # 「ハイライト解除」アクション
        self.search_clear_action = QAction("ハイライト解除", self)
        self.search_clear_action.triggered.connect(self._on_search_clear)
        toolbar.addAction(self.search_clear_action)

        self.addToolBar(toolbar)

    def _on_search_triggered(self) -> None:
        """検索バーの Enter または「次を検索」ボタンが押されたときに呼ばれる。"""
        text = self.search_bar.text().strip()
        if not text:
            return

        page = self.view.page()
        if page is None:
            return

        # 検索語が変わった時だけ、以前のハイライトを消す（毎回消すと「次へ」が進まない）
        if text != self._search_text:
            page.findText("")

        self._search_text = text

        # フラグは環境差があるので使わない（最も安定する呼び方）
        page.findText(text, QWebEnginePage.FindFlag(0), self._handle_search_result)

    def _on_search_clear(self) -> None:
        """検索ハイライトをすべて解除します。"""
        page = self.view.page()
        if page is None:
            return
        page.findText("")
        self.statusBar().showMessage("検索ハイライトをクリアしました。", 3000)

    def _handle_search_result(self, found: bool) -> None:
        """findText の検索結果に応じてステータスバーを更新します。"""
        if not self._search_text:
            return

        if found:
            self.statusBar().showMessage(f"「{self._search_text}」が見つかりました。", 3000)
        else:
            self.statusBar().showMessage(f"「{self._search_text}」はこれ以上見つかりませんでした。", 3000)

    # ------------------------------------------------------------------
    # メニュー/アクション
    # ------------------------------------------------------------------
    def _create_actions(self) -> None:
        # ファイルを開く
        self.open_file_action = QAction("ファイルを開く...", self)
        self.open_file_action.setShortcut(QKeySequence.StandardKey.Open)
        self.open_file_action.triggered.connect(self._on_open_file_menu)

        # フォルダを開く
        self.open_folder_action = QAction("フォルダを開く...", self)
        self.open_folder_action.triggered.connect(self._on_open_folder_clicked)

        # HTML 保存
        self.save_html_action = QAction("HTML として保存", self)
        self.save_html_action.triggered.connect(self._on_save_html_menu)

        # PDF 保存
        self.save_pdf_action = QAction("PDF として保存", self)
        self.save_pdf_action.triggered.connect(self._on_save_pdf_menu)

        # 終了
        self.exit_action = QAction("終了", self)
        self.exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        self.exit_action.triggered.connect(self.close)

        # --- ヘルプ ---
        self.help_show_action = QAction("使い方（README）", self)
        self.help_show_action.triggered.connect(self._on_help_show_readme)

        self.help_open_readme_action = QAction("README.md を開く", self)
        self.help_open_readme_action.triggered.connect(self._on_help_open_readme)

        self.about_action = QAction("バージョン情報", self)
        self.about_action.triggered.connect(self._on_help_about)

    def _create_menus(self) -> None:
        menubar = self.menuBar()

        # ファイルメニュー（開く系）
        file_menu = menubar.addMenu("ファイル(&F)")
        file_menu.addAction(self.open_file_action)
        file_menu.addAction(self.open_folder_action)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_action)

        # 出力メニュー（保存系）
        export_menu = menubar.addMenu("出力(&E)")
        export_menu.addAction(self.save_html_action)
        export_menu.addAction(self.save_pdf_action)

        # ヘルプメニュー
        help_menu = menubar.addMenu("ヘルプ(&H)")
        help_menu.addAction(self.help_show_action)
        help_menu.addAction(self.help_open_readme_action)
        help_menu.addSeparator()
        help_menu.addAction(self.about_action)

    # ------------------------------------------------------------------
    # ヘルプ（README / About）
    # ------------------------------------------------------------------
    def _readme_path(self) -> Path:
        """
        README.md の場所：
        - 基本は main_window.py と同じフォルダを想定
        """
        return Path(__file__).resolve().parent / "README.md"

    def _load_readme_text(self) -> str:
        path = self._readme_path()
        if not path.exists():
            return (
                "README.md が見つかりませんでした。\n\n"
                f"想定パス:\n{path}\n\n"
                "README.md を main_window.py と同じフォルダに置いてください。"
            )
        # MarkdownなのでUTF-8で読み込み（社内運用なら基本これでOK）
        return path.read_text(encoding="utf-8", errors="replace")

    def _on_help_show_readme(self) -> None:
        text = self._load_readme_text()
        dlg = HelpReadmeDialog("使い方（README）", text, parent=self)
        dlg.exec()

    def _on_help_open_readme(self) -> None:
        path = self._readme_path()
        if not path.exists():
            QMessageBox.information(
                self,
                "README.md が見つかりません",
                "README.md が見つかりませんでした。\n\n"
                f"想定パス:\n{path}\n\n"
                "README.md を main_window.py と同じフォルダに置いてください。",
            )
            return

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _on_help_about(self) -> None:
        QMessageBox.about(
            self,
            "バージョン情報",
            f"{self.APP_NAME}\n"
            f"Version: {__version__}\n\n"
            "XML + XSL を HTML 表示し、必要に応じて HTML / PDF として保存できるビューアです。",
        )

    # ------------------------------------------------------------------
    # フォルダ遷移（戻る/進む/上へ）
    # ------------------------------------------------------------------
    def _update_folder_nav_buttons(self) -> None:
        can_back = self._folder_history_index > 0
        can_forward = 0 <= self._folder_history_index < (len(self._folder_history) - 1)
        can_up = self.current_folder is not None and self.current_folder.parent != self.current_folder

        self.btn_back.setEnabled(can_back)
        self.btn_forward.setEnabled(can_forward)
        self.btn_up.setEnabled(can_up)

    def _navigate_to_folder(self, folder_path: Path, push_history: bool = True) -> None:
        """指定フォルダをツリーに表示し、必要なら履歴に積む。"""
        if not folder_path.exists() or not folder_path.is_dir():
            QMessageBox.warning(self, "フォルダが存在しません", str(folder_path))
            return

        folder_path = folder_path.resolve()

        if push_history:
            # 進む履歴を削除（ブラウザと同じ挙動）
            if 0 <= self._folder_history_index < len(self._folder_history) - 1:
                self._folder_history = self._folder_history[: self._folder_history_index + 1]

            # 同一フォルダの連続pushは避ける
            if self._folder_history and self._folder_history[-1] == folder_path:
                pass
            else:
                self._folder_history.append(folder_path)
                self._folder_history_index = len(self._folder_history) - 1

        self.current_folder = folder_path
        self._populate_tree(folder_path)
        self.statusBar().showMessage(f"フォルダを表示: {folder_path}")

        self._update_folder_nav_buttons()

    def _on_folder_back(self) -> None:
        if self._folder_history_index <= 0:
            return
        self._folder_history_index -= 1
        folder = self._folder_history[self._folder_history_index]
        self._navigate_to_folder(folder, push_history=False)

    def _on_folder_forward(self) -> None:
        if self._folder_history_index < 0:
            return
        if self._folder_history_index >= len(self._folder_history) - 1:
            return
        self._folder_history_index += 1
        folder = self._folder_history[self._folder_history_index]
        self._navigate_to_folder(folder, push_history=False)

    def _on_folder_up(self) -> None:
        if self.current_folder is None:
            return
        parent = self.current_folder.parent
        if parent == self.current_folder:
            return
        self._navigate_to_folder(parent, push_history=True)

    # ------------------------------------------------------------------
    # 表示用HTML（tempプレビュー）関連
    # ------------------------------------------------------------------
    def _get_preview_html_path(self, xml_path: Path) -> Path:
        """
        表示用のHTMLを一時フォルダに作成する。
        同名XMLが別フォルダにあっても衝突しないように、フルパスのハッシュを付与する。
        """
        temp_root = Path(tempfile.gettempdir()) / self.APP_NAME
        temp_root.mkdir(parents=True, exist_ok=True)

        h = hashlib.sha1(str(xml_path).encode("utf-8")).hexdigest()[:10]
        return temp_root / f"{xml_path.stem}__preview__{h}.html"
    
    def _show_empty_message(self) -> None:
        """右側表示が空の時の案内メッセージを表示する。"""
        msg = (
            "左側のメニューよりフォルダを開いてxmlファイルを選択するか、"
            "xmlファイルをここにドラッグ・アンド・ドロップしてください。"
        )

        html = f"""
        <!doctype html>
        <html>
        <head>
        <meta charset="utf-8" />
        <style>
            html, body {{
            height: 100%;
            margin: 0;
            }}
            body {{
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: sans-serif;
            background: #ffffff;
            color: #333;
            }}
            .box {{
            max-width: 720px;
            padding: 24px 28px;
            border: 1px solid #ddd;
            border-radius: 12px;
            line-height: 1.8;
            font-size: 16px;
            }}
        </style>
        </head>
        <body>
        <div class="box">{msg}</div>
        </body>
        </html>
        """
        self.view.setHtml(html)
        self.statusBar().showMessage("準備完了")



    # ------------------------------------------------------------------
    # フォルダ / ファイル選択
    # ------------------------------------------------------------------
    def _on_open_folder_clicked(self) -> None:
        last_dir = self.settings.value("last_dir", "")
        initial_dir = last_dir if last_dir else str(Path.home())

        folder = QFileDialog.getExistingDirectory(
            self,
            "XML ファイルを含むフォルダを選択してください",
            initial_dir,
        )
        if not folder:
            return

        folder_path = Path(folder)
        self.settings.setValue("last_dir", str(folder_path))
        self._navigate_to_folder(folder_path, push_history=True)

    def _on_open_file_menu(self) -> None:
        last_dir = self.settings.value("last_dir", "")
        initial_dir = last_dir if last_dir else str(Path.home())

        xml_path_str, _ = QFileDialog.getOpenFileName(
            self,
            "XML ファイルを選択してください",
            initial_dir,
            "XML Files (*.xml);;All Files (*)",
        )
        if not xml_path_str:
            return

        xml_path = Path(xml_path_str)
        if not xml_path.exists():
            QMessageBox.warning(self, "ファイルが存在しません", str(xml_path))
            return

        folder_path = xml_path.parent
        self.settings.setValue("last_dir", str(folder_path))

        self._navigate_to_folder(folder_path, push_history=True)
        self.open_xml(xml_path)

    # ------------------------------------------------------------------
    # ファイルツリー関連
    # ------------------------------------------------------------------
    def _populate_tree(self, root_folder: Path) -> None:
        """root_folder 以下の XML ファイルをツリーに列挙する。"""
        self.tree.clear()

        root_item = QTreeWidgetItem([str(root_folder)])
        root_item.setData(0, Qt.ItemDataRole.UserRole, root_folder)
        self.tree.addTopLevelItem(root_item)

        for path in sorted(root_folder.rglob("*.xml")):
            relative = path.relative_to(root_folder)
            parts = relative.parts

            parent_item = root_item
            for i, part in enumerate(parts):
                if i == len(parts) - 1:
                    file_item = QTreeWidgetItem([part])
                    file_item.setData(0, Qt.ItemDataRole.UserRole, path)
                    parent_item.addChild(file_item)
                else:
                    found_child = None
                    for j in range(parent_item.childCount()):
                        child = parent_item.child(j)
                        if child.text(0) == part:
                            found_child = child
                            break
                    if found_child is None:
                        new_folder = QTreeWidgetItem([part])
                        new_folder.setData(
                            0,
                            Qt.ItemDataRole.UserRole,
                            parent_item.data(0, Qt.ItemDataRole.UserRole) / part,
                        )
                        parent_item.addChild(new_folder)
                        parent_item = new_folder
                    else:
                        parent_item = found_child

        self.tree.expandItem(root_item)

        self.current_folder = root_folder.resolve()
        self._update_folder_nav_buttons()

    def _on_tree_item_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        """ツリーで XML ファイルをダブルクリックしたときに、そのファイルを表示する。"""
        path_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not isinstance(path_data, Path):
            return
        if path_data.is_file() and path_data.suffix.lower() == ".xml":
            self.open_xml(path_data)

    # ------------------------------------------------------------------
    # D&D 入口（要望①）
    # ------------------------------------------------------------------
    def open_xml_via_drop(self, xml_path: Path) -> None:
        """
        HTML表示エリアへのドロップから呼ばれる入口。
        仕様：
        - xml 1ファイルのみ
        - 同フォルダに同名 .xsl が必須
        - ない場合はエラー
        """
        if not xml_path.exists() or not xml_path.is_file():
            QMessageBox.information(self, "ドロップエラー", "ファイルが存在しません。")
            return

        if xml_path.suffix.lower() != ".xml":
            QMessageBox.information(self, "ドロップエラー", "XMLファイル（.xml）をドロップしてください。")
            return

        xsl_path = xml_path.with_suffix(".xsl")
        if not xsl_path.exists():
            QMessageBox.critical(
                self,
                "スタイルシートが見つかりません",
                "同名のスタイルシート（.xsl）が同じフォルダに見つかりませんでした。\n\n"
                f"XML:\n{xml_path}\n\n"
                f"必要なXSL:\n{xsl_path}",
            )
            return

        # フォルダ表示を追従（last_dirも更新）
        folder_path = xml_path.parent
        self.settings.setValue("last_dir", str(folder_path))
        self._navigate_to_folder(folder_path, push_history=True)

        # 表示
        self.open_xml(xml_path)

    # ------------------------------------------------------------------
    # XML を開いて表示（表示はtempプレビューのみ）
    # ------------------------------------------------------------------
    def open_xml(self, xml_path: Path) -> None:
        """指定された XML ファイルを変換して右側ビューに表示する（表示用HTMLは temp のみ）"""
        if not xml_path.exists():
            QMessageBox.warning(self, "ファイルが存在しません", str(xml_path))
            return

        self.current_xml = xml_path

        preview_html = self._get_preview_html_path(xml_path)

        try:
            self.transformer.transform_to_html_file(xml_path, output_path=preview_html)
        except Exception as e:
            QMessageBox.critical(
                self,
                "変換エラー",
                "XML から HTML への変換に失敗しました。\n\n"
                f"XML: {xml_path}\n\n"
                f"エラー詳細:\n{e}",
            )
            self._show_empty_message()
            return

        self.view.setUrl(QUrl.fromLocalFile(str(preview_html)))
        self.statusBar().showMessage(f"表示中（プレビュー）: {xml_path}")

    # ------------------------------------------------------------------
    # HTML / PDF 保存関連
    # ------------------------------------------------------------------
    def _require_current_xml(self) -> Optional[Path]:
        """現在の XML ファイルが選択されていない場合は警告して None を返す。"""
        if self.current_xml is None:
            QMessageBox.information(
                self,
                "XML が未選択",
                "変換対象の XML ファイルが選択されていません。\n\n"
                "左側のファイルツリーから XML ファイルを選択してください。",
            )
            return None
        return self.current_xml

    def _on_save_html_menu(self) -> None:
        xml_path = self._require_current_xml()
        if not xml_path:
            return
        self._save_html_for_current_xml(xml_path)

    def _on_save_pdf_menu(self) -> None:
        xml_path = self._require_current_xml()
        if not xml_path:
            return
        self._save_pdf_for_current_xml(xml_path)

    def _save_html_for_current_xml(self, xml_path: Path) -> None:
        """保存操作をした時だけ、XMLと同じフォルダに正式保存する"""
        html_path = xml_path.with_suffix(".html")

        try:
            self.transformer.transform_to_html_file(xml_path, output_path=html_path)
        except Exception as e:
            QMessageBox.critical(
                self,
                "HTML 変換エラー",
                "HTML ファイルの出力に失敗しました。\n\n"
                f"出力先:\n{html_path}\n\n"
                f"エラー詳細:\n{e}",
            )
            return

        reply = QMessageBox.question(
            self,
            "HTML 変換完了",
            "HTML への変換が終了しました。\n"
            "標準のブラウザで HTML を開きますか？\n\n"
            f"出力先:\n{html_path}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )

        self.statusBar().showMessage(f"HTML 変換が終了しました: {html_path}")

        if reply == QMessageBox.StandardButton.Yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(html_path)))

    def _save_pdf_for_current_xml(self, xml_path: Path) -> None:
        """現在の表示内容を PDF として保存する"""
        pdf_path = xml_path.with_suffix(".pdf")

        try:
            page = self.view.page()
            if page is None:
                raise RuntimeError("Web ページがまだ読み込まれていません。")

            def _on_pdf_generated(data: bytes) -> None:
                try:
                    pdf_path.write_bytes(data)
                except Exception as e:
                    QMessageBox.critical(
                        self,
                        "PDF 書き込みエラー",
                        "PDF データの書き込みに失敗しました。\n\n"
                        f"出力先:\n{pdf_path}\n\n"
                        f"エラー詳細:\n{e}",
                    )
                    return

                reply = QMessageBox.question(
                    self,
                    "PDF 変換完了",
                    "PDF への変換が終了しました。\n"
                    "標準の PDF ビューワで開きますか？\n\n"
                    f"出力先:\n{pdf_path}",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )

                self.statusBar().showMessage(f"PDF 変換が終了しました: {pdf_path}")

                if reply == QMessageBox.StandardButton.Yes:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(pdf_path)))

            page.printToPdf(_on_pdf_generated)

        except Exception as e:
            QMessageBox.critical(
                self,
                "PDF 変換エラー",
                "PDF ファイルの出力に失敗しました。\n\n"
                f"出力先:\n{pdf_path}\n\n"
                f"エラー詳細:\n{e}",
            )
            return

        self.statusBar().showMessage(f"PDF 変換を開始しました: {pdf_path}")
