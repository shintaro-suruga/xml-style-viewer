from __future__ import annotations

from pathlib import Path
import re
import shlex
from lxml import etree

# アプリ側で強制的に当てたい CSS
CUSTOM_CSS = """
/* 必要最小限の折り返しのみ。視覚デザインはXSLに完全に従う */
pre.oshirase {
    white-space: pre-wrap !important;
}
"""


class XmlXsltTransformer:
    """
    XML 内の xml-stylesheet を読み取り、
    「XMLファイル名と同名のXSL」を前提としたルールで
    XSLT 変換して HTML を返すクラス。
    """

    def __init__(self, xml_path: str | Path | None = None):
        # xml_path は optional にする（MainWindow 側で引数なし生成できるように）
        self.xml_path: Path | None = Path(xml_path) if xml_path is not None else None

    # --- 内部ヘルパー ---

    def _require_xml_path(self, xml_path: str | Path | None) -> Path:
        """
        メソッド引数 xml_path があればそれを採用、
        なければ self.xml_path を使う。どちらも無ければエラー。
        """
        if xml_path is not None:
            return Path(xml_path)
        if self.xml_path is not None:
            return self.xml_path
        raise TypeError("xml_path が指定されていません。transform_to_* の引数で指定してください。")

    def _load_xml_tree(self, xml_path: Path) -> etree._ElementTree:
        """XMLを読み込んでツリーにして返す"""
        if not xml_path.exists():
            raise FileNotFoundError(f"XMLファイルが見つかりません: {xml_path}")
        return etree.parse(str(xml_path))

    def _get_expected_xsl_name(self, xml_path: Path) -> str:
        """
        運用ルールに基づく「期待されるXSLファイル名」を返す。
        例: 7100001.xml -> 7100001.xsl / henrei.xml -> henrei.xsl
        """
        return f"{xml_path.stem}.xsl"

    def _read_stylesheet_href_from_pi(self, xml_tree: etree._ElementTree) -> str | None:
        """
        xml-stylesheet 処理命令から href を読み取る。
        見つからなければ None を返す。
        """
        root = xml_tree.getroot()
        node = root.getprevious()

        while node is not None:
            if isinstance(node, etree._ProcessingInstruction) and node.target == "xml-stylesheet":
                # 例: node.text -> 'type="text/xsl" href="7100001.xsl"'
                if node.text:
                    parts = shlex.split(node.text)
                    attrs: dict[str, str] = {}
                    for part in parts:
                        if "=" in part:
                            key, value = part.split("=", 1)
                            attrs[key] = value
                    return attrs.get("href")
            node = node.getprevious()

        return None

    def _resolve_xsl_path(self, xml_path: Path, xml_tree: etree._ElementTree) -> Path:
        """
        運用ルールに基づいて XSL ファイルのパスを決定する。

        優先順位:
        1. xml-stylesheet の href がある場合:
           - ファイル名が「期待されるXSL名」と違う → エラー
           - 同じ → そのパスを採用
        2. xml-stylesheet が無い場合:
           - 同じフォルダの「期待されるXSL名」があれば採用
           - 無ければエラー
        """
        expected_xsl_name = self._get_expected_xsl_name(xml_path)
        href_value = self._read_stylesheet_href_from_pi(xml_tree)

        if href_value:
            href_path = Path(href_value)
            href_name_only = href_path.name

            if href_name_only != expected_xsl_name:
                raise ValueError(
                    f"xml-stylesheet の href が運用ルールと異なります。\n"
                    f"  XMLファイル: {xml_path.name}\n"
                    f"  期待されるXSL名: {expected_xsl_name}\n"
                    f"  実際のhref: {href_value}"
                )

            xsl_path = (xml_path.parent / href_path).resolve()
        else:
            xsl_path = (xml_path.parent / expected_xsl_name).resolve()

        if not xsl_path.exists():
            raise FileNotFoundError(
                f"スタイルシートファイルが見つかりません。\n"
                f"  XMLファイル: {xml_path}\n"
                f"  探したXSLパス: {xsl_path}"
            )

        return xsl_path

    # --- HTML 後処理ヘルパー ---

    def _inject_custom_css(self, html: str) -> str:
        """
        生成された HTML の <head> 直前にアプリ専用 CSS を差し込む。
        <head> が無い場合は、先頭にスタイルだけ付ける簡易実装。
        """
        style_block = "<style>\n" + CUSTOM_CSS + "\n</style>\n"

        lower_html = html.lower()
        idx = lower_html.find("</head>")

        if idx != -1:
            return html[:idx] + style_block + html[idx:]
        else:
            return style_block + html

    def _has_meta_charset(self, html: str) -> bool:
        """meta charset がすでにあるか簡易判定"""
        return bool(re.search(r"<meta\s+[^>]*charset\s*=", html, re.IGNORECASE))

    def _inject_meta_charset(self, html: str, charset: str) -> str:
        """
        <head> 内に <meta charset="..."> を注入する。
        すでに meta charset があれば何もしない。
        """
        if self._has_meta_charset(html):
            return html

        meta = f'<meta charset="{charset}">\n'

        lower_html = html.lower()
        head_open_idx = lower_html.find("<head")
        if head_open_idx == -1:
            # headが無いなら先頭に追加（最低限の保険）
            return meta + html

        # <head ...> の ">" の位置を探す
        head_end = lower_html.find(">", head_open_idx)
        if head_end == -1:
            return meta + html

        insert_pos = head_end + 1
        return html[:insert_pos] + "\n" + meta + html[insert_pos:]

    def _normalize_charset(self, enc: str | None) -> str:
        """docinfo などで取れた encoding を保存向けに正規化する"""
        if not enc:
            return "UTF-8"
        e = enc.strip()
        # よくある揺れを少しだけ正規化
        if e.lower() in ("shift_jis", "shift-jis", "sjis", "ms932", "cp932"):
            return "Shift_JIS"
        return e

    def _get_output_encoding_from_xsl(self, xsl_path: Path) -> str:
        """
        XSL側の <xsl:output encoding="..."> を最優先で読む。
        無ければ UTF-8。
        """
        try:
            xsl_tree = etree.parse(str(xsl_path))
            root = xsl_tree.getroot()
            ns = root.nsmap.get("xsl", "http://www.w3.org/1999/XSL/Transform")
            out = root.find(f"{{{ns}}}output")
            if out is not None:
                enc = out.get("encoding")
                return self._normalize_charset(enc)
        except Exception:
            pass
        return "UTF-8"

    # --- 公開メソッド ---

    def transform_to_html_string(self, xml_path: str | Path | None = None) -> str:
        """XSLT 変換して HTML 文字列を返すメイン関数（xml_path は引数で渡せる）"""
        xml_path_p = self._require_xml_path(xml_path)

        xml_tree = self._load_xml_tree(xml_path_p)
        xsl_path = self._resolve_xsl_path(xml_path_p, xml_tree)

        xsl_tree = etree.parse(str(xsl_path))
        transform = etree.XSLT(xsl_tree)
        result_tree = transform(xml_tree)

        html_str = str(result_tree)
        html_str = self._inject_custom_css(html_str)

        return html_str

    def transform_to_html_file(
        self,
        xml_path: str | Path,
        output_path: str | Path | None = None,
    ) -> Path:
        """
        XSLT 変換して HTML を保存し、そのパスを返す。

        - output_path を省略した場合: XML と同名 .html
        - XSL側で encoding 指定がある場合は、その encoding で保存する
        """
        xml_path_p = Path(xml_path)

        html_str = self.transform_to_html_string(xml_path_p)

        xml_tree = self._load_xml_tree(xml_path_p)
        xsl_path = self._resolve_xsl_path(xml_path_p, xml_tree)
        out_enc = self._get_output_encoding_from_xsl(xsl_path)

        html_str = self._inject_meta_charset(html_str, out_enc)

        if output_path is None:
            output_path = xml_path_p.with_suffix(".html")

        output_path_p = Path(output_path)
        output_path_p.write_bytes(html_str.encode(out_enc, errors="replace"))
        return output_path_p

    def transform_to_debug_html_file(
        self,
        xml_path: str | Path,
        output_path: str | Path | None = None,
    ) -> Path:
        """互換のため残す：.debug.html を出力する"""
        xml_path_p = Path(xml_path)
        if output_path is None:
            output_path = xml_path_p.with_suffix(".debug.html")
        return self.transform_to_html_file(xml_path_p, output_path=output_path)


class XmlToStyledHtmlTransformer(XmlXsltTransformer):
    """互換性維持のための別名クラス。"""
    pass
