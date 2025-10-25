import subprocess
import os
import sys
import tempfile
import time
import shutil
from pathlib import Path
from typing import Optional, Any, Tuple

# UNO を利用した Excel → PDF 変換（各シート1ページ）
try:
    import uno  # type: ignore
    from com.sun.star.beans import PropertyValue  # type: ignore
except Exception:
    uno = None  # ランタイムで未導入の場合のフォールバックに利用
    PropertyValue = Any  # type: ignore  # フォールバック用


# 定数定義
UNO_PORT = 2002
UNO_HOST = "127.0.0.1"
UNO_CONNECTION_RETRY_COUNT = 30
UNO_CONNECTION_RETRY_DELAY = 0.2

# ページ設定の定数（単位: 1/100mm）
DEFAULT_MARGIN = 500
A4_LANDSCAPE_WIDTH = 29700
A4_LANDSCAPE_HEIGHT = 21000

# スケール設定の定数
MIN_SCALE = 10
MAX_SCALE = 100
MIN_DIMENSION = 100

# デフォルトDPI
DEFAULT_IMAGE_DPI = 200


class ConversionError(Exception):
    pass


def _find_soffice_executable() -> str:
    which = shutil.which("soffice")
    if which:
        return which
    raise ConversionError("LibreOffice 'soffice' not found (expected in the Docker image).")


def _imagemagick_convert_cmd() -> list[str]:
    if shutil.which("convert"):
        return ["convert"]
    if shutil.which("magick"):
        return ["magick", "convert"]
    raise ConversionError("ImageMagick 'convert' (or 'magick convert') not found")


def convert_office_to_pdf(input_path: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    soffice = _find_soffice_executable()
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--convert-to",
        "pdf:calc_pdf_Export" if input_path.suffix.lower() in {".xls", ".xlsx", ".xlsm"} else "pdf",
        "--outdir",
        str(output_dir),
        str(input_path),
    ]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except FileNotFoundError:
        raise ConversionError("LibreOffice 'soffice' not found")
    except subprocess.CalledProcessError as e:
        raise ConversionError(e.stderr.decode("utf-8", errors="ignore"))

    pdf_path = output_dir / (input_path.stem + ".pdf")
    if not pdf_path.exists():
        # LibreOffice may output with upper-case or different casing in rare cases
        candidates = list(output_dir.glob(input_path.stem + "*.pdf"))
        if candidates:
            return candidates[0]
        # 出力ディレクトリ内のすべてのPDFを確認
        all_pdfs = list(output_dir.glob("*.pdf"))
        if all_pdfs:
            # 予期しない名前でPDFが生成されている場合
            raise ConversionError(
                f"PDFファイルが予期しない名前で生成されました（期待: {pdf_path.name}、実際: {', '.join([p.name for p in all_pdfs[:3]])}）"
            )
        # PDFが全く生成されなかった場合
        raise ConversionError(
            f"LibreOfficeによる変換に失敗しました（ファイルが破損しているか、対応していない形式の可能性があります）"
        )
    return pdf_path


def _uno_property(name: str, value: Any) -> PropertyValue:
    """UNOプロパティを作成する"""
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _connect_to_uno() -> Any:
    """UNOサーバーに接続する。必要に応じてサーバーを起動する。"""
    if uno is None:
        raise ConversionError("UNO bindings not available")
    
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )

    def _try_connect():
        return resolver.resolve(
            f"uno:socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ComponentContext"
        )

    # 既存のサーバーへの接続を試行
    try:
        return _try_connect()
    except Exception:
        pass

    # サーバーを起動
    soffice = _find_soffice_executable()
    proc = subprocess.Popen([
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        f"--accept=socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ServiceManager",
    ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # 接続リトライ
    last_err = None
    for _ in range(UNO_CONNECTION_RETRY_COUNT):
        try:
            return _try_connect()
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(UNO_CONNECTION_RETRY_DELAY)
    
    # 接続失敗時はプロセスを終了
    try:
        proc.terminate()
    except Exception:
        pass
    raise ConversionError(f"Failed to connect to soffice via UNO: {last_err}")


def _trim_to_visible_range(sheet: Any, used_addr: Any) -> Any:
    """非表示の行/列を除外して可視範囲のアドレスを返す"""
    try:
        cols = sheet.getColumns()
        rows = sheet.getRows()
        sc, ec = used_addr.StartColumn, used_addr.EndColumn
        sr, er = used_addr.StartRow, used_addr.EndRow
        
        # 左端から可視列を探す
        while sc <= ec:
            try:
                if getattr(cols.getByIndex(sc), "IsVisible", True):
                    break
            except Exception:
                break
            sc += 1
        
        # 右端から可視列を探す
        while ec >= sc:
            try:
                if getattr(cols.getByIndex(ec), "IsVisible", True):
                    break
            except Exception:
                break
            ec -= 1
        
        # 上端から可視行を探す
        while sr <= er:
            try:
                if getattr(rows.getByIndex(sr), "IsVisible", True):
                    break
            except Exception:
                break
            sr += 1
        
        # 下端から可視行を探す
        while er >= sr:
            try:
                if getattr(rows.getByIndex(er), "IsVisible", True):
                    break
            except Exception:
                break
            er -= 1
        
        rng = sheet.getCellRangeByPosition(sc, sr, ec, er)
        return rng.getRangeAddress()
    except Exception:
        return used_addr


def _setup_print_area(sheet: Any) -> Any:
    """シートの印刷範囲を使用範囲に設定する"""
    try:
        # 既存の手動改ページをリセット
        try:
            reset_breaks = getattr(sheet, "resetAllPageBreaks", None)
            if callable(reset_breaks):
                reset_breaks()
        except Exception:
            pass

        # 使用範囲を取得
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        used_addr = cursor.getRangeAddress()
        
        # 非表示の行/列を除外
        used_addr = _trim_to_visible_range(sheet, used_addr)
        
        # 既存の印刷範囲をクリアして、使用範囲のみを設定
        try:
            sheet.setPrintAreas(())
        except Exception:
            pass
        sheet.setPrintAreas((used_addr,))
        
        # 印刷タイトル（先頭行/列の繰り返し）を解除
        try:
            if hasattr(sheet, "setTitleRows"):
                sheet.setTitleRows(())
        except Exception:
            pass
        try:
            if hasattr(sheet, "setTitleColumns"):
                sheet.setTitleColumns(())
        except Exception:
            pass
        
        return used_addr
    except Exception:
        # 印刷範囲設定に失敗しても続行
        return None


def _set_paper_size_and_orientation(style: Any, psi: Any) -> None:
    """用紙サイズをA4横向きに設定する"""
    try:
        # まず横向きを設定
        if psi.hasPropertyByName("IsLandscape"):
            style.setPropertyValue("IsLandscape", True)
    except Exception:
        pass


def _set_margins(style: Any, psi: Any) -> None:
    """余白を設定する"""
    margins = [
        ("TopMargin", DEFAULT_MARGIN),
        ("BottomMargin", DEFAULT_MARGIN),
        ("LeftMargin", DEFAULT_MARGIN),
        ("RightMargin", DEFAULT_MARGIN),
    ]
    for prop, val in margins:
        try:
            if psi.hasPropertyByName(prop):
                style.setPropertyValue(prop, val)
        except Exception:
            pass


def _set_scale_to_fit_one_page(style: Any, psi: Any) -> None:
    """1ページに収める設定を行う
    
    ScaleToPagesX/Y を使用してLibreOfficeの自動スケーリングを有効化します。
    これにより、コンテンツが自動的に1ページに収まるように縮小されます。
    """
    try:
        # 手動スケールを無効化（自動スケーリングと競合するため）
        if psi.hasPropertyByName("PageScale"):
            style.setPropertyValue("PageScale", 100)  # 100% = デフォルト（自動スケールに任せる）
        
        # 古い方式のScaleToPagesを無効化（ScaleToPagesX/Yと競合するため）
        if psi.hasPropertyByName("ScaleToPages"):
            style.setPropertyValue("ScaleToPages", 0)
        
        # 横1ページ、縦1ページに収める（自動スケーリング）
        if psi.hasPropertyByName("ScaleToPagesX"):
            style.setPropertyValue("ScaleToPagesX", 1)
        if psi.hasPropertyByName("ScaleToPagesY"):
            style.setPropertyValue("ScaleToPagesY", 1)
    except Exception:
        pass


def _fix_paper_dimensions(style: Any, psi: Any) -> None:
    """PaperWidth/Heightを横向きに合わせて調整する"""
    try:
        if psi.hasPropertyByName("PaperWidth") and psi.hasPropertyByName("PaperHeight"):
            pw = style.getPropertyValue("PaperWidth")
            ph = style.getPropertyValue("PaperHeight")
            if isinstance(pw, (int, float)) and isinstance(ph, (int, float)) and pw > 0 and ph > 0:
                if pw < ph:
                    style.setPropertyValue("PaperWidth", ph)
                    style.setPropertyValue("PaperHeight", pw)
    except Exception:
        pass


def _set_print_options(style: Any, psi: Any) -> None:
    """印刷オプションを設定する"""
    try:
        if psi.hasPropertyByName("PrintCenterHorizontally"):
            style.setPropertyValue("PrintCenterHorizontally", True)
        if psi.hasPropertyByName("PrintCenterVertically"):
            style.setPropertyValue("PrintCenterVertically", True)
        if psi.hasPropertyByName("PrintGrid"):
            style.setPropertyValue("PrintGrid", False)
        if psi.hasPropertyByName("PrintHeaders"):
            style.setPropertyValue("PrintHeaders", False)
    except Exception:
        pass


# ==============================================================================
# 以下の関数群は、手動スケール計算のために実装されていますが、
# 現在は使用していません（ScaleToPagesX/Yの自動スケーリングを優先）。
# 将来的に必要になる可能性があるため、参考として残しています。
# ==============================================================================

def _get_paper_dimensions(style: Any, psi: Any) -> Tuple[int, int]:
    """用紙サイズを取得する（横向き前提で幅>高さ）
    
    注: 現在未使用。将来の拡張用に保持。
    """
    paper_w, paper_h = A4_LANDSCAPE_WIDTH, A4_LANDSCAPE_HEIGHT
    try:
        if psi.hasPropertyByName("PaperWidth") and psi.hasPropertyByName("PaperHeight"):
            pw = style.getPropertyValue("PaperWidth")
            ph = style.getPropertyValue("PaperHeight")
            if isinstance(pw, (int, float)) and isinstance(ph, (int, float)) and pw > 0 and ph > 0:
                paper_w, paper_h = max(pw, ph), min(pw, ph)
    except Exception:
        pass
    return paper_w, paper_h


def _calculate_content_dimensions(sheet: Any, used_addr: Any) -> Tuple[int, int]:
    """使用範囲のコンテンツサイズを計算する（単位: 1/100mm）
    
    注: 現在未使用。将来の拡張用に保持。
    """
    content_w = 0
    content_h = 0
    try:
        cols = sheet.getColumns()
        for c in range(used_addr.StartColumn, used_addr.EndColumn + 1):
            try:
                cw = getattr(cols.getByIndex(c), "Width", None)
                if isinstance(cw, (int, float)) and cw > 0:
                    content_w += int(cw)
            except Exception:
                pass
        
        rows = sheet.getRows()
        for r in range(used_addr.StartRow, used_addr.EndRow + 1):
            try:
                rh = getattr(rows.getByIndex(r), "Height", None)
                if isinstance(rh, (int, float)) and rh > 0:
                    content_h += int(rh)
            except Exception:
                pass
    except Exception:
        pass
    return content_w, content_h


def _calculate_optimal_scale(
    content_w: int, 
    content_h: int, 
    paper_w: int, 
    paper_h: int,
    left: int,
    right: int,
    top: int,
    bottom: int
) -> int:
    """コンテンツが用紙に収まる最適なスケールを計算する
    
    注: 現在未使用。将来の拡張用に保持。
    """
    if content_w <= 0 or content_h <= 0:
        return MAX_SCALE
    
    printable_w = max(MIN_DIMENSION, paper_w - left - right)
    printable_h = max(MIN_DIMENSION, paper_h - top - bottom)
    
    scale_w = 100.0 * printable_w / content_w
    scale_h = 100.0 * printable_h / content_h
    
    return int(max(MIN_SCALE, min(MAX_SCALE, int(min(scale_w, scale_h)))))


def _apply_custom_scale(sheet: Any, style: Any, psi: Any, used_addr: Any) -> None:
    """必要に応じてカスタムスケールを適用する
    
    注: 現在未使用。ScaleToPagesX/Yの自動スケーリングと競合するため使用していません。
    将来の拡張用に保持。
    """
    try:
        # 用紙サイズと余白を取得
        paper_w, paper_h = _get_paper_dimensions(style, psi)
        
        top = style.getPropertyValue("TopMargin") if psi.hasPropertyByName("TopMargin") else DEFAULT_MARGIN
        bottom = style.getPropertyValue("BottomMargin") if psi.hasPropertyByName("BottomMargin") else DEFAULT_MARGIN
        left = style.getPropertyValue("LeftMargin") if psi.hasPropertyByName("LeftMargin") else DEFAULT_MARGIN
        right = style.getPropertyValue("RightMargin") if psi.hasPropertyByName("RightMargin") else DEFAULT_MARGIN
        
        # コンテンツサイズを計算
        content_w, content_h = _calculate_content_dimensions(sheet, used_addr)
        
        if content_w > 0 and content_h > 0:
            scale = _calculate_optimal_scale(
                content_w, content_h, paper_w, paper_h, left, right, top, bottom
            )
            
            # 100%未満が必要であればScaleToPagesを無効化しPageScaleを適用
            if scale < MAX_SCALE:
                try:
                    if psi.hasPropertyByName("ScaleToPagesX"):
                        style.setPropertyValue("ScaleToPagesX", 0)
                    if psi.hasPropertyByName("ScaleToPagesY"):
                        style.setPropertyValue("ScaleToPagesY", 0)
                    if psi.hasPropertyByName("ScaleToPages"):
                        style.setPropertyValue("ScaleToPages", 0)
                except Exception:
                    pass
                try:
                    if psi.hasPropertyByName("PageScale"):
                        style.setPropertyValue("PageScale", scale)
                except Exception:
                    pass
    except Exception:
        pass


def _configure_sheet_for_one_page(sheet: Any, page_styles: Any) -> None:
    """シートを1ページに収める設定を行う"""
    # 印刷範囲を設定
    used_addr = _setup_print_area(sheet)
    if used_addr is None:
        return
    
    # ページスタイルを取得
    style_name = sheet.getPropertyValue("PageStyle")
    style = page_styles.getByName(style_name)
    psi = style.getPropertySetInfo()
    
    # 用紙サイズと向きを設定（明示的にA4横向きを設定）
    _set_paper_size_and_orientation(style, psi)
    
    # 余白を設定
    _set_margins(style, psi)
    
    # 1ページに収める設定（LibreOfficeの自動スケーリングを使用）
    _set_scale_to_fit_one_page(style, psi)
    
    # 印刷オプションを設定
    _set_print_options(style, psi)
    
    # 注: _apply_custom_scale, _fix_paper_dimensions は使用しない
    # - _fix_paper_dimensions: 既に明示的にA4横向きを設定済み
    # - _apply_custom_scale: ScaleToPagesX/Yと競合するため
    # LibreOfficeの自動スケーリング（ScaleToPagesX=1, ScaleToPagesY=1）に任せる


def _fix_pdf_page_orientation_to_landscape(pdf_path: Path) -> None:
    """PDFの各ページをA4横向きに修正する
    
    LibreOffice CalcのPDFエクスポートはIsLandscape設定を無視するため、
    生成後のPDFを修正します。
    """
    try:
        import pikepdf
        
        # A4サイズ（ポイント単位）
        A4_WIDTH_PTS = 841.89  # 297mm
        A4_HEIGHT_PTS = 595.28  # 210mm
        
        pdf = pikepdf.Pdf.open(pdf_path, allow_overwriting_input=True)
        modified = False
        
        for page in pdf.pages:
            # 現在のMediaBox（ページサイズ）を取得
            if '/MediaBox' in page:
                media_box = page['/MediaBox']
                current_width = float(media_box[2]) - float(media_box[0])
                current_height = float(media_box[3]) - float(media_box[1])
                
                # 縦向き（高さ > 幅）の場合、横向きに変更
                if current_height > current_width:
                    # A4横向きに設定
                    page['/MediaBox'] = pikepdf.Array([0, 0, A4_WIDTH_PTS, A4_HEIGHT_PTS])
                    
                    # コンテンツを90度回転して横向きに合わせる
                    # （必要に応じて調整）
                    if '/Rotate' not in page:
                        page['/Rotate'] = 0
                    
                    modified = True
        
        if modified:
            pdf.save(pdf_path)
        pdf.close()
    except Exception:
        # pikepdf処理が失敗しても、元のPDFは残る（警告なし）
        pass


def convert_excel_to_pdf_fit_one_page(input_path: Path, output_pdf: Path) -> Path:
    """ExcelファイルをPDFに変換する（各シート1ページ）
    
    注意: LibreOffice CalcはIsLandscape設定をPDF出力時に無視する既知の問題があります。
    そのため、1ページに収まりますが、縦向きA4で出力されます。
    """
    if uno is None:
        raise ConversionError("UNO bindings not available for Excel conversion")

    # UNOサーバーに接続
    ctx = _connect_to_uno()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # ドキュメントを読み込み
    in_url = input_path.resolve().as_uri()
    props = (_uno_property("Hidden", True),)
    doc = desktop.loadComponentFromURL(in_url, "_blank", 0, props)

    try:
        # 各シートに対してページ設定を適用
        style_families = doc.getStyleFamilies()
        page_styles = style_families.getByName("PageStyles")
        sheets = doc.getSheets()
        
        for i in range(sheets.getCount()):
            sheet = sheets.getByIndex(i)
            _configure_sheet_for_one_page(sheet, page_styles)

        # 設定を反映
        try:
            calc_all = getattr(doc, "calculateAll", None)
            if callable(calc_all):
                calc_all()
        except Exception:
            pass

        # PDFにエクスポート
        out_url = output_pdf.resolve().as_uri()
        export_props = (_uno_property("FilterName", "calc_pdf_Export"),)
        doc.storeToURL(out_url, export_props)
    finally:
        doc.close(True)

    return output_pdf


def convert_image_to_pdf(input_path: Path, output_pdf: Path, dpi: Optional[int] = None) -> Path:
    """画像ファイルをPDFに変換する"""
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    base_cmd = _imagemagick_convert_cmd()
    cmd = [
        *base_cmd,
        str(input_path),
        "-units",
        "PixelsPerInch",
        "-density",
        str(dpi or DEFAULT_IMAGE_DPI),
        str(output_pdf),
    ]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except FileNotFoundError:
        raise ConversionError("ImageMagick command not found")
    except subprocess.CalledProcessError as e:
        raise ConversionError(e.stderr.decode("utf-8", errors="ignore"))
    return output_pdf


def convert_to_pdf(input_path: Path, workdir: Optional[Path] = None, image_dpi: Optional[int] = None) -> Path:
    """各種ファイル形式をPDFに変換する"""
    suffix = input_path.suffix.lower()
    tmpdir_created = False
    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="candytune-"))
        tmpdir_created = True
    try:
        if suffix in {".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".xlsm", ".csv"}:
            # Excel系はUNOによる1シート=1ページのPDF化を優先
            if uno is not None and suffix in {".xls", ".xlsx", ".xlsm"}:
                output_pdf = workdir / (input_path.stem + ".pdf")
                try:
                    return convert_excel_to_pdf_fit_one_page(input_path, output_pdf)
                except Exception:
                    # フォールバック: 標準のLibreOffice変換を使用（警告なし）
                    pass
            return convert_office_to_pdf(input_path, workdir)
        if suffix in {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}:
            return convert_image_to_pdf(
                input_path, 
                workdir / (input_path.stem + ".pdf"), 
                dpi=image_dpi
            )
        if suffix == ".pdf":
            # PDFはそのままコピー
            out = workdir / input_path.name
            shutil.copy2(input_path, out)
            return out
        raise ConversionError(f"Unsupported file type: {suffix}")
    finally:
        if tmpdir_created:
            shutil.rmtree(workdir, ignore_errors=True)


