import subprocess
import os
import tempfile
import time
import shutil
from pathlib import Path
from typing import Optional

# UNO を利用した Excel → PDF 変換（各シート1ページ）
try:
    import uno  # type: ignore
    from com.sun.star.beans import PropertyValue  # type: ignore
except Exception:
    uno = None  # ランタイムで未導入の場合のフォールバックに利用


class ConversionError(Exception):
    pass


def _find_soffice_executable() -> str:
    # 優先: 環境変数
    for env_key in ("CANDYTUNE_SOFFICE", "SOFFICE_PATH"):
        v = os.environ.get(env_key)
        if v and Path(v).exists():
            return v
    # PATH 検索
    which = shutil.which("soffice")
    if which:
        return which
    # macOS 既定パス
    mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(mac_path).exists():
        return mac_path
    # Windows の一般的パス（存在チェックのみ）
    win_paths = [
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]
    for p in win_paths:
        if Path(p).exists():
            return p
    raise ConversionError(
        "LibreOffice 'soffice' が見つかりません。インストールするか、CANDYTUNE_SOFFICE/SOFFICE_PATH でパスを指定してください。"
    )


def _imagemagick_convert_cmd() -> list[str]:
    # 通常: convert。見つからなければ 'magick convert'
    if shutil.which("convert"):
        return ["convert"]
    if shutil.which("magick"):
        return ["magick", "convert"]
    raise ConversionError("ImageMagick の 'convert' (または 'magick convert') が見つかりません")


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
        raise ConversionError("LibreOffice 'soffice' が見つかりません")
    except subprocess.CalledProcessError as e:
        raise ConversionError(e.stderr.decode("utf-8", errors="ignore"))

    pdf_path = output_dir / (input_path.stem + ".pdf")
    if not pdf_path.exists():
        # LibreOffice may output with upper-case or different casing in rare cases
        candidates = list(output_dir.glob(input_path.stem + "*.pdf"))
        if candidates:
            return candidates[0]
        raise ConversionError("PDF output not found")
    return pdf_path


def _uno_property(name: str, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def convert_excel_to_pdf_fit_one_page(input_path: Path, output_pdf: Path) -> Path:
    if uno is None:
        raise ConversionError("UNO bindings not available for Excel conversion")

    # 既存の soffice に接続を試行。失敗したら受け口付きで起動して再試行。
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)

    def _connect():
        return resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")

    try:
        ctx = _connect()
    except Exception:
        # 起動
        soffice = _find_soffice_executable()
        proc = subprocess.Popen([
            soffice,
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager",
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # 接続リトライ
        last_err = None
        for _ in range(30):
            try:
                ctx = _connect()
                break
            except Exception as e:  # noqa: BLE001
                last_err = e
                time.sleep(0.2)
        else:
            try:
                proc.terminate()
            except Exception:
                pass
            raise ConversionError(f"Failed to connect to soffice via UNO: {last_err}")

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    in_url = input_path.resolve().as_uri()
    props = (_uno_property("Hidden", True),)
    doc = desktop.loadComponentFromURL(in_url, "_blank", 0, props)

    try:
        # PageStyles を取得して各シートのページスタイルにスケール設定を適用
        style_families = doc.getStyleFamilies()
        page_styles = style_families.getByName("PageStyles")
        sheets = doc.getSheets()
        # 単純化: 用紙はA4固定、横向き固定のみ（環境によっては viewer 側回転あり）
        desired_paper = "A4"
        for i in range(sheets.getCount()):
            sheet = sheets.getByIndex(i)
            # 単純化: アクティブ化は行わない

            # 既存の手動改ページをリセット（存在する環境のみ）
            try:
                reset_breaks = getattr(sheet, "resetAllPageBreaks", None)
                if callable(reset_breaks):
                    reset_breaks()
            except Exception:
                pass

            # 使用範囲を自動で印刷範囲に設定（手動の印刷範囲や改ページの影響を抑制）
            try:
                cursor = sheet.createCursor()
                cursor.gotoStartOfUsedArea(False)
                cursor.gotoEndOfUsedArea(True)
                used_addr = cursor.getRangeAddress()
                # 端の非表示 行/列 をトリム（印刷面積を縮めない）
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
                    used_addr = rng.getRangeAddress()
                except Exception:
                    pass
                # いったん既存の印刷範囲をクリアしてから、単一の使用範囲のみを設定
                try:
                    sheet.setPrintAreas(())
                except Exception:
                    pass
                # setPrintAreas は sequence<CellRangeAddress> を受け取る
                sheet.setPrintAreas((used_addr,))
                # 印刷タイトル（先頭行/列の繰り返し）を解除（存在する環境のみ）
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
            except Exception:
                # 印刷範囲設定に失敗しても続行
                pass

            # ページスタイル調整
            style_name = sheet.getPropertyValue("PageStyle")
            style = page_styles.getByName(style_name)
            psi = style.getPropertySetInfo()

            # 用紙サイズをA4に設定（横向き固定）
            try:
                if psi.hasPropertyByName("PaperFormat"):
                    a4 = uno.getConstantByName("com.sun.star.view.PaperFormat.A4")
                    style.setPropertyValue("PaperFormat", a4)
                # 明示的に横向きを強制
                if psi.hasPropertyByName("IsLandscape"):
                    style.setPropertyValue("IsLandscape", True)
                if psi.hasPropertyByName("PaperOrientation"):
                    landscape = uno.getConstantByName("com.sun.star.view.PaperOrientation.LANDSCAPE")
                    style.setPropertyValue("PaperOrientation", landscape)
            except Exception:
                pass

            # 余白を小さく（単位: 1/100 mm）
            for prop, val in (
                ("TopMargin", 500),
                ("BottomMargin", 500),
                ("LeftMargin", 500),
                ("RightMargin", 500),
            ):
                try:
                    if psi.hasPropertyByName(prop):
                        style.setPropertyValue(prop, val)
                except Exception:
                    pass

            # 1ページに収める（X=1, Y=1 を優先。ScaleToPages は 0=無効）
            try:
                if psi.hasPropertyByName("ScaleToPagesX"):
                    style.setPropertyValue("ScaleToPagesX", 1)
                if psi.hasPropertyByName("ScaleToPagesY"):
                    style.setPropertyValue("ScaleToPagesY", 1)
                if psi.hasPropertyByName("ScaleToPages"):
                    style.setPropertyValue("ScaleToPages", 0)
                # PageScale が効いていると干渉するケースがあるため 0 クリア（0 は無効）
                if psi.hasPropertyByName("PageScale"):
                    style.setPropertyValue("PageScale", 0)
            except Exception:
                pass

            # PaperWidth/Height の整合（横向きで幅>高さになるように入替）
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

            # センタリング
            try:
                if psi.hasPropertyByName("PrintCenterHorizontally"):
                    style.setPropertyValue("PrintCenterHorizontally", True)
                if psi.hasPropertyByName("PrintCenterVertically"):
                    style.setPropertyValue("PrintCenterVertically", True)
                # 余計な印刷情報を抑止（存在時）
                if psi.hasPropertyByName("PrintGrid"):
                    style.setPropertyValue("PrintGrid", False)
                if psi.hasPropertyByName("PrintHeaders"):
                    style.setPropertyValue("PrintHeaders", False)
            except Exception:
                pass

            # Fit-to-one-page が効かない場合の保険として PageScale を自動算出
            try:
                # 用紙サイズ（A4 landscape）と余白から印刷可能領域を算出（単位: 1/100mm）
                paper_w, paper_h = (29700, 21000)
                try:
                    # もしユーザ定義サイズ等により PaperWidth/Height が取得できればそれを優先
                    if psi.hasPropertyByName("PaperWidth") and psi.hasPropertyByName("PaperHeight"):
                        pw = style.getPropertyValue("PaperWidth")
                        ph = style.getPropertyValue("PaperHeight")
                        if isinstance(pw, (int, float)) and isinstance(ph, (int, float)) and pw > 0 and ph > 0:
                            # landscape 前提で幅>高さになるように入れ替え
                            paper_w, paper_h = (max(pw, ph), min(pw, ph))
                except Exception:
                    pass

                top = style.getPropertyValue("TopMargin") if psi.hasPropertyByName("TopMargin") else 500
                bottom = style.getPropertyValue("BottomMargin") if psi.hasPropertyByName("BottomMargin") else 500
                left = style.getPropertyValue("LeftMargin") if psi.hasPropertyByName("LeftMargin") else 500
                right = style.getPropertyValue("RightMargin") if psi.hasPropertyByName("RightMargin") else 500
                printable_w = max(100, paper_w - left - right)
                printable_h = max(100, paper_h - top - bottom)

                # 使用範囲の実寸法を列幅/行高から計算
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

                if content_w > 0 and content_h > 0:
                    def best_scale_for(pw: int, ph: int) -> int:
                        tw = max(100, pw - left - right)
                        th = max(100, ph - top - bottom)
                        sw = 100.0 * tw / content_w
                        sh = 100.0 * th / content_h
                        return int(max(10, min(100, int(min(sw, sh)))))

                    scale = best_scale_for(paper_w, paper_h)
                    # 100%未満が必要であれば ScaleToPages* を無効化し PageScale を適用
                    if scale < 100:
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

        # 設定反映のため再計算（存在する環境のみ）
        try:
            calc_all = getattr(doc, "calculateAll", None)
            if callable(calc_all):
                calc_all()
        except Exception:
            pass
        finally:
            try:
                if hasattr(doc, "unlockControllers"):
                    doc.unlockControllers()
            except Exception:
                pass

        out_url = output_pdf.resolve().as_uri()
        # PDFエクスポートは最小限の指定のみ
        export_props = (_uno_property("FilterName", "calc_pdf_Export"),)
        doc.storeToURL(out_url, export_props)
    finally:
        doc.close(True)

    return output_pdf


def convert_image_to_pdf(input_path: Path, output_pdf: Path, dpi: Optional[int] = 200) -> Path:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    base_cmd = _imagemagick_convert_cmd()
    cmd = [
        *base_cmd,
        str(input_path),
        "-units",
        "PixelsPerInch",
        "-density",
        str(dpi or 200),
        str(output_pdf),
    ]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except FileNotFoundError:
        raise ConversionError("ImageMagick のコマンドが見つかりません")
    except subprocess.CalledProcessError as e:
        raise ConversionError(e.stderr.decode("utf-8", errors="ignore"))
    return output_pdf


def convert_to_pdf(input_path: Path, workdir: Optional[Path] = None, image_dpi: Optional[int] = None) -> Path:
    suffix = input_path.suffix.lower()
    tmpdir_created = False
    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="candytune-"))
        tmpdir_created = True
    try:
        if suffix in {".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".xlsm", ".csv"}:
            # Excel 系は UNO による1シート=1ページのPDF化を優先
            if uno is not None and suffix in {".xls", ".xlsx", ".xlsm"}:
                output_pdf = (workdir / (input_path.stem + ".pdf"))
                try:
                    return convert_excel_to_pdf_fit_one_page(input_path, output_pdf)
                except Exception:
                    # 単純化: フォールバック許容
                    pass
            return convert_office_to_pdf(input_path, workdir)
        if suffix in {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}:
            return convert_image_to_pdf(input_path, workdir / (input_path.stem + ".pdf"), dpi=image_dpi or 200)
        if suffix == ".pdf":
            # no-op copy
            out = workdir / input_path.name
            shutil.copy2(input_path, out)
            return out
        raise ConversionError(f"Unsupported file type: {suffix}")
    finally:
        if tmpdir_created:
            shutil.rmtree(workdir, ignore_errors=True)


