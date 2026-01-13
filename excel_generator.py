
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

# --- 定数定義 (Consts) ---
# シート名や色コードなどを定数として定義し、再利用性とメンテナンス性を向上させます。
SHEET_DB = "管理台帳"
SHEET_LABEL = "ラベル印刷"
HEADER_FILL_COLOR = "4472C4"  # 青色
HEADER_FONT_COLOR = "FFFFFF"  # 白色
INPUT_CELL_FILL_COLOR = "FFFF00"  # 黄色

def _apply_header_style(ws: Worksheet):
    """ヘッダー行にスタイルを適用するヘルパー関数"""
    header_fill = PatternFill(start_color=HEADER_FILL_COLOR, end_color=HEADER_FILL_COLOR, fill_type="solid")
    header_font = Font(color=HEADER_FONT_COLOR, bold=True)

    for cell in ws[1]:  # 1行目の全セルをループ
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

def _generate_qr_code_formula(data_parts_formula: str) -> str:
    """QuickChart APIを使用したQRコードのExcel数式を生成するヘルパー関数"""
    # URLエンコードし、IMAGE関数で画像として表示する数式を組み立てる
    return f'IMAGE("https://quickchart.io/qr?text=" & ENCODEURL({data_parts_formula}) & "&size=150")'

def _apply_border_to_range(ws: Worksheet, cell_range: str):
    """指定された範囲の外側に中太の罫線を引くヘルパー関数"""
    thin_side = Side(style='thin')
    medium_side = Side(style='medium')

    rows = list(ws[cell_range])

    # 範囲内の全セルに一旦細い罫線を引く（結合セル対策）
    for row in rows:
        for cell in row:
            cell.border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

    # 外枠を中太罫線で上書き
    top_row = rows[0]
    bottom_row = rows[-1]
    left_col_idx = top_row[0].column
    right_col_idx = top_row[-1].column

    for cell in top_row:
        cell.border = Border(top=medium_side, left=cell.border.left, right=cell.border.right, bottom=cell.border.bottom)
    for cell in bottom_row:
        cell.border = Border(bottom=medium_side, left=cell.border.left, right=cell.border.right, top=cell.border.top)

    for row in ws.iter_rows(min_row=rows[0][0].row, max_row=rows[-1][0].row, min_col=left_col_idx, max_col=left_col_idx):
        for cell in row:
            cell.border = Border(left=medium_side, top=cell.border.top, bottom=cell.border.bottom, right=cell.border.right)

    for row in ws.iter_rows(min_row=rows[0][0].row, max_row=rows[-1][0].row, min_col=right_col_idx, max_col=right_col_idx):
        for cell in row:
            cell.border = Border(right=medium_side, top=cell.border.top, bottom=cell.border.bottom, left=cell.border.left)


def create_database_sheet(wb: openpyxl.Workbook):
    """管理台帳シートを作成・設定する"""
    ws_db = wb.active
    ws_db.title = SHEET_DB

    # ヘッダー作成
    headers = ["ID", "品名", "品番", "数量", "依頼者", "ステータス", "QRコード"]
    ws_db.append(headers)

    # デザイン適用
    _apply_header_style(ws_db)

    # サンプルデータ
    sample_data = [
        ["S-001", "Fabric A", "F-100", "2 pcs", "T.Yamada", "発注済"],
        ["S-002", "Plastic B", "P-200", "10 pcs", "S.Suzuki", "発注済"],
        ["S-003", "Metal C", "M-300", "1 set", "K.Sato", "到着済"]
    ]
    for row in sample_data:
        ws_db.append(row)

    # QRコード生成数式の埋め込み（データがある行まで自動処理）
    # ws_db.max_row は現在データがある最終行を返す
    # +1 することで、次の新しい行にも数式が設定される（手入力用）
    # 今回はサンプルデータ分のみ設定する
    for i in range(2, ws_db.max_row + 1):
        cell = ws_db.cell(row=i, column=7)

        # QRコードに含めるデータ部分の数式を組み立て
        data_part = (
            f'"ID: " & A{i} & CHAR(10) & '
            f'"品名: " & B{i} & CHAR(10) & '
            f'"品番: " & C{i} & CHAR(10) & '
            f'"数量: " & D{i} & CHAR(10) & '
            f'"依頼者: " & E{i}'
        )

        formula = f'=IF(A{i}="","", {_generate_qr_code_formula(data_part)})'
        cell.value = formula
        ws_db.row_dimensions[i].height = 60

    # 列幅の調整
    ws_db.column_dimensions['A'].width = 12
    ws_db.column_dimensions['B'].width = 25
    ws_db.column_dimensions['E'].width = 15
    ws_db.column_dimensions['G'].width = 15

def create_label_sheet(wb: openpyxl.Workbook):
    """ラベル印刷シートを作成・設定する"""
    ws_lbl = wb.create_sheet(SHEET_LABEL)

    # ID入力欄 (B1)
    ws_lbl["A1"] = "印刷ID:"
    ws_lbl["B1"] = "S-001"

    # スタイル設定
    input_cell = ws_lbl["B1"]
    input_cell.fill = PatternFill(start_color=INPUT_CELL_FILL_COLOR, end_color=INPUT_CELL_FILL_COLOR, fill_type="solid")
    input_cell.font = Font(bold=True, size=12)
    input_cell.alignment = Alignment(horizontal="center")

    # ラベルタイトル
    ws_lbl.merge_cells("B3:E3")
    title_cell = ws_lbl["B3"]
    title_cell.value = "SAMPLE MANAGEMENT LABEL"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # QRコード表示エリア
    ws_lbl.merge_cells("E4:E7")
    qr_cell = ws_lbl["E4"]

    # QRコードに含めるデータ部分の数式 (XLOOKUP使用)
    # ここでもヘルパー関数を使い、ロジックを共通化
    label_qr_data = (
        f'"ID: " & B1 & CHAR(10) & '
        f'"品名: " & XLOOKUP(B1,{SHEET_DB}!A:A,{SHEET_DB}!B:B,"") & CHAR(10) & '
        f'"品番: " & XLOOKUP(B1,{SHEET_DB}!A:A,{SHEET_DB}!C:C,"") & CHAR(10) & '
        f'"数量: " & XLOOKUP(B1,{SHEET_DB}!A:A,{SHEET_DB}!D:D,"") & CHAR(10) & '
        f'"依頼者: " & XLOOKUP(B1,{SHEET_DB}!A:A,{SHEET_DB}!E:E,"")'
    )
    qr_cell.value = f'=IF(B1="","", {_generate_qr_code_formula(label_qr_data)})'
    qr_cell.alignment = Alignment(horizontal="center", vertical="center")

    # データ項目
    labels = {
        4: ("ID:", '=B1'),
        5: ("品名:", f'=XLOOKUP($B$1,{SHEET_DB}!A:A,{SHEET_DB}!B:B,"")'),
        6: ("品番:", f'=XLOOKUP($B$1,{SHEET_DB}!A:A,{SHEET_DB}!C:C,"")'),
        7: ("数量:", f'=XLOOKUP($B$1,{SHEET_DB}!A:A,{SHEET_DB}!D:D,"")')
    }

    for r, (txt, formula) in labels.items():
        ws_lbl.cell(row=r, column=2).value = txt
        ws_lbl.cell(row=r, column=2).font = Font(size=9, color="555555")
        ws_lbl.cell(row=r, column=2).alignment = Alignment(horizontal="right", vertical="center")

        ws_lbl.merge_cells(f"C{r}:D{r}")
        data_cell = ws_lbl.cell(row=r, column=3)
        data_cell.value = formula
        data_cell.font = Font(bold=True, size=11)
        data_cell.alignment = Alignment(horizontal="left", vertical="center")

    # 依頼者
    ws_lbl.merge_cells("B8:E8")
    ws_lbl["B8"].value = "ATTN (REQUESTER):"
    ws_lbl["B8"].font = Font(size=8)

    ws_lbl.merge_cells("B9:E9")
    attn_val = ws_lbl["B9"]
    attn_val.value = f'=XLOOKUP($B$1,{SHEET_DB}!A:A,{SHEET_DB}!E:E,"")'
    attn_val.font = Font(bold=True, size=16)
    attn_val.alignment = Alignment(horizontal="center", vertical="center")

    # 罫線の適用
    _apply_border_to_range(ws_lbl, "B3:E9")

    # 列幅調整
    ws_lbl.column_dimensions['A'].width = 8
    ws_lbl.column_dimensions['B'].width = 10
    ws_lbl.column_dimensions['C'].width = 15
    ws_lbl.column_dimensions['D'].width = 15
    ws_lbl.column_dimensions['E'].width = 18

def create_final_excel(filename="Sample_System_Final_Improved.xlsx"):
    """Excelワークブックを作成し、2つのシートを追加して保存するメイン関数"""
    wb = openpyxl.Workbook()

    create_database_sheet(wb)
    create_label_sheet(wb)

    wb.save(filename)
    print(f"作成完了: {filename}")

if __name__ == "__main__":
    create_final_excel()
