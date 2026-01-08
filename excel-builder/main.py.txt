from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
import io, zipfile
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

app = FastAPI()

CATEGORIES = [
    "완제품",
    "완제품(수탁)",
    "반제품",
    "반제품(구매)",
    "반제품(수탁)",
    "상품",
    "주원료",
    "부원료",
    "포장재",
]

# receipt 계산 열(문자 기준)
RECEIPT_COLS = {
    "in":  ["AC","AD","AF","AG","AP","AQ"],   # (AC+AD)-(AF+AG)-(AP+AQ)
    "out": ["BD","BE","BG","BH","BS","BT"],   # (BD+BE)-(BG+BH)-(BS+BT)
    "end": ["BY"],                            # BY
}

# wip 합계 열
WIP_COLS = {"in": "P", "out": "Q", "end": "R"}

def n(v):
    """숫자 변환: None/공백/문자면 0, 숫자면 float"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if s == "":
            return 0.0
        try:
            return float(s)
        except:
            return 0.0
    return 0.0

def col(ws, col_letter, row):
    return ws.cell(row=row, column=column_index_from_string(col_letter)).value

def sum_column_numeric(ws, col_letter, start_row=1):
    total = 0.0
    for r in range(start_row, ws.max_row + 1):
        total += n(col(ws, col_letter, r))
    return total

def build_plnt_map(plnt_wb):
    """plnt.xlsx: 자재내역 우선 key -> 실제플랜트(F열)"""
    ws = plnt_wb.active  # Sheet1
    # 헤더 2행, 데이터 3행부터
    # 자재, 자재내역, 실제플랜트(F)
    # 자재/자재내역이 정확히 몇 열인지 고정이 아니라면 헤더에서 찾는 게 안전
    header_row = 2
    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=c).value
        if isinstance(val, str):
            headers[val.strip()] = c

    # 필수 헤더명(사용자 제공): "자재", "자재내역"
    if "자재" not in headers or "자재내역" not in headers:
        # 그래도 진행은 하되 치환은 0건 처리
        return {}, {"plnt_header_missing": True}

    mat_col = headers["자재"]
    mat_desc_col = headers["자재내역"]
    real_plnt_col = 6  # F열 고정

    m = {}
    for r in range(3, ws.max_row + 1):
        mat = ws.cell(row=r, column=mat_col).value
        mat_desc = ws.cell(row=r, column=mat_desc_col).value
        real_plnt = ws.cell(row=r, column=real_plnt_col).value
        key1 = str(mat_desc).strip() if mat_desc not in (None, "") else ""
        key2 = str(mat).strip() if mat not in (None, "") else ""
        rp = str(real_plnt).strip() if real_plnt not in (None, "") else ""
        if rp:
            if key1:
                m[key1] = rp
            if key2:
                m[key2] = rp
    return m, {"plnt_header_missing": False}

def parse_receipt(receipt_wb, plnt_map):
    ws = receipt_wb.active  # Sheet1
    # 6행 헤더, 7행부터 데이터
    header_row = 6
    data_start = 7

    # 헤더에서 "내역", "자재", "자재내역" 위치 찾기 (안정성)
    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=c).value
        if isinstance(val, str) and val.strip():
            headers[val.strip()] = c

    for req in ["내역", "자재", "자재내역"]:
        if req not in headers:
            raise HTTPException(status_code=400, detail=f"receipt.xlsx에서 헤더 '{req}'를 찾지 못했습니다(6행).")

    cat_col = headers["내역"]
    mat_col = headers["자재"]
    mat_desc_col = headers["자재내역"]

    sums = {k: {"in": 0.0, "out": 0.0, "end": 0.0, "rows": 0} for k in CATEGORIES}
    unknown = {}

    # 완제품/완제품(수탁) 치환 검증용 카운트
    repl_total = 0
    repl_hit = 0

    for r in range(data_start, ws.max_row + 1):
        cat = ws.cell(row=r, column=cat_col).value
        if cat is None:
            continue
        cat_s = str(cat).strip()
        if cat_s == "":
            continue

        # 행별 입고/출고/기말 계산
        in_v = (n(col(ws, "AC", r)) + n(col(ws, "AD", r))) - (n(col(ws, "AF", r)) + n(col(ws, "AG", r))) - (n(col(ws, "AP", r)) + n(col(ws, "AQ", r)))
        out_v = (n(col(ws, "BD", r)) + n(col(ws, "BE", r))) - (n(col(ws, "BG", r)) + n(col(ws, "BH", r))) - (n(col(ws, "BS", r)) + n(col(ws, "BT", r)))
        end_v = n(col(ws, "BY", r))

        # 치환 로직(완제품 계열만): 결과에 직접 사용하지 않더라도 검증/향후 확장용
        if "완제품" in cat_s:
            repl_total += 1
            mat = ws.cell(row=r, column=mat_col).value
            mat_desc = ws.cell(row=r, column=mat_desc_col).value
            key = str(mat_desc).strip() if mat_desc not in (None, "") else str(mat).strip() if mat not in (None, "") else ""
            if key and key in plnt_map:
                repl_hit += 1

        if cat_s in sums:
            sums[cat_s]["in"] += in_v
            sums[cat_s]["out"] += out_v
            sums[cat_s]["end"] += end_v
            sums[cat_s]["rows"] += 1
        else:
            unknown.setdefault(cat_s, {"in": 0.0, "out": 0.0, "end": 0.0, "rows": 0})
            unknown[cat_s]["in"] += in_v
            unknown[cat_s]["out"] += out_v
            unknown[cat_s]["end"] += end_v
            unknown[cat_s]["rows"] += 1

    return sums, unknown, {"repl_total": repl_total, "repl_hit": repl_hit}

def write_values_and_formulas(template_wb, receipt_sums, wip_sums):
    ws = template_wb.active  # Sheet1 (단일 시트)

    # 1) 공란 유지 셀(명시적으로 비우기)
    blanks = ["B5", "B10", "B13", "B16", "B20", "B23", "B45", "C44", "D44", "C45", "D45"]
    for addr in blanks:
        ws[addr].value = None

    # 2) receipt 기말(B3~)
    ws["B3"].value  = receipt_sums["완제품"]["end"]
    ws["B4"].value  = receipt_sums["완제품(수탁)"]["end"]
    ws["B7"].value  = receipt_sums["반제품"]["end"]
    ws["B8"].value  = receipt_sums["반제품(구매)"]["end"]
    ws["B9"].value  = receipt_sums["반제품(수탁)"]["end"]
    ws["B12"].value = receipt_sums["상품"]["end"]
    ws["B15"].value = receipt_sums["주원료"]["end"]
    ws["B18"].value = receipt_sums["부원료"]["end"]
    ws["B19"].value = receipt_sums["포장재"]["end"]
    ws["B22"].value = wip_sums["end"]

    # 3) 37~43 입고/출고(C/D)
    ws["C37"].value = receipt_sums["완제품"]["in"]
    ws["D37"].value = receipt_sums["완제품"]["out"]

    ws["C38"].value = receipt_sums["반제품"]["in"] + receipt_sums["반제품(수탁)"]["in"]
    ws["D38"].value = receipt_sums["반제품"]["out"] + receipt_sums["반제품(수탁)"]["out"]

    ws["C39"].value = receipt_sums["반제품(구매)"]["in"]
    ws["D39"].value = receipt_sums["반제품(구매)"]["out"]

    ws["C40"].value = receipt_sums["상품"]["in"]
    ws["D40"].value = receipt_sums["상품"]["out"]

    ws["C41"].value = receipt_sums["주원료"]["in"] + receipt_sums["부원료"]["in"]
    ws["D41"].value = receipt_sums["주원료"]["out"] + receipt_sums["부원료"]["out"]

    ws["C42"].value = receipt_sums["포장재"]["in"]
    ws["D42"].value = receipt_sums["포장재"]["out"]

    ws["C43"].value = wip_sums["in"]
    ws["D43"].value = wip_sums["out"]

    # 4) 계/합계 수식(B열 상단)
    ws["B6"].value  = "=B3+B4-B5"
    ws["B11"].value = "=B7+B8+B9-B10"
    ws["B14"].value = "=B12-B13"
    ws["B17"].value = "=B15-B16"
    ws["B21"].value = "=B19-B20"
    ws["B24"].value = "=B6+B11+B14+B18+B21+B22+B23"

    # 5) 37~46 기말(B) 수식
    ws["B37"].value = "=B3+B4"
    ws["B38"].value = "=B7+B9"
    ws["B39"].value = "=B8"
    ws["B40"].value = "=B12"
    ws["B41"].value = "=B15+B18"
    ws["B42"].value = "=B19"
    ws["B43"].value = "=B22"
    ws["B44"].value = "=B23"
    ws["B46"].value = "=SUM(B37:B45)"

    # 6) 46행 합계 수식(C/D)
    ws["C46"].value = "=SUM(C37:C43)"
    ws["D46"].value = "=SUM(D37:D43)"

    # 7) 재고일수(근본 수식): 기말/출고*30
    for row in range(37, 44):
        ws[f"E{row}"].value = f'=IF(D{row}=0,"",B{row}/D{row}*30)'
    ws["E46"].value = '=IF(D46=0,"",B46/D46*30)'

    # 8) 매출원가율(%): (매출원가/매출액)*100
    # 당월: E27/E29, 누적: E28/E30 (누적 값이 비어있어도 수식은 세팅)
    ws["E31"].value = '=IF(E29=0,"",E27/E29*100)'
    ws["E32"].value = '=IF(E30=0,"",E28/E30*100)'

    # 9) 회전일(근본 수식): 재고/매출원가*30, 재고/매출액*30
    ws["E25"].value = '=IF(E27=0,"",B24/E27*30)'
    ws["E26"].value = '=IF(E29=0,"",B24/E29*30)'

    # 열 때 재계산 유도(안정장치)
    template_wb.calculation.fullCalcOnLoad = True

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/build")
async def build(zip_file: UploadFile = File(...)):
    """
    zip_file: package.zip
      - receipt.xlsx
      - plnt.xlsx
      - wip.xlsx
      - template.xlsx
    return: final xlsx
    """
    if not zip_file.filename.lower().endswith(".zip"):
        raise HTTPException(status_code=400, detail="zip 파일(package.zip)을 업로드해야 합니다.")

    raw = await zip_file.read()
    try:
        z = zipfile.ZipFile(io.BytesIO(raw))
    except Exception:
        raise HTTPException(status_code=400, detail="zip 파일을 열 수 없습니다. 손상되었거나 형식이 다릅니다.")

    required = ["receipt.xlsx", "plnt.xlsx", "wip.xlsx", "template.xlsx"]
    names_lower = {n.lower(): n for n in z.namelist()}
    for r in required:
        if r not in names_lower:
            raise HTTPException(status_code=400, detail=f"zip 안에 '{r}' 파일이 없습니다. (필수: {required})")

    def read_xlsx_bytes(name_lower):
        return z.read(names_lower[name_lower])

    receipt_bytes = read_xlsx_bytes("receipt.xlsx")
    plnt_bytes    = read_xlsx_bytes("plnt.xlsx")
    wip_bytes     = read_xlsx_bytes("wip.xlsx")
    template_bytes= read_xlsx_bytes("template.xlsx")

    # 워크북 로드
    receipt_wb  = load_workbook(io.BytesIO(receipt_bytes), data_only=False)
    plnt_wb     = load_workbook(io.BytesIO(plnt_bytes), data_only=False)
    wip_wb      = load_workbook(io.BytesIO(wip_bytes), data_only=False)
    template_wb = load_workbook(io.BytesIO(template_bytes), data_only=False)

    # plnt 매핑
    plnt_map, plnt_meta = build_plnt_map(plnt_wb)

    # receipt 집계
    receipt_sums, unknown_cats, repl_meta = parse_receipt(receipt_wb, plnt_map)

    # wip 합계(P/Q/R)
    wip_ws = wip_wb.active
    wip_sums = {
        "in":  sum_column_numeric(wip_ws, WIP_COLS["in"], start_row=1),
        "out": sum_column_numeric(wip_ws, WIP_COLS["out"], start_row=1),
        "end": sum_column_numeric(wip_ws, WIP_COLS["end"], start_row=1),
    }

    # template에 값/수식 쓰기
    write_values_and_formulas(template_wb, receipt_sums, wip_sums)

    # 저장
    out = io.BytesIO()
    template_wb.save(out)
    out.seek(0)

    # 간단 검증 요약(응답 헤더로도 가능하지만, 우선 JSON이 아닌 파일 응답이므로 생략)
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="final_close_report.xlsx"'}
    )
