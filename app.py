import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from copy import copy as cp
from io import BytesIO
import re

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Phân tích tồn kho – Giày Tuấn",
    page_icon="👟",
    layout="wide"
)

st.title("👟 Phân tích tồn kho – Đề xuất điều chuyển & nhập thêm")
st.caption("Upload file **Chi tiết số lượng nhập xuất tồn kho** (gồm sheet `6_thang` và `3_thang`) để phân tích.")

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
SIZE_DIGITS = ['5','6','7','8','9','0','1','2','3','4']
SIZE_MAP    = {'5':35,'6':36,'7':37,'8':38,'9':39,'0':40,'1':41,'2':42,'3':43,'4':44}
C_BLUE  = '1F3864'; C_RED = 'C00000'; C_GREEN = '375623'

DANH_GIA_COLOR = {
    'BÁN CHẠY':    ('C6EFCE','006100'),
    'BÁN KHÁ':     ('DDEBF7','1F497D'),
    'BÁN VỪA':     ('FFEB9C','9C5700'),
    'BÁN CHẬM':    ('FCE4D6','833C00'),
    'ĐIỀU CHUYỂN': ('FFD9D9','C00000'),
}

thin   = Side(style='thin', color='AAAAAA')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def F(bold=False, size=9, color='000000'):
    return Font(bold=bold, size=size, color=color, name='Arial')
def FILL(h): return PatternFill('solid', fgColor=h)
def AL(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def cw(ws, col, w):
    ws.column_dimensions[get_column_letter(col)].width = w
def title_row(ws, row, c1, c2, text, bg, fg='FFFFFF', size=12, rh=30):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    ws.row_dimensions[row].height = rh
    c = ws.cell(row=row, column=c1, value=text)
    c.font = Font(bold=True, size=size, color=fg, name='Arial')
    c.fill = FILL(bg); c.alignment = AL(wrap=True)
def hdr(ws, r, c, v, bg, fg='FFFFFF'):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(bold=True, size=9, color=fg, name='Arial')
    cell.fill = FILL(bg); cell.alignment = AL(wrap=True); cell.border = BORDER
    return cell
def dc(ws, r, c, v, bg='FFFFFF', bold=False, color='000000', al='center'):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(bold=bold, size=9, color=color, name='Arial')
    cell.fill = FILL(bg); cell.alignment = AL(h=al); cell.border = BORDER
    return cell

# ─────────────────────────────────────────────────────────────
# PARSE SKU
# ─────────────────────────────────────────────────────────────
P_STD = re.compile(r'^([A-Z]+)([A-Z]\d)(-/-)([A-Z0-9]+)(\d)$')
P_SFX = re.compile(r'^([A-Z]+)([A-Z]\d)(?:HM|KG)-([A-Z0-9]+)(\d)$')

def parse_sku(sku):
    s = str(sku).strip()
    m = P_STD.match(s)
    if m: px,_,_,cd,sd = m.groups(); return f'{px}-/-{cd}', sd
    m = P_SFX.match(s)
    if m: px,_,cd,sd = m.groups(); return f'{px}-/-{cd}', sd
    return s[:-1], s[-1]

def ten_mau(ten):
    t = re.sub(r'-màu\s+\S+', '', str(ten))
    t = re.sub(r'-[Ss]ize\s*\S+', '', t)
    return t.strip()

# ─────────────────────────────────────────────────────────────
# CORE ANALYSIS
# ─────────────────────────────────────────────────────────────
COLS = ['ma_sku','ten_hang','dvt','nhom',
        'ton_dau','nhap_tong','nhap_mua','nhap_dc','nhap_tralai',
        'nhap_ck','nhap_kiemke','nhap_khac',
        'xuat_tong','xuat_ban','xuat_dc','xuat_tralai',
        'xuat_ck','xuat_kiemke','xuat_huy','xuat_khac','ton_cuoi']

def load_sheet(file_bytes, sheet_name):
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
    d  = df.iloc[9:].copy(); d.columns = COLS
    d  = d[d['ma_sku'].notna() & (d['ma_sku'] != 'Mã SKU')].copy()
    for c in ['ton_dau','nhap_tong','xuat_ban','ton_cuoi']:
        d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0).astype(int)
    parsed = d['ma_sku'].apply(lambda x: pd.Series(parse_sku(x), index=['ma_hang','size_digit']))
    d['ma_hang']    = parsed['ma_hang']
    d['size_digit'] = parsed['size_digit']
    d['ten_mau']    = d['ten_hang'].apply(ten_mau)
    return d

def agg_mh(df):
    return df.groupby(['nhom','ma_hang']).agg(
        ten_mau = ('ten_mau',   'first'),
        ban     = ('xuat_ban',  'sum'),
        ton     = ('ton_cuoi',  'sum'),
        ton_dau = ('ton_dau',   'sum'),
        nhap    = ('nhap_tong', 'sum'),
    ).reset_index()

def st_val(ban, ton):
    d = ban + ton
    return round(ban / d, 4) if d > 0 else 0.0

def vong_quay(ban, thang, ton):
    if ton <= 0: return 0.0
    return round((ban / thang) / ton, 4)

def danh_gia_fn(row):
    ton, ban3 = row['ton_3t'], row['ban_3t']
    st3, st6   = row['st3m'],  row['st6m']
    vq3        = row['vq3m']

    GOI_Y = {
        'BÁN CHẠY':    'Cân nhắc nhập thêm',
        'BÁN KHÁ':     'Theo dõi – nhập khi gần hết',
        'BÁN VỪA':     'Duy trì mức tồn hiện tại',
        'BÁN CHẬM':    'Xem xét điều chuyển',
        'ĐIỀU CHUYỂN': 'Ưu tiên điều chuyển đi',
    }

    if ton > 0 and ban3 == 0:
        dg = 'ĐIỀU CHUYỂN'
        cb = 'Đã bán trong 6T nhưng ngưng 3T' if st6 >= 0.25 else 'Bán kém cả 3T và 6T'
        return dg, cb, GOI_Y[dg]

    trend = row['trend']
    cb = ('📈 Đang tăng tốt' if trend >= 0.15 else
          '📉 Đang giảm' if trend <= -0.15 else '')

    if (st6 >= 0.70 and st3 >= 0.70) or vq3 >= 1.0: dg = 'BÁN CHẠY'
    elif (st6 >= 0.55 and st3 >= 0.50) or vq3 >= 0.5: dg = 'BÁN KHÁ'
    elif (st6 >= 0.35 and st3 >= 0.30) or vq3 >= 0.2: dg = 'BÁN VỪA'
    else: dg = 'BÁN CHẬM'
    return dg, cb, GOI_Y[dg]

def run_analysis(file_bytes):
    TODAY = date.today().strftime('%d/%m/%Y')

    # Load
    df6r = load_sheet(file_bytes, '6_thang')
    df3r = load_sheet(file_bytes, '3_thang')

    # Aggregate
    a6 = agg_mh(df6r).rename(columns={'ban':'ban_6t','ton':'ton_6t','ton_dau':'ton_dau_6t','nhap':'nhap_6t'})
    a3 = agg_mh(df3r).rename(columns={'ban':'ban_3t','ton':'ton_3t','ton_dau':'ton_dau_3t','nhap':'nhap_3t'})
    agg = a6.merge(a3[['nhom','ma_hang','ban_3t','ton_3t','ton_dau_3t','nhap_3t']],
                   on=['nhom','ma_hang'], how='outer')
    agg['ten_mau'] = agg['ten_mau'].fillna('')
    for c in ['ban_6t','ton_6t','ban_3t','ton_3t','ton_dau_6t','nhap_6t','ton_dau_3t','nhap_3t']:
        agg[c] = agg[c].fillna(0).astype(int)

    # Indicators
    agg['st3m']  = agg.apply(lambda r: st_val(r['ban_3t'], r['ton_3t']), axis=1)
    agg['st6m']  = agg.apply(lambda r: st_val(r['ban_6t'], r['ton_6t']), axis=1)
    agg['vq3m']  = agg.apply(lambda r: vong_quay(r['ban_3t'], 3, r['ton_3t']), axis=1)
    agg['vq6m']  = agg.apply(lambda r: vong_quay(r['ban_6t'], 6, r['ton_6t']), axis=1)
    agg['trend'] = (agg['st3m'] - agg['st6m']).round(4)

    results = agg.apply(lambda r: pd.Series(danh_gia_fn(r), index=['danh_gia','canh_bao','goi_y']), axis=1)
    agg = pd.concat([agg, results], axis=1)
    agg = agg.sort_values(['nhom','ma_hang']).reset_index(drop=True)

    # Size priority
    size_df = df6r[['nhom','ma_hang','ma_sku','size_digit','xuat_ban','ton_cuoi']].copy()
    size_df = size_df.rename(columns={'xuat_ban':'ban_6t','ton_cuoi':'ton'})
    size_max = size_df.groupby('ma_hang')['ban_6t'].max().rename('max_ban')
    size_df  = size_df.merge(size_max, on='ma_hang', how='left')
    size_df['size_index'] = size_df.apply(
        lambda r: round(r['ban_6t']/r['max_ban'],4) if r['max_ban']>0 else 0.0, axis=1)
    size_df['xep_loai'] = size_df['size_index'].apply(
        lambda si: 'Size chạy' if si>=0.75 else ('Size TB' if si>=0.40 else ('Size chậm' if si>0 else 'Không bán')))
    size_df = size_df.sort_values(['nhom','ma_hang','size_digit'])

    def ton_size_pivot(df, mahang_list):
        sub = df[df['ma_hang'].isin(mahang_list)]
        pv  = (sub.groupby(['ma_hang','size_digit'])['ton_cuoi']
                   .sum().unstack(fill_value=0)
                   .reindex(columns=SIZE_DIGITS, fill_value=0))
        pv.columns = [f'ton_{c}' for c in SIZE_DIGITS]
        return pv

    def si_pivot_fn(mahang_list):
        sub = size_df[size_df['ma_hang'].isin(mahang_list)]
        if sub.empty: return pd.DataFrame()
        pv  = (sub.groupby(['ma_hang','size_digit'])['size_index']
                   .max().unstack(fill_value=0)
                   .reindex(columns=SIZE_DIGITS, fill_value=0))
        pv.columns = [f'si_{c}' for c in SIZE_DIGITS]
        return pv

    def calc_bs_row(ma_hang, ton_pv, si_pv):
        result = {}
        for sd in SIZE_DIGITS:
            ton = int(ton_pv.loc[ma_hang, f'ton_{sd}']) if ma_hang in ton_pv.index else 0
            si  = float(si_pv.loc[ma_hang, f'si_{sd}']) if (not si_pv.empty and ma_hang in si_pv.index) else 0.0
            if ton == 0 and si == 0:
                result[f'bs_{sd}'] = 0; continue
            min_bs = max(0, 2 - ton)
            extra  = 2 if si>=0.75 else (1 if si>=0.40 else 0)
            result[f'bs_{sd}'] = min_bs + extra
        return result

    def build_sheet_data(mahang_list):
        if not mahang_list:
            return pd.DataFrame()
        dc6 = df6r[df6r['ma_hang'].isin(mahang_list)].copy()
        dc3 = df3r[df3r['ma_hang'].isin(mahang_list)].copy()

        def agg_dc(df):
            return df.groupby('ma_hang').agg(
                ten_mau  = ('ten_mau',   'first'),
                nhom     = ('nhom',      'first'),
                ton_dau  = ('ton_dau',   'sum'),
                nhap     = ('nhap_tong', 'sum'),
                xuat_ban = ('xuat_ban',  'sum'),
                ton_cuoi = ('ton_cuoi',  'sum'),
            ).reset_index()

        a6x = agg_dc(dc6).rename(columns={'ton_dau':'ton_dau_6t','nhap':'nhap_6t',
                                           'xuat_ban':'xuat_6t','ton_cuoi':'ton_cuoi_6t'})
        a3x = agg_dc(dc3).rename(columns={'ton_dau':'ton_dau_3t','nhap':'nhap_3t',
                                           'xuat_ban':'xuat_3t','ton_cuoi':'ton_cuoi_3t'})
        final = a6x.merge(a3x[['ma_hang','ton_dau_3t','nhap_3t','xuat_3t','ton_cuoi_3t']],
                          on='ma_hang', how='left')
        for c in ['ton_dau_3t','nhap_3t','xuat_3t','ton_cuoi_3t']:
            final[c] = final[c].fillna(0).astype(int)

        ton_pv = ton_size_pivot(dc6, mahang_list)
        si_pv  = si_pivot_fn(mahang_list)
        ton_pv_r = ton_pv.reset_index()
        final = final.merge(ton_pv_r, on='ma_hang', how='left')
        for sd in SIZE_DIGITS:
            col = f'ton_{sd}'
            if col not in final.columns: final[col] = 0
            final[col] = final[col].fillna(0).astype(int)

        bs_rows = [calc_bs_row(mh, ton_pv, si_pv) for mh in final['ma_hang']]
        bs_df   = pd.DataFrame(bs_rows, index=final.index)
        final   = pd.concat([final, bs_df], axis=1)
        for sd in SIZE_DIGITS:
            col = f'bs_{sd}'
            if col not in final.columns: final[col] = 0
            final[col] = final[col].fillna(0).astype(int)
        final['bs_tong'] = final[[f'bs_{sd}' for sd in SIZE_DIGITS]].sum(axis=1)
        return final.sort_values('ton_cuoi_6t', ascending=False).reset_index(drop=True)

    mh_dc = agg[agg['danh_gia']=='ĐIỀU CHUYỂN']['ma_hang'].tolist()
    mh_nt = agg[agg['danh_gia']=='BÁN CHẠY']['ma_hang'].tolist()
    data_dc = build_sheet_data(mh_dc)
    data_nt = build_sheet_data(mh_nt)

    # ── Build Excel output ────────────────────────────────────
    wb = openpyxl.Workbook()

    # Sheet 1: BAN_3_THANG_TONGHOP
    ws1 = wb.active; ws1.title = 'BAN_3_THANG_TONGHOP'
    title_row(ws1,1,1,6,'BÁN 3 THÁNG – TỔNG HỢP THEO MÃ HÀNG',C_BLUE)
    ws1.merge_cells('A2:F2')
    ws1.cell(2,1,f'Ngày phân tích: {TODAY}').font=F(size=9,color='555555')
    for i,(h,w) in enumerate(zip(['Nhóm','Mã hàng','Tên hàng hóa','Bán 3T','Tồn 3T','ST3M'],
                                  [15,22,48,10,10,10]),1):
        hdr(ws1,3,i,h,C_BLUE); cw(ws1,i,w)
    ws1.row_dimensions[3].height=22
    for ri,row in enumerate(agg.itertuples(),1):
        er=ri+3; bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        for ci,v in enumerate([row.nhom,row.ma_hang,row.ten_mau,row.ban_3t,row.ton_3t,f'{row.st3m:.1%}'],1):
            dc(ws1,er,ci,v,bg=bg,al='center' if ci in(4,5,6) else 'left')
        ws1.row_dimensions[er].height=15
    ws1.freeze_panes='A4'

    # Sheet 2: BAN_6_THANG_TONGHOP
    ws2=wb.create_sheet('BAN_6_THANG_TONGHOP')
    title_row(ws2,1,1,6,'BÁN 6 THÁNG – TỔNG HỢP THEO MÃ HÀNG',C_BLUE)
    ws2.merge_cells('A2:F2')
    ws2.cell(2,1,f'Ngày phân tích: {TODAY}').font=F(size=9,color='555555')
    for i,(h,w) in enumerate(zip(['Nhóm','Mã hàng','Tên hàng hóa','Bán 6T','Tồn 6T','ST6M'],
                                  [15,22,48,10,10,10]),1):
        hdr(ws2,3,i,h,C_BLUE); cw(ws2,i,w)
    ws2.row_dimensions[3].height=22
    for ri,row in enumerate(agg.itertuples(),1):
        er=ri+3; bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        for ci,v in enumerate([row.nhom,row.ma_hang,row.ten_mau,row.ban_6t,row.ton_6t,f'{row.st6m:.1%}'],1):
            dc(ws2,er,ci,v,bg=bg,al='center' if ci in(4,5,6) else 'left')
        ws2.row_dimensions[er].height=15
    ws2.freeze_panes='A4'

    # Sheet 3: DANH_GIA_TIEU_THU
    ws3=wb.create_sheet('DANH_GIA_TIEU_THU')
    title_row(ws3,1,1,14,'ĐÁNH GIÁ TIÊU THỤ – ST + VÒNG QUAY TỒN KHO',C_BLUE)
    ws3.merge_cells('A2:N2')
    ws3.cell(2,1,f'ST=Bán/(Bán+Tồn)  |  Vòng quay=(Bán TB/tháng)÷Tồn  |  Ngày: {TODAY}').font=F(size=9,color='555555')
    HDG=['Nhóm','Mã hàng','Tên hàng hóa','Bán 3T','Bán 6T','Tồn',
         'ST3M','ST6M','VQ 3T','VQ 6T','Trend','Đánh giá','Cảnh báo','Gợi ý']
    WDG=[15,22,45,9,9,9,8,8,9,9,9,14,35,25]
    for i,(h,w) in enumerate(zip(HDG,WDG),1):
        hdr(ws3,3,i,h,C_BLUE); cw(ws3,i,w)
    ws3.row_dimensions[3].height=25
    for ri,row in enumerate(agg.itertuples(),1):
        er=ri+3; dg=row.danh_gia
        bg_dg,fg_dg=DANH_GIA_COLOR.get(dg,('FFFFFF','000000'))
        bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        trend_s=f'+{row.trend:.1%}' if row.trend>=0 else f'{row.trend:.1%}'
        vals=[row.nhom,row.ma_hang,row.ten_mau,row.ban_3t,row.ban_6t,row.ton_3t,
              f'{row.st3m:.1%}',f'{row.st6m:.1%}',f'{row.vq3m:.2f}',f'{row.vq6m:.2f}',trend_s,
              dg,row.canh_bao,row.goi_y]
        for ci,v in enumerate(vals,1):
            if ci==12: dc(ws3,er,ci,v,bg=bg_dg,bold=True,color=fg_dg)
            elif ci==11:
                cbg='E2EFDA' if row.trend>=0.15 else ('FFD9D9' if row.trend<=-0.15 else bg)
                dc(ws3,er,ci,v,bg=cbg)
            else: dc(ws3,er,ci,v,bg=bg,al='center' if ci in(4,5,6,7,8,9,10) else 'left')
        ws3.row_dimensions[er].height=15
    ws3.freeze_panes='A4'
    ws3.auto_filter.ref=f'A3:N{len(agg)+3}'

    # Sheet 4: SIZE_PRIORITY
    ws4=wb.create_sheet('SIZE_PRIORITY')
    title_row(ws4,1,1,8,'PHÂN TÍCH SIZE – CHỈ SỐ ƯU TIÊN',C_BLUE)
    ws4.merge_cells('A2:H2')
    ws4.cell(2,1,'Size Index = Bán 6T size / Max bán 6T các size cùng mã hàng').font=F(size=9,color='555555')
    for i,(h,w) in enumerate(zip(['Nhóm','Mã hàng','SKU','Size','Bán 6T','Tồn','Size Index','Xếp loại'],
                                   [15,22,24,6,9,9,11,14]),1):
        hdr(ws4,3,i,h,C_BLUE); cw(ws4,i,w)
    ws4.row_dimensions[3].height=22
    XLC={'Size chạy':('C6EFCE','006100'),'Size TB':('DDEBF7','1F497D'),
         'Size chậm':('FCE4D6','833C00'),'Không bán':('F2F2F2','888888')}
    for ri,row in enumerate(size_df.itertuples(),1):
        er=ri+3; bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        xl=row.xep_loai; bg_xl,fg_xl=XLC.get(xl,('FFFFFF','000000'))
        sz=SIZE_MAP.get(row.size_digit, row.size_digit)
        for ci,v in enumerate([row.nhom,row.ma_hang,row.ma_sku,sz,row.ban_6t,row.ton,f'{row.size_index:.1%}',xl],1):
            if ci==8: dc(ws4,er,ci,v,bg=bg_xl,bold=True,color=fg_xl)
            else: dc(ws4,er,ci,v,bg=bg,al='center' if ci in(4,5,6,7) else 'left')
        ws4.row_dimensions[er].height=15
    ws4.freeze_panes='A4'
    ws4.auto_filter.ref=f'A3:H{len(size_df)+3}'

    # Sheet 5 & 6: Phiếu (không cần file mẫu gốc – tự build layout)
    TON_COL = {sd: 11+i for i,sd in enumerate(SIZE_DIGITS)}
    BS_COL  = {sd: 21+i for i,sd in enumerate(SIZE_DIGITS)}
    DATA_ROW = 7

    def build_phieu_sheet(wb, sheet_name, title_text, bg_title, df_data):
        ws = wb.create_sheet(sheet_name)
        title_row(ws,1,1,31,title_text,bg_title,rh=30)
        ws.merge_cells('A2:AE2')
        ws.cell(2,1,f'Ngày lập: {TODAY} | Cửa hàng: Giày Tuấn').font=F(size=9,color='555555')

        # Header row 3 (nhóm cột)
        for txt,c1,c2,bg in [
            ('STT',1,1,C_BLUE),('Mã hàng',2,2,C_BLUE),
            ('N-X 3 tháng',3,6,bg_title),('N-X 6 tháng',7,10,C_BLUE),
            ('Chi tiết tồn kho theo size',11,20,'2F5496'),
            ('Bổ sung / Đề xuất theo size',21,30,'375623'),
            ('Tổng',31,31,'7F6000')]:
            ws.merge_cells(start_row=3,start_column=c1,end_row=3,end_column=c2)
            cell=ws.cell(3,c1,txt)
            cell.font=Font(bold=True,size=9,color='FFFFFF',name='Arial')
            cell.fill=FILL(bg); cell.alignment=AL(wrap=True); cell.border=BORDER

        # Header row 4 (chi tiết)
        h4=['STT','Mã hàng','Tồn đầu','Nhập','Xuất','Tồn cuối',
            'Tồn đầu','Nhập','Xuất','Tồn cuối']
        h4 += [str(SIZE_MAP[s]) for s in SIZE_DIGITS]  # size 35-44
        h4 += [str(SIZE_MAP[s]) for s in SIZE_DIGITS]  # bs size 35-44
        h4 += ['Tổng']
        widths=[5,20]+[9]*8+[7]*10+[7]*10+[8]
        for ci,(h,w) in enumerate(zip(h4,widths),1):
            hdr(ws,4,ci,h,bg_title if ci<=10 else ('2F5496' if ci<=20 else ('375623' if ci<=30 else '7F6000')))
            cw(ws,ci,w)
        ws.row_dimensions[3].height=20; ws.row_dimensions[4].height=22

        if df_data.empty:
            ws.cell(DATA_ROW,1,'Không có dữ liệu').font=F(size=10,color='888888')
            return ws

        for ri,row in enumerate(df_data.itertuples(),0):
            er=DATA_ROW+ri; bg='FFFFFF' if ri%2==0 else 'FFF9F0'
            dc(ws,er,1,ri+1,bg=bg)
            dc(ws,er,2,row.ma_hang,bg=bg,al='left')
            dc(ws,er,3,row.ton_dau_3t,bg=bg)
            dc(ws,er,4,row.nhap_3t,bg=bg)
            dc(ws,er,5,row.xuat_3t,bg=bg)
            dc(ws,er,6,row.ton_cuoi_3t,bg=bg,bold=True)
            dc(ws,er,7,row.ton_dau_6t,bg=bg)
            dc(ws,er,8,row.nhap_6t,bg=bg)
            dc(ws,er,9,row.xuat_6t,bg=bg)
            dc(ws,er,10,row.ton_cuoi_6t,bg=bg,bold=True)
            for sd,col in TON_COL.items():
                val=getattr(row,f'ton_{sd}',0)
                cbg='FFD9D9' if val>0 else bg
                dc(ws,er,col,val,bg=cbg,bold=(val>0))
            for sd,col in BS_COL.items():
                val=getattr(row,f'bs_{sd}',0)
                cbg='E2EFDA' if val>0 else bg
                dc(ws,er,col,val,bg=cbg,bold=(val>0),color=(C_GREEN if val>0 else '000000'))
            dc(ws,er,31,row.bs_tong,bg='FFF2CC',bold=True,color=(C_RED if row.bs_tong>0 else '000000'))
            ws.row_dimensions[er].height=16

        ws.freeze_panes='A5'
        return ws

    build_phieu_sheet(wb,'DE_XUAT_DIEU_CHUYEN',
                      '📋 PHIẾU ĐỀ XUẤT ĐIỀU CHUYỂN – Hàng tồn kho không bán 3T',
                      C_RED, data_dc)
    build_phieu_sheet(wb,'DE_XUAT_NHAP_THEM',
                      '📋 PHIẾU ĐỀ XUẤT NHẬP THÊM – Hàng bán chạy cần bổ sung',
                      C_GREEN, data_nt)

    # Save to bytes
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, agg

# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "📂 Upload file Chi tiết số lượng nhập xuất tồn kho (.xlsx)",
    type=["xlsx","xls"],
    help="File phải có 2 sheet: `6_thang` và `3_thang`"
)

if uploaded:
    st.success(f"✅ Đã upload: **{uploaded.name}**")

    col1, col2 = st.columns([1,3])
    with col1:
        run_btn = st.button("🚀 Chạy phân tích", use_container_width=True, type="primary")

    if run_btn:
        with st.spinner("Đang phân tích dữ liệu..."):
            try:
                file_bytes = BytesIO(uploaded.read())
                result_buf, agg = run_analysis(file_bytes)

                # Metrics
                count = agg['danh_gia'].value_counts().to_dict()
                st.markdown("---")
                st.subheader("📊 Kết quả phân tích")
                m1,m2,m3,m4,m5 = st.columns(5)
                m1.metric("🔥 Bán chạy",    count.get('BÁN CHẠY',0))
                m2.metric("✅ Bán khá",     count.get('BÁN KHÁ',0))
                m3.metric("🔶 Bán vừa",     count.get('BÁN VỪA',0))
                m4.metric("🟠 Bán chậm",    count.get('BÁN CHẬM',0))
                m5.metric("🔴 Điều chuyển", count.get('ĐIỀU CHUYỂN',0))

                # Preview top hàng cần điều chuyển
                st.markdown("#### Top 10 mã hàng cần điều chuyển (tồn nhiều, không bán 3T)")
                dc_preview = agg[agg['danh_gia']=='ĐIỀU CHUYỂN'][
                    ['nhom','ma_hang','ten_mau','ban_3t','ban_6t','ton_3t','st6m','canh_bao']
                ].rename(columns={'nhom':'Nhóm','ma_hang':'Mã hàng','ten_mau':'Tên',
                                   'ban_3t':'Bán 3T','ban_6t':'Bán 6T','ton_3t':'Tồn',
                                   'st6m':'ST6M','canh_bao':'Cảnh báo'})
                dc_preview['ST6M'] = dc_preview['ST6M'].map('{:.1%}'.format)
                st.dataframe(dc_preview.head(10), use_container_width=True, hide_index=True)

                # Preview top hàng bán chạy
                st.markdown("#### Top 10 mã hàng bán chạy (cần nhập thêm)")
                nt_preview = agg[agg['danh_gia']=='BÁN CHẠY'][
                    ['nhom','ma_hang','ten_mau','ban_3t','ban_6t','ton_3t','st3m','st6m','vq3m']
                ].rename(columns={'nhom':'Nhóm','ma_hang':'Mã hàng','ten_mau':'Tên',
                                   'ban_3t':'Bán 3T','ban_6t':'Bán 6T','ton_3t':'Tồn',
                                   'st3m':'ST3M','st6m':'ST6M','vq3m':'Vòng quay 3T'})
                nt_preview['ST3M'] = nt_preview['ST3M'].map('{:.1%}'.format)
                nt_preview['ST6M'] = nt_preview['ST6M'].map('{:.1%}'.format)
                nt_preview['Vòng quay 3T'] = nt_preview['Vòng quay 3T'].map('{:.2f}'.format)
                st.dataframe(nt_preview.sort_values('Bán 3T', ascending=False).head(10),
                             use_container_width=True, hide_index=True)

                # Download
                st.markdown("---")
                fname = f"Phan_Tich_Ton_Kho_{date.today().strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    label="⬇️ Tải xuống file kết quả (.xlsx)",
                    data=result_buf,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                st.caption(f"File gồm 6 sheet: BAN_3_THANG_TONGHOP | BAN_6_THANG_TONGHOP | DANH_GIA_TIEU_THU | SIZE_PRIORITY | DE_XUAT_DIEU_CHUYEN | DE_XUAT_NHAP_THEM")

            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý: {e}")
                st.exception(e)
else:
    st.info("👆 Vui lòng upload file Excel để bắt đầu phân tích.")
    with st.expander("ℹ️ Hướng dẫn sử dụng"):
        st.markdown("""
**Yêu cầu file upload:**
- Định dạng `.xlsx`
- Phải có **2 sheet**: `6_thang` và `3_thang`
- Dữ liệu bắt đầu từ **dòng 10** (header ở dòng 9)
- Gồm các cột: Mã SKU, Tên hàng, ĐVT, Nhóm, Tồn đầu, Nhập, Xuất, Tồn cuối...

**Output gồm 6 sheet:**

| Sheet | Nội dung |
|---|---|
| BAN_3_THANG_TONGHOP | Tổng hợp bán 3 tháng theo mã hàng |
| BAN_6_THANG_TONGHOP | Tổng hợp bán 6 tháng theo mã hàng |
| DANH_GIA_TIEU_THU | Đánh giá + ST + Vòng quay + Trend |
| SIZE_PRIORITY | Size Index theo từng SKU |
| DE_XUAT_DIEU_CHUYEN | Phiếu hàng tồn không bán 3T → điều chuyển đi |
| DE_XUAT_NHAP_THEM | Phiếu hàng bán chạy → nhập bổ sung |

**Công thức đánh giá:**
- `ST = Bán / (Bán + Tồn)` — tỷ lệ bán ra
- `Vòng quay = (Bán TB/tháng) ÷ Tồn` — tốc độ luân chuyển
- **BÁN CHẠY**: ST6≥70% & ST3≥70% hoặc VQ3≥1.0
- **ĐIỀU CHUYỂN**: Còn tồn nhưng 3 tháng không bán
        """)
