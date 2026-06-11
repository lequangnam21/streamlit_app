import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from copy import copy as cp
from io import BytesIO
import re

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Phân tích tồn kho – Giày Tuấn", page_icon="👟", layout="wide")
st.title("👟 Phân tích tồn kho – Đề xuất điều chuyển & nhập thêm")
st.caption("Upload file **Chi tiết số lượng nhập xuất tồn kho** (gồm sheet `6_thang` và `3_thang`) để phân tích.")

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
SIZE_DIGITS = ['5','6','7','8','9','0','1','2','3','4']
SIZE_MAP    = {'5':35,'6':36,'7':37,'8':38,'9':39,'0':40,'1':41,'2':42,'3':43,'4':44}
C_BLUE = '1F3864'; C_RED = 'C00000'; C_GREEN = '375623'
TOP_N  = 10   # top mã mỗi nhóm

DANH_GIA_COLOR = {
    'BÁN CHẠY':    ('C6EFCE','006100'),
    'BÁN KHÁ':     ('DDEBF7','1F497D'),
    'BÁN VỪA':     ('FFEB9C','9C5700'),
    'BÁN CHẬM':    ('FCE4D6','833C00'),
    'ĐIỀU CHUYỂN': ('FFD9D9','C00000'),
    'HÀNG MỚI CHƯA ĐÁNH GIÁ': ('E2F0D9','548235'),
}

# 17 màu xen kẽ cho nhóm hàng (đậm/nhạt luân phiên)
GROUP_COLORS = [
    ('D9E1F2','1F3864'),('FCE4D6','833C00'),('E2EFDA','375623'),('FFF2CC','7F6000'),
    ('DDEBF7','1F497D'),('FFD9D9','C00000'),('EBF1DE','4F6228'),('FFF9C4','5C4A00'),
    ('D6E4F0','1A4971'),('FAD7A0','784212'),('D5F5E3','1E8449'),('FDEDEC','922B21'),
    ('E8DAEF','6C3483'),('D1F2EB','117A65'),('FDEBD0','784212'),('EAF2F8','1B4F72'),
    ('F9EBEA','922B21'),
]

thin   = Side(style='thin',   color='AAAAAA')
thick  = Side(style='medium', color='666666')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
BORDER_THICK_TOP = Border(left=thin, right=thin, top=thick, bottom=thin)

def F(bold=False, size=9, color='000000'):
    return Font(bold=bold, size=size, color=color, name='Arial')
def FILL(h): return PatternFill('solid', fgColor=h)
def AL(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def cw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w
def title_row(ws, row, c1, c2, text, bg, fg='FFFFFF', size=12, rh=30):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    ws.row_dimensions[row].height = rh
    c = ws.cell(row=row, column=c1, value=text)
    c.font = Font(bold=True, size=size, color=fg, name='Arial')
    c.fill = FILL(bg); c.alignment = AL(wrap=True)
def hdr(ws, r, c, v, bg, fg='FFFFFF', wrap=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(bold=True, size=9, color=fg, name='Arial')
    cell.fill = FILL(bg); cell.alignment = AL(wrap=wrap); cell.border = BORDER
    return cell
def dc(ws, r, c, v, bg='FFFFFF', bold=False, color='000000', al='center', thick_top=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(bold=bold, size=9, color=color, name='Arial')
    cell.fill = FILL(bg); cell.alignment = AL(h=al)
    cell.border = BORDER_THICK_TOP if thick_top else BORDER
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
    t = re.sub(r'-màu\s+\S+', '', str(ten), flags=re.I)
    t = re.sub(r'-[Ss]ize\s*\S+', '', t, flags=re.I)
    t = re.sub(r'\s+(3[5-9]|4[0-4])$', '', t)
    return t.strip()

# ─────────────────────────────────────────────────────────────
# LOAD & PARSE
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
        ten_mau = ('ten_mau', 'first'),
        ban     = ('xuat_ban',  'sum'),
        ton     = ('ton_cuoi',  'sum'),
        ton_dau = ('ton_dau',   'sum'),
        nhap    = ('nhap_tong', 'sum'),
    ).reset_index()

def st_val(ban, ton):
    d = ban + ton; return round(ban/d, 4) if d > 0 else 0.0

def vong_quay(ban, thang, ton):
    if ton <= 0: return 0.0
    return round((ban/thang)/ton, 4)

# ─────────────────────────────────────────────────────────────
# ĐÁNH GIÁ
# ─────────────────────────────────────────────────────────────
def danh_gia_fn(row):

    ton = row['ton_3t']
    ban3 = row['ban_3t']
    ban6 = row['ban_6t']

    st3 = row['st3m']
    st6 = row['st6m']
    mos = row['mos3m']

    ton_dau_3t = row['ton_dau_3t']
    nhap_3t = row['nhap_3t']

    GOI_Y = {
        'BÁN CHẠY': 'Cân nhắc nhập thêm',
        'BÁN KHÁ': 'Theo dõi – nhập khi gần hết',
        'BÁN VỪA': 'Duy trì mức tồn hiện tại',
        'BÁN CHẬM': 'Theo dõi thêm',
        'ĐIỀU CHUYỂN': 'Ưu tiên điều chuyển đi',
        'HÀNG MỚI CHƯA ĐÁNH GIÁ': 'Chưa đủ dữ liệu đánh giá',
    }

    if ton_dau_3t == 0 and nhap_3t > 0:
        return (
            'HÀNG MỚI CHƯA ĐÁNH GIÁ',
            'Mới nhập trong 3 tháng',
            GOI_Y['HÀNG MỚI CHƯA ĐÁNH GIÁ']
        )

    if (
        ton > 0
        and ban3 == 0
        and ban6 <= 1
        and st6 < 0.30
        and ton_dau_3t > 0
    ):
        return (
            'ĐIỀU CHUYỂN',
            'Tồn lâu, không bán',
            GOI_Y['ĐIỀU CHUYỂN']
        )

    trend = row['trend']
    cb = '📈 Đang tăng tốt' if trend >= 0.15 else ('📉 Đang giảm' if trend <= -0.15 else '')

    if st3 >= 0.60 and mos < 1.5:
        dg = 'BÁN CHẠY'
    elif st3 >= 0.45 and mos < 3:
        dg = 'BÁN KHÁ'
    elif 3 <= mos <= 6:
        dg = 'BÁN VỪA'
    else:
        dg = 'BÁN CHẬM'

    return dg, cb, GOI_Y[dg]

# ─────────────────────────────────────────────────────────────
# MAIN ANALYSIS
# ─────────────────────────────────────────────────────────────
def run_analysis(file_bytes):
    TODAY = date.today().strftime('%d/%m/%Y')

    df6r = load_sheet(file_bytes, '6_thang')
    df3r = load_sheet(file_bytes, '3_thang')

    a6 = agg_mh(df6r).rename(columns={'ban':'ban_6t','ton':'ton_6t','ton_dau':'ton_dau_6t','nhap':'nhap_6t'})
    a3 = agg_mh(df3r).rename(columns={'ban':'ban_3t','ton':'ton_3t','ton_dau':'ton_dau_3t','nhap':'nhap_3t'})
    agg = a6.merge(a3[['nhom','ma_hang','ban_3t','ton_3t','ton_dau_3t','nhap_3t']],
                   on=['nhom','ma_hang'], how='outer')
    agg['ten_mau'] = agg['ten_mau'].fillna('')
    for c in ['ban_6t','ton_6t','ban_3t','ton_3t','ton_dau_6t','nhap_6t','ton_dau_3t','nhap_3t']:
        agg[c] = agg[c].fillna(0).astype(int)

    agg['st3m']  = agg.apply(lambda r: st_val(r['ban_3t'], r['ton_3t']), axis=1)
    agg['st6m']  = agg.apply(lambda r: st_val(r['ban_6t'], r['ton_6t']), axis=1)
    agg['vq3m']  = agg.apply(lambda r: vong_quay(r['ban_3t'], 3, r['ton_3t']), axis=1)
    agg['vq6m']  = agg.apply(lambda r: vong_quay(r['ban_6t'], 6, r['ton_6t']), axis=1)
    agg['trend'] = (agg['st3m'] - agg['st6m']).round(4)
    agg['mos3m'] = agg.apply(
        lambda r: round(r['ton_3t']/(r['ban_3t']/3), 2) if r['ban_3t'] > 0 else 999, axis=1)

    results = agg.apply(
        lambda r: pd.Series(danh_gia_fn(r), index=['danh_gia','canh_bao','goi_y']), axis=1)
    agg = pd.concat([agg, results], axis=1).reset_index(drop=True)

    # ── Size priority (gộp tất cả màu cùng size_digit) ────────
    size_df = (df6r
        .groupby(['nhom','ma_hang','size_digit'], as_index=False)
        .agg(ban_6t=('xuat_ban','sum'), ton=('ton_cuoi','sum')))
    size_max = size_df.groupby('ma_hang')['ban_6t'].max().rename('max_ban')
    size_df  = size_df.merge(size_max, on='ma_hang', how='left')
    size_df['size_index'] = size_df.apply(
        lambda r: round(r['ban_6t']/r['max_ban'],4) if r['max_ban']>0 else 0.0, axis=1)
    size_df['xep_loai'] = size_df['size_index'].apply(
        lambda si: 'Size chạy' if si>=0.75 else ('Size TB' if si>=0.40 else ('Size chậm' if si>0 else 'Không bán')))

    # ── Helper: pivot tồn / size_index theo size_digit ────────
    def ton_size_pivot(df_sku, mahang_list):
        sub = df_sku[df_sku['ma_hang'].isin(mahang_list)]
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

    # ── Bảng size hợp lệ của từng mã hàng (từ SKU thực tế 6T) ─
    # valid_sizes[ma_hang] = set của size_digit thực sự tồn tại
    valid_sizes = (df6r.groupby('ma_hang')['size_digit']
                       .apply(set).to_dict())

    def calc_dc_row(ma_hang, ton_pv):
        # Điều chuyển: chỉ các size thực tế của mã hàng đó
        valid = valid_sizes.get(ma_hang, set())
        result = {}
        for sd in SIZE_DIGITS:
            if sd not in valid:
                result[f'dc_{sd}'] = 0
            else:
                result[f'dc_{sd}'] = int(ton_pv.loc[ma_hang, f'ton_{sd}']) \
                    if ma_hang in ton_pv.index else 0
        return result

    def calc_bs_row(ma_hang, ton_pv, si_pv):
        # Bổ sung: chỉ các size thực tế của mã hàng đó
        valid = valid_sizes.get(ma_hang, set())
        result = {}
        for sd in SIZE_DIGITS:
            if sd not in valid:
                # Size này không có trong mã hàng → không đề xuất
                result[f'bs_{sd}'] = 0
                continue
            ton = int(ton_pv.loc[ma_hang, f'ton_{sd}']) if ma_hang in ton_pv.index else 0
            si  = float(si_pv.loc[ma_hang, f'si_{sd}']) if (not si_pv.empty and ma_hang in si_pv.index) else 0.0
            # Tối thiểu 2 đôi mỗi size hợp lệ
            min_bs = max(0, 2 - ton)
            # Thêm theo size_index (bán chạy → nhập thêm)
            extra  = 2 if si >= 0.75 else (1 if si >= 0.40 else 0)
            result[f'bs_{sd}'] = min_bs + extra
        return result

    def build_data(mahang_list, is_dc=False):
        if not mahang_list: return pd.DataFrame()
        d6 = df6r[df6r['ma_hang'].isin(mahang_list)].copy()
        d3 = df3r[df3r['ma_hang'].isin(mahang_list)].copy()

        def agg_dc(df):
            return df.groupby('ma_hang').agg(
                ten_mau  = ('ten_mau',   'first'),
                nhom     = ('nhom',      'first'),
                ton_dau  = ('ton_dau',   'sum'),
                nhap     = ('nhap_tong', 'sum'),
                xuat_ban = ('xuat_ban',  'sum'),
                ton_cuoi = ('ton_cuoi',  'sum'),
            ).reset_index()

        f6 = agg_dc(d6).rename(columns={'ton_dau':'ton_dau_6t','nhap':'nhap_6t',
                                          'xuat_ban':'xuat_6t','ton_cuoi':'ton_cuoi_6t'})
        f3 = agg_dc(d3).rename(columns={'ton_dau':'ton_dau_3t','nhap':'nhap_3t',
                                          'xuat_ban':'xuat_3t','ton_cuoi':'ton_cuoi_3t'})
        final = f6.merge(f3[['ma_hang','ton_dau_3t','nhap_3t','xuat_3t','ton_cuoi_3t']],
                         on='ma_hang', how='left')
        for c in ['ton_dau_3t','nhap_3t','xuat_3t','ton_cuoi_3t']:
            final[c] = final[c].fillna(0).astype(int)

        ton_pv   = ton_size_pivot(d6, mahang_list)
        si_pv    = si_pivot_fn(mahang_list)
        final    = final.merge(ton_pv.reset_index(), on='ma_hang', how='left')
        for sd in SIZE_DIGITS:
            final[f'ton_{sd}'] = final.get(f'ton_{sd}', 0).fillna(0).astype(int)

        if is_dc:
            rows = [calc_dc_row(mh, ton_pv) for mh in final['ma_hang']]
            key  = 'dc'
        else:
            rows = [calc_bs_row(mh, ton_pv, si_pv) for mh in final['ma_hang']]
            key  = 'bs'

        extra = pd.DataFrame(rows, index=final.index)
        final = pd.concat([final, extra], axis=1)
        for sd in SIZE_DIGITS:
            final[f'{key}_{sd}'] = final.get(f'{key}_{sd}', 0).fillna(0).astype(int)
        final[f'{key}_tong'] = final[[f'{key}_{sd}' for sd in SIZE_DIGITS]].sum(axis=1)
        return final

    # ── Lọc và chọn TOP N mỗi nhóm ───────────────────────────
    
    def is_model_2024_2026(ma_hang):
        s = str(ma_hang)
        if '-/-' not in s:
            return False
        prefix = s.split('-/-')[0]
        m = re.search(r'(\d)$', prefix)
        if not m:
            return False
        return m.group(1) in ['4', '5', '6']

    def top_n_per_group(agg_df, dg_filter, sort_col, ascending, top_n=TOP_N):
        """Lấy top N mã hàng mỗi nhóm theo tiêu chí sort_col"""
        subset = agg_df[agg_df['danh_gia'] == dg_filter].copy()
        subset = (subset
            .sort_values(['nhom', sort_col], ascending=[True, ascending])
            .groupby('nhom')
            .head(top_n)
            .reset_index(drop=True))
        return subset['ma_hang'].tolist()

    # ĐIỀU CHUYỂN: chỉ lấy model năm 2024-2026, ưu tiên tồn nhiều nhất trong nhóm
    dc_filter = agg[
        (agg['danh_gia'] == 'ĐIỀU CHUYỂN')
        &
        (agg['ma_hang'].apply(is_model_2024_2026))
    ].copy()

    mh_dc = (
        dc_filter
        .sort_values(['nhom','ton_3t'], ascending=[True,False])
        .groupby('nhom')
        .head(TOP_N)
        ['ma_hang']
        .tolist()
    )
    # NHẬP THÊM: ưu tiên ST3M cao + MOS thấp (bán nhanh, sắp hết)
    mh_nt = top_n_per_group(agg, 'BÁN CHẠY',    'st3m',   ascending=False)

    data_dc = build_data(mh_dc, is_dc=True)
    data_nt = build_data(mh_nt, is_dc=False)

    # Gắn thêm thứ tự trong nhóm và màu nhóm
    nhom_list = sorted(agg['nhom'].dropna().unique().tolist())
    nhom_color_map = {nhom: GROUP_COLORS[i % len(GROUP_COLORS)]
                      for i, nhom in enumerate(nhom_list)}

    def attach_group_info(df):
        """Thêm rank trong nhóm và màu nhóm"""
        if df is None or df.empty or 'nhom' not in df.columns:
            return pd.DataFrame(columns=list(df.columns) + ['rank_grp','_bg','_fg']) if df is not None else pd.DataFrame()

        df = df.copy()
        df['rank_grp'] = df.groupby('nhom').cumcount() + 1
        df['_bg'], df['_fg'] = zip(*df['nhom'].map(
            lambda n: nhom_color_map.get(n, ('F5F5F5','000000'))))
        return df

    data_dc = attach_group_info(data_dc)
    data_nt = attach_group_info(data_nt)

    # ─────────────────────────────────────────────────────────
    # BUILD EXCEL
    # ─────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    # ── Sheet 1: BAN_3_THANG_TONGHOP ─────────────────────────
    ws1 = wb.active; ws1.title = 'BAN_3_THANG_TONGHOP'
    title_row(ws1,1,1,7,'BÁN 3 THÁNG – TỔNG HỢP THEO MÃ HÀNG',C_BLUE)
    ws1.merge_cells('A2:G2')
    ws1.cell(2,1,f'Ngày phân tích: {TODAY}').font=F(size=9,color='555555')
    for i,(h,w) in enumerate(zip(
        ['Nhóm','Mã hàng','Tên hàng hóa','Bán 3T','Tồn 3T','ST3M','Đánh giá'],
        [15,22,46,10,10,10,14]),1):
        hdr(ws1,3,i,h,C_BLUE); cw(ws1,i,w)
    ws1.row_dimensions[3].height=22
    for ri,row in enumerate(agg.sort_values(['nhom','ban_3t'],ascending=[True,False]).itertuples(),1):
        er=ri+3; bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        dg=row.danh_gia; bg_dg,fg_dg=DANH_GIA_COLOR.get(dg,('FFFFFF','000000'))
        for ci,v in enumerate([row.nhom,row.ma_hang,row.ten_mau,row.ban_3t,row.ton_3t,
                                f'{row.st3m:.1%}',dg],1):
            if ci==7: dc(ws1,er,ci,v,bg=bg_dg,bold=True,color=fg_dg)
            else: dc(ws1,er,ci,v,bg=bg,al='center' if ci in(4,5,6) else 'left')
        ws1.row_dimensions[er].height=15
    ws1.freeze_panes='A4'
    ws1.auto_filter.ref=f'A3:G{len(agg)+3}'

    # ── Sheet 2: BAN_6_THANG_TONGHOP ─────────────────────────
    ws2=wb.create_sheet('BAN_6_THANG_TONGHOP')
    title_row(ws2,1,1,7,'BÁN 6 THÁNG – TỔNG HỢP THEO MÃ HÀNG',C_BLUE)
    ws2.merge_cells('A2:G2')
    ws2.cell(2,1,f'Ngày phân tích: {TODAY}').font=F(size=9,color='555555')
    for i,(h,w) in enumerate(zip(
        ['Nhóm','Mã hàng','Tên hàng hóa','Bán 6T','Tồn 6T','ST6M','Đánh giá'],
        [15,22,46,10,10,10,14]),1):
        hdr(ws2,3,i,h,C_BLUE); cw(ws2,i,w)
    ws2.row_dimensions[3].height=22
    for ri,row in enumerate(agg.sort_values(['nhom','ban_6t'],ascending=[True,False]).itertuples(),1):
        er=ri+3; bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        dg=row.danh_gia; bg_dg,fg_dg=DANH_GIA_COLOR.get(dg,('FFFFFF','000000'))
        for ci,v in enumerate([row.nhom,row.ma_hang,row.ten_mau,row.ban_6t,row.ton_6t,
                                f'{row.st6m:.1%}',dg],1):
            if ci==7: dc(ws2,er,ci,v,bg=bg_dg,bold=True,color=fg_dg)
            else: dc(ws2,er,ci,v,bg=bg,al='center' if ci in(4,5,6) else 'left')
        ws2.row_dimensions[er].height=15
    ws2.freeze_panes='A4'
    ws2.auto_filter.ref=f'A3:G{len(agg)+3}'

    # ── Sheet 3: DANH_GIA_TIEU_THU ───────────────────────────
    ws3=wb.create_sheet('DANH_GIA_TIEU_THU')
    title_row(ws3,1,1,14,'ĐÁNH GIÁ TIÊU THỤ – ST + MOS + TREND',C_BLUE)
    ws3.merge_cells('A2:N2')
    ws3.cell(2,1,
        f'ST=Bán/(Bán+Tồn) | MOS=Tồn÷(BánTB/tháng) | Trend=ST3M-ST6M | Ngày: {TODAY}'
    ).font=F(size=9,color='555555')
    HDG=['Nhóm','Mã hàng','Tên hàng hóa','Bán 3T','Bán 6T','Tồn',
         'ST3M','ST6M','MOS','VQ 3T','Trend','Đánh giá','Cảnh báo','Gợi ý']
    WDG=[15,22,44,9,9,9,8,8,8,9,9,14,35,25]
    for i,(h,w) in enumerate(zip(HDG,WDG),1):
        hdr(ws3,3,i,h,C_BLUE); cw(ws3,i,w)
    ws3.row_dimensions[3].height=25
    for ri,row in enumerate(agg.sort_values(['nhom','danh_gia','st3m'],
                                             ascending=[True,True,False]).itertuples(),1):
        er=ri+3; dg=row.danh_gia
        bg_dg,fg_dg=DANH_GIA_COLOR.get(dg,('FFFFFF','000000'))
        bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        trend_s=f'+{row.trend:.1%}' if row.trend>=0 else f'{row.trend:.1%}'
        mos_s = f'{row.mos3m:.1f}' if row.mos3m < 999 else '–'
        vals=[row.nhom,row.ma_hang,row.ten_mau,row.ban_3t,row.ban_6t,row.ton_3t,
              f'{row.st3m:.1%}',f'{row.st6m:.1%}',mos_s,f'{row.vq3m:.2f}',trend_s,
              dg,row.canh_bao,row.goi_y]
        for ci,v in enumerate(vals,1):
            if ci==12: dc(ws3,er,ci,v,bg=bg_dg,bold=True,color=fg_dg)
            elif ci==11:
                cbg='E2EFDA' if row.trend>=0.15 else('FFD9D9' if row.trend<=-0.15 else bg)
                dc(ws3,er,ci,v,bg=cbg)
            else: dc(ws3,er,ci,v,bg=bg,al='center' if ci in(4,5,6,7,8,9,10) else 'left')
        ws3.row_dimensions[er].height=15
    ws3.freeze_panes='A4'
    ws3.auto_filter.ref=f'A3:N{len(agg)+3}'

    # ── Sheet 4: SIZE_PRIORITY ────────────────────────────────
    ws4=wb.create_sheet('SIZE_PRIORITY')
    title_row(ws4,1,1,8,'PHÂN TÍCH SIZE – CHỈ SỐ ƯU TIÊN (gộp tất cả màu cùng size)',C_BLUE)
    ws4.merge_cells('A2:H2')
    ws4.cell(2,1,'Size Index = Tổng bán 6T (tất cả màu) / Max bán 6T các size trong mã hàng'
             ).font=F(size=9,color='555555')
    for i,(h,w) in enumerate(zip(
        ['Nhóm','Mã hàng','Size\n(digit)','Size\n(số)','Tổng bán\n6T','Tồn\ncác màu','Size Index','Xếp loại'],
        [15,22,8,8,12,11,11,14]),1):
        hdr(ws4,3,i,h,C_BLUE); cw(ws4,i,w)
    ws4.row_dimensions[3].height=30
    XLC={'Size chạy':('C6EFCE','006100'),'Size TB':('DDEBF7','1F497D'),
         'Size chậm':('FCE4D6','833C00'),'Không bán':('F2F2F2','888888')}
    for ri,row in enumerate(size_df.sort_values(['nhom','ma_hang','size_digit']).itertuples(),1):
        er=ri+3; bg='F5F5F5' if ri%2==0 else 'FFFFFF'
        xl=row.xep_loai; bg_xl,fg_xl=XLC.get(xl,('FFFFFF','000000'))
        sz_num=SIZE_MAP.get(row.size_digit,'?')
        for ci,v in enumerate([row.nhom,row.ma_hang,row.size_digit,sz_num,
                                row.ban_6t,row.ton,f'{row.size_index:.1%}',xl],1):
            if ci==8: dc(ws4,er,ci,v,bg=bg_xl,bold=True,color=fg_xl)
            else: dc(ws4,er,ci,v,bg=bg,al='center' if ci in(3,4,5,6,7) else 'left')
        ws4.row_dimensions[er].height=15
    ws4.freeze_panes='A4'
    ws4.auto_filter.ref=f'A3:H{len(size_df)+3}'

    # ── Hàm build phiếu (Sheet 5 & 6) ────────────────────────
    TON_COL = {sd: 12+i for i,sd in enumerate(SIZE_DIGITS)}   # col 13-22
    ACT_COL = {sd: 22+i for i,sd in enumerate(SIZE_DIGITS)}   # V-AE
    TONG_COL = 32
    DATA_ROW = 6

    def build_phieu(wb, sheet_name, title_text, bg_title, df_data, is_dc):
        ws = wb.create_sheet(sheet_name)
        n_cols = TONG_COL
        title_row(ws,1,1,n_cols,title_text,bg_title,rh=32)
        ws.merge_cells(f'A2:{get_column_letter(n_cols)}2')
        ws.cell(2,1,
            f'Ngày lập: {TODAY} | Top {TOP_N} mã hàng mỗi nhóm | '
            + ('Ưu tiên: Tồn nhiều nhất' if is_dc else 'Ưu tiên: ST3M cao, MOS thấp')
        ).font=F(size=9,color='555555')

        act_label = 'Đề xuất điều chuyển theo size' if is_dc else 'Đề xuất nhập thêm theo size'
        tong_label = 'Tổng\nđiều chuyển' if is_dc else 'Tổng\nnhập thêm'

        # Header nhóm cột (row 3)
        groups = [
            ('STT',          1,  1,  C_BLUE),
            ('Nhóm / Mã hàng',2, 3,  bg_title),
            ('N-X 3 tháng',  4,  7,  '2E4057'),
            ('N-X 6 tháng',  8,  11, '2F5496'),
            ('Tồn kho theo size',12,21,'34495E'),
            (act_label,     22, 31, '1A5276'),
            (tong_label,    32, 32, '7D6608'),
        ]
        for txt,c1,c2,bg in groups:
            ws.merge_cells(start_row=3,start_column=c1,end_row=3,end_column=c2)
            cell=ws.cell(3,c1,txt)
            cell.font=Font(bold=True,size=9,color='FFFFFF',name='Arial')
            cell.fill=FILL(bg); cell.alignment=AL(wrap=True); cell.border=BORDER

        # Header chi tiết (row 4)
        h4 = ['STT','Nhóm','Mã hàng',
              'Tồn đầu','Nhập','Xuất','Tồn cuối',   # 3T
              'Tồn đầu','Nhập','Xuất','Tồn cuối',   # 6T
              ] + [str(SIZE_MAP[s]) for s in SIZE_DIGITS] \
                + [str(SIZE_MAP[s]) for s in SIZE_DIGITS] \
                + ['Tổng']
        widths = [5,15,20]+[9]*8+[7]*10+[7]*10+[10]
        hdr_bgs = (
            [C_BLUE,bg_title,bg_title]+
            ['2E4057']*4+['2F5496']*4+
            ['34495E']*10+['1A5276']*10+['7D6608']
        )
        for ci,(h,w,bg) in enumerate(zip(h4,widths,hdr_bgs),1):
            hdr(ws,4,ci,h,bg); cw(ws,ci,w)
        ws.row_dimensions[3].height=22; ws.row_dimensions[4].height=24

        if df_data.empty:
            ws.cell(DATA_ROW,1,'Không có dữ liệu').font=F(size=10,color='888888')
            return ws

        key = 'dc' if is_dc else 'bs'
        prev_nhom = None

        for ri,row in enumerate(df_data.itertuples(),0):
            er = DATA_ROW + ri
            nhom_now = row.nhom
            is_first_in_group = (nhom_now != prev_nhom)
            thick_top = is_first_in_group and ri > 0   # đường kẻ đậm giữa các nhóm

            # Màu nền nhóm
            bg_grp, fg_grp = nhom_color_map.get(nhom_now, ('F5F5F5','000000'))
            # Xen kẽ sáng/tối trong nhóm
            bg_row = bg_grp if row.rank_grp % 2 == 1 else _lighten(bg_grp)

            def wr(col, val, bold=False, color='000000', al='center', bg_ov=None):
                dc(ws, er, col, val,
                   bg=bg_ov if bg_ov else bg_row,
                   bold=bold, color=color, al=al,
                   thick_top=thick_top)

            # Nhóm header khi bắt đầu nhóm mới – merge ô nhóm
            if is_first_in_group:
                # Tô header nhóm đậm hơn
                nhom_cell = ws.cell(er, 2, nhom_now)
                nhom_cell.font = Font(bold=True, size=9, color=fg_grp, name='Arial')
                nhom_cell.fill = FILL(bg_grp)
                nhom_cell.alignment = AL(h='left')
                nhom_cell.border = BORDER_THICK_TOP if thick_top else BORDER
            else:
                wr(2, '', al='left')

            wr(1,  row.rank_grp)
            wr(3,  row.ma_hang, al='left')
            wr(4,  row.ton_dau_3t)
            wr(5,  row.nhap_3t)
            wr(6,  row.xuat_3t)
            wr(7,  row.ton_cuoi_3t, bold=True)
            wr(8,  row.ton_dau_6t)
            wr(9,  row.nhap_6t)
            wr(10, row.xuat_6t)
            wr(11, row.ton_cuoi_6t, bold=True)

            # Tồn theo size (col 12-21)
            for sd,col in TON_COL.items():
                val = getattr(row, f'ton_{sd}', 0)
                cbg = 'FFD9D9' if val > 0 else bg_row
                wr(col, val, bold=(val>0), bg_ov=cbg)

            # Đề xuất theo size (col 22-31 → ACT_COL maps to 23+i)
            for sd,col in ACT_COL.items():
                val = getattr(row, f'{key}_{sd}', 0)
                cbg = 'E2EFDA' if val > 0 else bg_row
                clr = C_GREEN if val > 0 else '000000'
                wr(col, val, bold=(val>0), color=clr, bg_ov=cbg)

            tong = getattr(row, f'{key}_tong', 0)
            wr(TONG_COL, tong,
               bold=True, color=(C_RED if tong>0 else '000000'),
               bg_ov='FFF2CC')

            ws.row_dimensions[er].height = 16
            prev_nhom = nhom_now

        ws.freeze_panes = 'A5'
        ws.auto_filter.ref = f'A4:{get_column_letter(TONG_COL)}4'
        return ws

    build_phieu(wb,'DE_XUAT_DIEU_CHUYEN',
                f'📋 PHIẾU ĐỀ XUẤT ĐIỀU CHUYỂN – Top {TOP_N} mã mỗi nhóm (tồn nhiều, không bán 3T)',
                C_RED, data_dc, is_dc=True)
    build_phieu(wb,'DE_XUAT_NHAP_THEM',
                f'📋 PHIẾU ĐỀ XUẤT NHẬP THÊM – Top {TOP_N} mã mỗi nhóm (bán chạy nhất)',
                C_GREEN, data_nt, is_dc=False)

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf, agg

def _lighten(hex6):
    """Trả về màu nhạt hơn (~15%) để xen kẽ dòng trong nhóm"""
    r=int(hex6[0:2],16); g=int(hex6[2:4],16); b=int(hex6[4:6],16)
    r=min(255,r+30); g=min(255,g+30); b=min(255,b+30)
    return f'{r:02X}{g:02X}{b:02X}'

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
    with st.columns([1,3])[0]:
        run_btn = st.button("🚀 Chạy phân tích", use_container_width=True, type="primary")

    if run_btn:
        with st.spinner("Đang phân tích dữ liệu..."):
            try:
                file_bytes = BytesIO(uploaded.read())
                result_buf, agg = run_analysis(file_bytes)

                count = agg['danh_gia'].value_counts().to_dict()
                st.markdown("---")
                st.subheader("📊 Kết quả phân tích")
                m1,m2,m3,m4,m5 = st.columns(5)
                m1.metric("🔥 Bán chạy",    count.get('BÁN CHẠY',0))
                m2.metric("✅ Bán khá",     count.get('BÁN KHÁ',0))
                m3.metric("🔶 Bán vừa",     count.get('BÁN VỪA',0))
                m4.metric("🟠 Bán chậm",    count.get('BÁN CHẬM',0))
                m5.metric("🔴 Điều chuyển", count.get('ĐIỀU CHUYỂN',0))

                st.markdown("#### 🔴 Top điều chuyển theo nhóm")
                dc_prev = (agg[agg['danh_gia']=='ĐIỀU CHUYỂN']
                    .sort_values(['nhom','ton_3t'],ascending=[True,False])
                    .groupby('nhom').head(5)
                    [['nhom','ma_hang','ten_mau','ban_3t','ton_3t','st6m']]
                    .rename(columns={'nhom':'Nhóm','ma_hang':'Mã hàng','ten_mau':'Tên',
                                     'ban_3t':'Bán 3T','ton_3t':'Tồn','st6m':'ST6M'}))
                dc_prev['ST6M'] = dc_prev['ST6M'].map('{:.1%}'.format)
                st.dataframe(dc_prev, use_container_width=True, hide_index=True)

                st.markdown("#### 🔥 Top bán chạy theo nhóm")
                nt_prev = (agg[agg['danh_gia']=='BÁN CHẠY']
                    .sort_values(['nhom','st3m'],ascending=[True,False])
                    .groupby('nhom').head(5)
                    [['nhom','ma_hang','ten_mau','ban_3t','ton_3t','st3m','mos3m']]
                    .rename(columns={'nhom':'Nhóm','ma_hang':'Mã hàng','ten_mau':'Tên',
                                     'ban_3t':'Bán 3T','ton_3t':'Tồn','st3m':'ST3M','mos3m':'MOS'}))
                nt_prev['ST3M'] = nt_prev['ST3M'].map('{:.1%}'.format)
                nt_prev['MOS']  = nt_prev['MOS'].apply(lambda x: f'{x:.1f}' if x<999 else '–')
                st.dataframe(nt_prev, use_container_width=True, hide_index=True)

                st.markdown("---")
                fname = f"Phan_Tich_Ton_Kho_{date.today().strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    label="⬇️ Tải xuống file kết quả (.xlsx)",
                    data=result_buf, file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary"
                )
                st.caption("File gồm 6 sheet: BAN_3_THANG_TONGHOP | BAN_6_THANG_TONGHOP | DANH_GIA_TIEU_THU | SIZE_PRIORITY | DE_XUAT_DIEU_CHUYEN | DE_XUAT_NHAP_THEM")

            except Exception as e:
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)
else:
    st.info("👆 Vui lòng upload file Excel để bắt đầu phân tích.")
    with st.expander("ℹ️ Hướng dẫn"):
        st.markdown(f"""
**Yêu cầu:** File `.xlsx` có 2 sheet `6_thang` và `3_thang`, dữ liệu từ dòng 10.

**6 sheet output:**
| Sheet | Nội dung |
|---|---|
| BAN_3_THANG_TONGHOP | Tổng hợp 3T theo mã hàng |
| BAN_6_THANG_TONGHOP | Tổng hợp 6T theo mã hàng |
| DANH_GIA_TIEU_THU | ST + MOS + Trend + Đánh giá |
| SIZE_PRIORITY | Size Index (gộp tất cả màu) |
| DE_XUAT_DIEU_CHUYEN | Top {TOP_N} mỗi nhóm – tồn không bán |
| DE_XUAT_NHAP_THEM | Top {TOP_N} mỗi nhóm – bán chạy nhất |

**Tiêu chí:**
- `ST = Bán/(Bán+Tồn)` — `MOS = Tồn÷(BánTB/tháng)`
- **BÁN CHẠY**: ST3M≥60% và MOS<1.5 tháng
- **ĐIỀU CHUYỂN**: Còn tồn, không bán 3T, ST6M<30%
        """)
