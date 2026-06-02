import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from copy import copy
from collections import defaultdict
from io import BytesIO

# ==========================
# WEB UI
# ==========================
st.set_page_config(

    page_title=
    "Đề xuất bổ sung tồn kho",

    layout=
    "wide"
)

st.title(

    "ĐỀ XUẤT BỔ SUNG TỒN KHO - GIẢM TỒN"
)

col1,col2,col3 = st.columns(3)

with col1:

    DOANH_THU_FILE = st.file_uploader(

        "Upload Doanh thu",

        type=[
            "xlsx",
            "xls"
        ]
    )

with col2:

    TON_KHO_FILE = st.file_uploader(

        "Upload nhập tồn hệ thống",

        type=[
            "xlsx",
            "xls"
        ]
    )

with col3:

    MAU_FILE = st.file_uploader(

        "Upload Mẫu file kết quả",

        type=[
            "xlsx",
            "xls"
        ]
    )

RUN = st.button(

    "RUN PHÂN TÍCH",

    use_container_width=True
)

# ==========================
# CONFIG
# ==========================
KHO_TONG = (

    "CÔNG TY TNHH SX-TM-DV ĐĂNG TUẤN"
)

SIZE_MAP = {

    "5":35,
    "6":36,
    "7":37,
    "8":38,
    "9":39,
    "0":40,
    "1":41,
    "2":42,
    "3":43,
    "4":44
}

SIZE_ORDER = [

    35,36,37,38,39,

    40,41,42,43,44
]

# ==========================
# HELPER
# ==========================
def get_model(ma):

    ma = str(ma)

    if "-/-" in ma:

        return ma.split(
            "-/-"
        )[0]

    if "-" in ma:

        return ma.split(
            "-"
        )[0]

    return ma


def is_hm(ma):

    return (

        "HM"

        in

        str(ma).upper()
    )


def get_base_code(ma):

    ma = str(ma)

    if len(ma)==0:

        return ma

    last_char = ma[-1]

    if last_char in SIZE_MAP:

        return ma[:-1]

    return ma


def get_size_from_code(ma):

    ma = str(ma)

    if len(ma)==0:

        return None

    last_char = ma[-1]

    return SIZE_MAP.get(
        last_char
    )


def danh_gia_st(st):

    if st < 0.25:

        return (
            "Mẫu yếu - Không đề xuất"
        )

    elif st < 0.5:

        return (
            "Mẫu ổn"
        )

    elif st < 1:

        return (
            "Mẫu khoẻ"
        )

    return (
        "Mẫu cực khoẻ"
    )


def copy_row_style(

    ws,
    source_row,
    target_row,
    max_col=34
):

    for col in range(

        1,
        max_col+1
    ):

        s = ws.cell(
            source_row,
            col
        )

        t = ws.cell(
            target_row,
            col
        )

        if s.has_style:

            t._style = copy(
                s._style
            )

        t.number_format = (
            s.number_format
        )

        t.font = copy(
            s.font
        )

        t.fill = copy(
            s.fill
        )

        t.border = copy(
            s.border
        )

        t.alignment = copy(
            s.alignment
        )

        t.protection = copy(
            s.protection
        )

# ==========================
# MAIN PROCESS
# ==========================
if RUN:

    if (

        DOANH_THU_FILE is None

        or

        TON_KHO_FILE is None

        or

        MAU_FILE is None
    ):

        st.error(
            "Thiếu file upload."
        )

        st.stop()

    with st.spinner(

        "Đang xử lý..."
    ):

        # ==========================
        # LOAD DATA
        # ==========================
        df_1 = pd.read_excel(
            DOANH_THU_FILE,
            sheet_name="1_thang"
        )

        df_3 = pd.read_excel(
            DOANH_THU_FILE,
            sheet_name="3_thang"
        )

        df_6 = pd.read_excel(
            DOANH_THU_FILE,
            sheet_name="6_thang"
        )

        df_kho = pd.read_excel(
            TON_KHO_FILE
        )

        # ==========================
        # FILTER HM
        # ==========================
        df_1 = df_1[
            ~df_1["Ma_hang"]
            .astype(str)
            .str.contains(
                "HM",
                case=False
            )
        ]

        df_3 = df_3[
            ~df_3["Ma_hang"]
            .astype(str)
            .str.contains(
                "HM",
                case=False
            )
        ]

        df_6 = df_6[
            ~df_6["Ma_hang"]
            .astype(str)
            .str.contains(
                "HM",
                case=False
            )
        ]

        df_kho = df_kho[
            ~df_kho["Ma_hang"]
            .astype(str)
            .str.contains(
                "HM",
                case=False
            )
        ]

        # ==========================
        # DETECT STORE
        # từ file doanh thu
        # ==========================
        CURRENT_STORE = None

        for df_dt in [

            df_1,
            df_3,
            df_6
        ]:

            if "Ten_CH" in df_dt.columns:

                stores = (

                    df_dt["Ten_CH"]

                    .dropna()

                    .astype(str)

                    .unique()

                    .tolist()
                )

                if len(stores)>0:

                    CURRENT_STORE = (
                        stores[0]
                    )

                    break

        if CURRENT_STORE is None:

            st.error(

                "Không detect được Ten_CH từ file doanh thu."
            )

            st.stop()

        st.info(

            f"Cửa hàng đang phân tích: "

            f"{CURRENT_STORE}"
        )

        # ==========================
        # CHỈ LẤY TỒN KHO
        # CỬA HÀNG ĐANG PHÂN TÍCH
        # ==========================
        df_kho = df_kho[

            df_kho["Ten_CH"]

            .astype(str)

            ==

            str(CURRENT_STORE)
        ].copy()
                # ==========================
        # SALES MAP
        # MODEL LEVEL
        # dùng để check mã giày
        # ==========================
        sale_1 = (

            df_1

            .groupby(
                "Ma_hang"
            )

            ["SL_ban"]

            .sum()

            .to_dict()
        )

        sale_3 = (

            df_3

            .groupby(
                "Ma_hang"
            )

            ["SL_ban"]

            .sum()

            .to_dict()
        )

        sale_6 = (

            df_6

            .groupby(
                "Ma_hang"
            )

            ["SL_ban"]

            .sum()

            .to_dict()
        )

        all_sales_codes = (

            set(
                sale_1.keys()
            )

            |

            set(
                sale_3.keys()
            )

            |

            set(
                sale_6.keys()
            )
        )

        # ==========================
        # INVENTORY MAP
        # DE_XUAT_BO_SUNG
        # giữ nguyên
        # ==========================
        ton_by_base = defaultdict(

            lambda:
            defaultdict(int)
        )

        detail_by_store = defaultdict(

            lambda:
            defaultdict(list)
        )

        available_sizes = defaultdict(
            set
        )

        nhap3_map = defaultdict(
            int
        )

        nhap6_map = defaultdict(
            int
        )

        # ==========================
        # HANG_CAN_GIAM_TON_KHO MAP
        # PATCH 2.6.2
        # MODEL = mã giày
        # SIZE = chi tiết tồn
        # ==========================
        giam_ton_store_model = defaultdict(list)

        for _, r in df_kho.iterrows():

            ma = str(
                r["Ma_hang"]
            )

            ton = (

                r["Ton"]

                if pd.notna(
                    r["Ton"]
                )

                else 0
            )

            nhap3 = (

                r["Nhap_3_thang"]

                if pd.notna(
                    r["Nhap_3_thang"]
                )

                else 0
            )

            nhap6 = (

                r["Nhap_6_thang"]

                if pd.notna(
                    r["Nhap_6_thang"]
                )

                else 0
            )

            base = get_base_code(
                ma
            )

            size = get_size_from_code(
                ma
            )

            # ==========================
            # DE_XUAT_BO_SUNG
            # giữ nguyên
            # ==========================
            if size is not None:

                ton_by_base[
                    base
                ][
                    size
                ] += ton

                available_sizes[
                    base
                ].add(
                    size
                )

                if ton > 0:

                    detail_by_store[
                        base
                    ][
                        CURRENT_STORE
                    ].append(

                        f"{ma}({int(ton)})"
                    )

                nhap3_map[
                    base
                ] += nhap3

                nhap6_map[
                    base
                ] += nhap6

            # ==========================
            # HANG_CAN_GIAM_TON_KHO
            # CHỈ XÉT CÓ TỒN
            # ==========================
            if ton <= 0:

                continue

            # ==========================
            # MODEL PHẢI LÀ MÃ GIÀY
            #
            # ví dụ:
            #
            # AA20C2-/-BG74
            # AA20C2-/-BG745
            # AA20C2-/-BG746
            #
            # model = AA20C2-/-BG74
            # ==========================
            model = base

            # ==========================
            # PHẢI CÓ TRONG DOANH THU
            # mới xét giảm tồn
            # ==========================
            if model not in all_sales_codes:

                continue

            # ==========================
            # BUILD MAP
            # ==========================
            giam_ton_store_model[
                model
            ].append({

                "ma_size":
                ma,

                "ton":
                ton,

                "nhap3":
                nhap3,

                "nhap6":
                nhap6
            })
                    # ==========================
        # BUILD RESULT
        # DE_XUAT_BO_SUNG
        # giữ nguyên logic 2.6
        # ==========================
        all_codes = sorted(

            set(
                sale_1.keys()
            )

            |

            set(
                sale_3.keys()
            )

            |

            set(
                sale_6.keys()
            ),

            key=lambda x:(

                get_model(x),

                x
            )
        )

        results = []

        for ma in all_codes:

            if is_hm(ma):

                continue

            sl1 = sale_1.get(
                ma,
                0
            )

            sl3 = sale_3.get(
                ma,
                0
            )

            sl6 = sale_6.get(
                ma,
                0
            )

            nhap3 = nhap3_map.get(
                ma,
                0
            )

            nhap6 = nhap6_map.get(
                ma,
                0
            )

            st3 = (

                sl3 / nhap3

                if nhap3 > 0

                else 0
            )

            st6 = (

                sl6 / nhap6

                if nhap6 > 0

                else 0
            )

            danh_gia = danh_gia_st(
                st6
            )

            # ==========================
            # PASS BỔ SUNG
            # giữ nguyên
            # ==========================
            is_new_product = (

                nhap6 > 0

                and

                (
                    nhap3 / nhap6
                ) >= 0.7
            )

            if is_new_product:

                pass_bo_sung = (
                    st3 >= 0.25
                )

            else:

                pass_bo_sung = (

                    st6 >= 0.25

                    or

                    (
                        st6 < 0.25
                        and
                        st3 >= 0.25
                    )
                )

            if not pass_bo_sung:

                continue

            ton_sizes = ton_by_base.get(
                ma,
                {}
            )

            model_sizes = sorted(

                available_sizes.get(
                    ma,
                    set()
                )
            )

            de_xuat_size = {}

            tong_dx = 0

            for s in model_sizes:

                ton_hien_tai = (

                    ton_sizes.get(
                        s,
                        0
                    )
                )

                need = max(

                    2 - ton_hien_tai,

                    0
                )

                de_xuat_size[
                    s
                ] = need

                tong_dx += need

            # ==========================
            # TRẠNG THÁI SIZE
            # ==========================
            size_count = sum(

                1

                for x in ton_sizes.values()

                if x > 0
            )

            trang_thai_size = (

                "Đủ size"

                if size_count >= 3

                else "Hụt size"
            )

            # ==========================
            # CHI TIẾT TỒN
            # chỉ cửa hàng đang phân tích
            # ==========================
            chi_tiet = []

            for ch, arr in detail_by_store.get(

                ma,

                {}

            ).items():

                if arr:

                    chi_tiet.append(

                        f"{ch}: "

                        +

                        ", ".join(arr)
                    )

            chi_tiet_text = "\n".join(
                chi_tiet
            )

            # ==========================
            # ĐIỀU CHUYỂN
            # chỉ CURRENT_STORE
            # ==========================
            dieu_chuyen_lines = []

            for size in model_sizes:

                need = de_xuat_size.get(
                    size,
                    0
                )

                if need <= 0:

                    continue

                size_code = list(
                    SIZE_MAP.keys()
                )[

                    list(
                        SIZE_MAP.values()
                    ).index(
                        size
                    )
                ]

                ma_size = (
                    f"{ma}{size_code}"
                )

                kho_rows = df_kho[

                    df_kho[
                        "Ma_hang"
                    ].astype(str)

                    ==

                    ma_size
                ]

                remain = need

                # ==========================
                # KHO TỔNG
                # ==========================
                kho_tong = kho_rows[

                    kho_rows[
                        "Ten_CH"
                    ]

                    ==

                    KHO_TONG
                ]

                if len(
                    kho_tong
                ) > 0:

                    ton_kho_tong = int(

                        kho_tong[
                            "Ton"
                        ].sum()
                    )

                    if ton_kho_tong > 0:

                        lay = min(

                            remain,

                            ton_kho_tong
                        )

                        dieu_chuyen_lines.append(

                            f"{KHO_TONG}: "

                            f"{ma_size} "

                            f"({lay})"
                        )

                        remain -= lay

            results.append({

                "ma":
                ma,

                "sl1":
                sl1,

                "sl3":
                sl3,

                "sl6":
                sl6,

                "st3":
                st3,

                "st6":
                st6,

                "danh_gia":
                danh_gia,

                "ton_sizes":
                ton_sizes,

                "dx_sizes":
                de_xuat_size,

                "tong_dx":
                tong_dx,

                "trang_thai":
                trang_thai_size,

                "chi_tiet":
                chi_tiet_text,

                "dieu_chuyen":

                "\n".join(
                    dieu_chuyen_lines
                )
            })

        st.success(

            f"Phân tích xong: "

            f"{len(results)} mẫu."
        )
                # ==========================
        # BUILD
        # HANG_CAN_GIAM_TON_KHO
        # 2.6.2 FINAL
        # ==========================
        giam_ton_results = []

        for model, rows in (

            giam_ton_store_model.items()
        ):

            # ==========================
            # MODEL LEVEL
            # AA20C2-/-BG74
            # ==========================
            sl1 = sale_1.get(
                model,
                0
            )

            sl3 = sale_3.get(
                model,
                0
            )

            sl6 = sale_6.get(
                model,
                0
            )

            # ==========================
            # CHỈ TÍNH TỒN
            # CỬA HÀNG ĐANG PHÂN TÍCH
            # ==========================
            tong_ton = sum(

                x["ton"]

                for x in rows
            )

            if tong_ton <= 0:

                continue

            # ==========================
            # NHẬP
            # STORE LEVEL
            # ==========================
            nhap3 = sum(

                x["nhap3"]

                for x in rows
            )

            nhap6 = sum(

                x["nhap6"]

                for x in rows
            )

            # ==========================
            # ST
            # MODEL LEVEL
            # ==========================
            st3 = (

                sl3 / nhap3

                if nhap3 > 0

                else 0
            )

            st6 = (

                sl6 / nhap6

                if nhap6 > 0

                else 0
            )

            # ==========================
            # TIÊU CHÍ GIẢM TỒN
            # OPTION A
            # ==========================
            loai_giam_ton = None

            model_not_sell = (

                sl3 == 0

                and

                sl6 == 0
            )

            if model_not_sell:

                loai_giam_ton = (

                    "Không bán 3T + 6T"
                )

            elif (

                st3 < 0.25

                and

                st6 < 0.25
            ):

                loai_giam_ton = (

                    "Bán chậm 3T + 6T"
                )

            elif (

                nhap6 > 0

                and

                (
                    nhap3 / nhap6
                ) >= 0.7
            ):

                loai_giam_ton = (

                    "Hàng mới - theo dõi thêm"
                )

            if loai_giam_ton is None:

                continue

            # ==========================
            # CHI TIẾT TỒN SIZE
            # chỉ size đang tồn
            # ==========================
            chi_tiet_ton_size = []

            for rr in rows:

                ton = rr["ton"]

                if ton <= 0:

                    continue

                chi_tiet_ton_size.append(

                    f'{rr["ma_size"]}'

                    f'({int(ton)})'
                )

            if len(
                chi_tiet_ton_size
            ) == 0:

                continue

            chi_tiet_ton_size_text = (

                ", ".join(
                    chi_tiet_ton_size
                )
            )

            # ==========================
            # OUTPUT
            # CURRENT STORE ONLY
            # ==========================
            giam_ton_results.append({

                "Ten_CH":
                CURRENT_STORE,

                "Ma_hang":
                model,

                "Tong_ton":
                tong_ton,

                "Ban_3":
                sl3,

                "ST_3M":
                st3,

                "Ban_6":
                sl6,

                "ST_6M":
                st6,

                "Loai_Giam_Ton":
                loai_giam_ton,

                "Chi_tiet_ton_size":
                chi_tiet_ton_size_text
            })
                    # ==========================
        # WRITE EXCEL
        # ==========================
        wb = load_workbook(
            MAU_FILE
        )

        # ==========================
        # SHEET DE_XUAT_BO_SUNG
        # ==========================
        ws = wb[
            "DE_XUAT_BO_SUNG"
        ]

        START_ROW = 9
        template_row = 9

        for r in range(

            START_ROW,

            ws.max_row + 1
        ):

            for c in range(
                1,
                35
            ):

                cell = ws.cell(
                    r,
                    c
                )

                if isinstance(
                    cell,
                    MergedCell
                ):

                    continue

                cell.value = None

        def safe_write(
            ws,
            r,
            c,
            value
        ):

            cell = ws.cell(
                r,
                c
            )

            if isinstance(
                cell,
                MergedCell
            ):

                return

            cell.value = value

        row = START_ROW

        for i,item in enumerate(

            results,

            start=1
        ):

            copy_row_style(
                ws,
                template_row,
                row
            )

            safe_write(
                ws,row,1,i
            )

            safe_write(
                ws,row,2,
                item["ma"]
            )

            safe_write(
                ws,row,3,
                item["sl1"]
            )

            safe_write(
                ws,row,4,
                item["sl3"]
            )

            safe_write(
                ws,row,5,
                item["sl6"]
            )

            for idx,size in enumerate(
                SIZE_ORDER
            ):

                safe_write(

                    ws,

                    row,

                    6+idx,

                    item[
                        "ton_sizes"
                    ].get(
                        size,
                        0
                    )
                )

            for idx,size in enumerate(
                SIZE_ORDER
            ):

                safe_write(

                    ws,

                    row,

                    17+idx,

                    item[
                        "dx_sizes"
                    ].get(
                        size,
                        0
                    )
                )

            safe_write(
                ws,
                row,
                27,
                item["dieu_chuyen"]
            )

            safe_write(
                ws,
                row,
                28,
                item["chi_tiet"]
            )

            safe_write(
                ws,
                row,
                29,
                item["tong_dx"]
            )

            safe_write(
                ws,
                row,
                30,
                item["st6"]
            )

            ws.cell(
                row,
                30
            ).number_format='0%'

            safe_write(
                ws,
                row,
                31,
                item["st3"]
            )

            ws.cell(
                row,
                31
            ).number_format='0%'

            safe_write(
                ws,
                row,
                32,
                item["danh_gia"]
            )

            safe_write(
                ws,
                row,
                33,
                item["trang_thai"]
            )

            row += 1

        # ==========================
        # SHEET
        # HANG_CAN_GIAM_TON_KHO
        # 2.6.2 FINAL
        # ==========================
        SHEET_NAME = (
            "hang_can_giam_ton_kho"
        )

        if SHEET_NAME in wb.sheetnames:

            ws2 = wb[
                SHEET_NAME
            ]

            ws2.delete_rows(
                1,
                ws2.max_row
            )

        else:

            ws2 = wb.create_sheet(
                SHEET_NAME
            )

        # ==========================
        # HEADER
        # ==========================
        headers = [

            "Ten_CH",

            "Ma_hang",

            "Tong_ton",

            "Ban_3",

            "ST_3M",

            "Ban_6",

            "ST_6M",

            "Loai_Giam_Ton",

            "Chi_tiet_ton_size"
        ]

        header_fill = PatternFill(

            start_color="D9EAD3",

            end_color="D9EAD3",

            fill_type="solid"
        )

        for col,header in enumerate(

            headers,

            start=1
        ):

            cell = ws2.cell(

                1,

                col,

                header
            )

            cell.fill = (
                header_fill
            )

        # ==========================
        # WRITE DATA
        # ==========================
        row2 = 2

        for item in giam_ton_results:

            ws2.cell(
                row2,
                1,
                item["Ten_CH"]
            )

            ws2.cell(
                row2,
                2,
                item["Ma_hang"]
            )

            ws2.cell(
                row2,
                3,
                item["Tong_ton"]
            )

            ws2.cell(
                row2,
                4,
                item["Ban_3"]
            )

            ws2.cell(
                row2,
                5,
                item["ST_3M"]
            )

            ws2.cell(
                row2,
                6,
                item["Ban_6"]
            )

            ws2.cell(
                row2,
                7,
                item["ST_6M"]
            )

            ws2.cell(
                row2,
                8,
                item[
                    "Loai_Giam_Ton"
                ]
            )

            ws2.cell(
                row2,
                9,
                item[
                    "Chi_tiet_ton_size"
                ]
            )

            ws2.cell(
                row2,
                5
            ).number_format='0%'

            ws2.cell(
                row2,
                7
            ).number_format='0%'

            row2 += 1

        # ==========================
        # SORT OUTPUT
        # ==========================
        ws2.auto_filter.ref = (

            f"A1:I{row2}"
        )

        # ==========================
        # AUTO WIDTH
        # ==========================
        for column_cells in (

            ws2.columns
        ):

            length = max(

                len(
                    str(cell.value)
                )

                if cell.value
                else 0

                for cell in column_cells
            )

            ws2.column_dimensions[

                column_cells[0]
                .column_letter

            ].width = min(
                length+5,
                60
            )

        # ==========================
        # EXPORT MEMORY
        # ==========================
        output = BytesIO()

        wb.save(
            output
        )

        output.seek(0)

        st.success(
            "Hoàn tất tạo file."
        )

        # ==========================
        # PREVIEW
        # ==========================
        preview_df = pd.DataFrame([

            {

                "Mã hàng":
                x["ma"],

                "SL1":
                x["sl1"],

                "SL3":
                x["sl3"],

                "SL6":
                x["sl6"],

                "ST6":
                round(
                    x["st6"],
                    2
                ),

                "Tổng DX":
                x["tong_dx"]
            }

            for x in results
        ])

        st.subheader(
            "Preview kết quả"
        )

        st.dataframe(

            preview_df,

            use_container_width=True
        )

        st.download_button(

            label=
            "DOWNLOAD ket_qua_de_xuat.xlsx",

            data=output,

            file_name=
            "ket_qua_de_xuat.xlsx",

            mime=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            use_container_width=True
        )
