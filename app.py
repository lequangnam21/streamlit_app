import streamlit as st
import pandas as pd
from io import BytesIO

# =====================
# CONFIG
# =====================

st.set_page_config(
    page_title="PHÂN TÍCH ĐIỀU CHUYỂN HÀNG",
    layout="wide"
)

st.title("PHÂN TÍCH ĐIỀU CHUYỂN HÀNG")

# =====================
# UPLOAD
# =====================

file_ban=st.file_uploader(
    "UPLOAD FILE BÁN HÀNG",
    type=["xlsx"]
)

file_ton=st.file_uploader(
    "UPLOAD FILE TỒN KHO",
    type=["xlsx"]
)

# =====================
# CLEAN
# =====================

def clean_columns(df):

    df.columns=(

        df.columns
        .astype(str)
        .str.strip()
        .str.replace(" ","_")

    )

    return df

# =====================
# RUN
# =====================

if file_ban and file_ton:

    try:

        # ---------------------
        # READ FILE
        # ---------------------

        sheet_1=pd.read_excel(
            file_ban,
            sheet_name="1_thang"
        )

        sheet_3=pd.read_excel(
            file_ban,
            sheet_name="3_thang"
        )

        sheet_6=pd.read_excel(
            file_ban,
            sheet_name="6_thang"
        )

        ton=pd.read_excel(file_ton)

        sheet_1=clean_columns(sheet_1)
        sheet_3=clean_columns(sheet_3)
        sheet_6=clean_columns(sheet_6)
        ton=clean_columns(ton)

        # ---------------------
        # CỬA HÀNG
        # ---------------------

        cua_hang=(

            sheet_1["Ten_CH"]
            .iloc[0]

        )

        st.info(
            f"Cửa hàng đang phân tích: {cua_hang}"
        )

        # ---------------------
        # BÁN 1 THÁNG
        # ---------------------

        ban_1=(

            sheet_1
            .groupby(
                ["Ma_hang","Ten_hang"],
                as_index=False
            )["SL_ban"]
            .sum()
            .rename(
                columns={
                    "SL_ban":"Ban_1"
                }
            )

        )

        # ---------------------
        # BÁN 3 THÁNG
        # ---------------------

        ban_3=(

            sheet_3
            .groupby(
                "Ma_hang",
                as_index=False
            )["SL_ban"]
            .sum()
            .rename(
                columns={
                    "SL_ban":"Ban_3"
                }
            )

        )

        # ---------------------
        # BÁN 6 THÁNG
        # ---------------------

        ban_6=(

            sheet_6
            .groupby(
                "Ma_hang",
                as_index=False
            )["SL_ban"]
            .sum()
            .rename(
                columns={
                    "SL_ban":"Ban_6"
                }
            )

        )

        # ---------------------
        # MERGE
        # ---------------------

        df=ban_1.merge(

            ban_3,

            on="Ma_hang",

            how="left"

        )

        df=df.merge(

            ban_6,

            on="Ma_hang",

            how="left"

        )

        df=df.fillna(0)

        # ---------------------
        # ĐIỀU KIỆN
        # ---------------------

        dieu_kien=(

            (df["Ban_1"]>=2)

            &

            (

                (df["Ban_3"]>1)

                |

                (df["Ban_6"]>1)

            )

        )

        df=df[dieu_kien].copy()

        # ---------------------
        # TỒN CHÍNH CH
        # ---------------------

        ton_ch=ton[
            ton["Ten_CH"]==cua_hang
        ].copy()

        ton_ngoai=ton[
            ton["Ten_CH"]!=cua_hang
        ].copy()

        # ---------------------
        # SIZE CHI TIẾT
        # ---------------------

        def lay_size(ma_hang):

            tmp=ton_ch[

                ton_ch["Ma_hang"]
                .astype(str)
                .str.startswith(
                    str(ma_hang)
                )

            ][[
                "Ma_hang",
                "Ton"
            ]]

            tmp=tmp.sort_values(
                "Ma_hang"
            )

            if len(tmp)==0:

                return ""

            return ", ".join(

                f'{r["Ma_hang"]}({int(r["Ton"])})'

                for _,r in tmp.iterrows()

            )

        df["Size_chi_tiet"]=(
            df["Ma_hang"]
            .apply(
                lay_size
            )
        )

        # ---------------------
        # TỔNG TỒN
        # ---------------------

        def tong_ton(ma_hang):

            tmp=ton_ch[

                ton_ch["Ma_hang"]
                .astype(str)
                .str.startswith(
                    str(ma_hang)
                )

            ]

            return tmp["Ton"].sum()

        df["Ton_tai_cua_hang"]=(

            df["Ma_hang"]
            .apply(
                tong_ton
            )

        )

        # ---------------------
        # LOOKUP ĐIỀU CHUYỂN
        # ---------------------

        def de_xuat(ma_hang):

            ds_size=(

                ton_ch[

                    ton_ch["Ma_hang"]
                    .astype(str)
                    .str.startswith(
                        str(ma_hang)
                    )

                ]["Ma_hang"]

                .unique()

            )

            ketqua=[]

            for size in ds_size:

                tmp=ton_ngoai[

                    (

                        ton_ngoai["Ma_hang"]
                        ==size

                    )

                    &

                    (

                        ton_ngoai["Ton"]>0

                    )

                ]

                if len(tmp)==0:

                    continue

                tmp=tmp.sort_values(

                    by="Ton",

                    ascending=False

                )

                noi=", ".join(

                    f'{r["Ten_CH"]}({int(r["Ton"])})'

                    for _,r in tmp.iterrows()

                )

                ketqua.append(

                    f"{size} → {noi}"

                )

            return " ; ".join(ketqua)

        df["Đề_xuất_nơi_lấy"]=(

            df["Ma_hang"]
            .apply(
                de_xuat
            )

        )

        # ---------------------
        # OUTPUT
        # ---------------------

        hien_thi=df[[

            "Ma_hang",
            "Ten_hang",
            "Ban_1",
            "Ban_3",
            "Ban_6",
            "Ton_tai_cua_hang",
            "Size_chi_tiet",
            "Đề_xuất_nơi_lấy"

        ]]

        hien_thi.columns=[

            "Mã hàng",
            "Tên hàng",
            "Bán 1 tháng",
            "Bán 3 tháng",
            "Bán 6 tháng",
            "Tồn tại cửa hàng",
            "Size chi tiết",
            "Đề xuất nơi lấy"

        ]

        st.success(
            "PHÂN TÍCH THÀNH CÔNG"
        )

        st.write(
            f"Số dòng đề xuất: {len(hien_thi)}"
        )

        st.dataframe(
            hien_thi,
            use_container_width=True,
            height=750
        )

        # ---------------------
        # EXPORT
        # ---------------------

        output=BytesIO()

        with pd.ExcelWriter(

            output,

            engine="openpyxl"

        ) as writer:

            hien_thi.to_excel(

                writer,

                index=False,

                sheet_name="De_xuat"

            )

        st.download_button(

            label="TẢI FILE ĐỀ XUẤT",

            data=output.getvalue(),

            file_name="de_xuat_dieu_chuyen.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

    except Exception as e:

        st.error("LỖI HỆ THỐNG")

        st.exception(e)