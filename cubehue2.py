import win32com.client
import pandas as pd
import plotly.graph_objects as go
import os

# ==========================
# KẾT NỐI SSAS
# ==========================

server_name = r'MERINGUE\ADMIN'
database_name = 'MultidimensionalProject4'

conn_str = (
    f"Provider=MSOLAP;"
    f"Data Source={server_name};"
    f"Initial Catalog={database_name};"
)

mdx_query = """
WITH 
SET [ProductList] AS [DIM SAN PHAM].[Ten SP].[Ten SP].Members
SET [QuarterList] AS [DIM NGAY].[Quy].[Quy].Members
SET [BranchList] AS [DIM CHI NHANH].[Ten CN].[Ten CN].Members

MEMBER [Measures].[Doanh Thu] AS [Measures].[Tong Tien], FORMAT_STRING = '#,##0'

MEMBER [Measures].[Xep Hang Theo Quy] AS
    Rank(
        [DIM SAN PHAM].[Ten SP].CurrentMember,
        Order([ProductList], [Measures].[Tong Tien], BDESC)
    )

MEMBER [Measures].[So Quy Trong Top10] AS
    Count(
        Filter(
            [QuarterList],
            ([Measures].[Xep Hang Theo Quy], [DIM CHI NHANH].[Ten CN].CurrentMember) <= 10
        )
    )

MEMBER [Measures].[Chien Luoc] AS
    CASE
        WHEN [Measures].[So Quy Trong Top10] >= 3 THEN 'CHU LUC'
        WHEN [Measures].[So Quy Trong Top10] = 2 THEN 'TIEM NANG'
        WHEN [Measures].[So Quy Trong Top10] = 1 THEN 'MUA VU'
        ELSE 'CAN XEM XET'
    END

MEMBER [Measures].[Khuyen Nghi] AS
    CASE
        WHEN [Measures].[Xep Hang Theo Quy] <= 3 THEN 'TANG KHUYEN MAI'
        WHEN [Measures].[Xep Hang Theo Quy] <= 5 THEN 'AP DUNG COMBO'
        WHEN [Measures].[Xep Hang Theo Quy] <= 10 THEN 'TANG KEM'
        WHEN [Measures].[Doanh Thu] > 0 THEN 'DIEU CHINH GIA HOAC CONG THUC'
        ELSE 'NGUNG VA THAY THE SP MOI'
    END

SELECT
{
    [Measures].[Doanh Thu],
    [Measures].[Xep Hang Theo Quy],
    [Measures].[So Quy Trong Top10],
    [Measures].[Chien Luoc],
    [Measures].[Khuyen Nghi]
} ON COLUMNS,

GENERATE(
    [QuarterList],
    [DIM NGAY].[Quy].CurrentMember *
    TOPCOUNT(
        [ProductList],
        10,
        [Measures].[Tong Tien]
    )
) ON ROWS,

[BranchList] ON PAGES

FROM [QLCTB]

WHERE ([DIM NGAY].[Nam].[Nam].&[2025])
"""

# ==========================
# LẤY DỮ LIỆU
# ==========================

print("Đang kết nối SSAS...")

conn = win32com.client.Dispatch("ADODB.Connection")
conn.Open(conn_str)

rs = win32com.client.Dispatch("ADODB.Recordset")
rs.Open(mdx_query, conn)

data = []

while not rs.EOF:

    chi_nhanh = rs.Fields(0).Value
    quy = rs.Fields(1).Value
    san_pham = rs.Fields(2).Value

    doanh_thu = rs.Fields(3).Value
    xep_hang = rs.Fields(4).Value
    so_quy_top10 = rs.Fields(5).Value
    chien_luoc = rs.Fields(6).Value
    khuyen_nghi = rs.Fields(7).Value

    data.append([
        chi_nhanh,
        quy,
        san_pham,
        doanh_thu,
        xep_hang,
        so_quy_top10,
        chien_luoc,
        khuyen_nghi
    ])

    rs.MoveNext()

rs.Close()
conn.Close()

# ==========================
# DATAFRAME
# ==========================

df = pd.DataFrame(
    data,
    columns=[
        "ChiNhanh",
        "Quy",
        "SanPham",
        "DoanhThu",
        "XepHang",
        "SoQuyTop10",
        "ChienLuoc",
        "KhuyenNghi"
    ]
)

df["DoanhThu"] = pd.to_numeric(df["DoanhThu"], errors="coerce")
df = df.dropna(subset=["DoanhThu"])

print(f"Lấy được {len(df)} dòng dữ liệu.")

# ==========================
# CHUẨN BỊ TRỤC
# ==========================

# Tách sản phẩm theo chi nhánh để tránh chồng cột

df["SanPham_CN"] = (
    df["SanPham"]
    + " | "
    + df["ChiNhanh"]
)

quarters = sorted(df["Quy"].unique())
products_branch = sorted(df["SanPham_CN"].unique())
branches = sorted(df["ChiNhanh"].unique())

quarter_map = {
    q: i
    for i, q in enumerate(quarters)
}

product_map = {
    p: i
    for i, p in enumerate(products_branch)
}

colors = [
    '#3366cc',
    '#dc3912',
    '#ff9900',
    '#109618',
    '#990099',
    '#0099c6',
    '#dd4477',
    '#66aa00'
]

branch_color = {
    b: colors[i % len(colors)]
    for i, b in enumerate(branches)
}

# ==========================
# BIỂU ĐỒ 3D BAR CHART
# ==========================

fig = go.Figure()

legend_added = set()

for _, row in df.iterrows():

    x_center = quarter_map[row["Quy"]]

    y_center = product_map[
        row["SanPham_CN"]
    ]

    z_height = row["DoanhThu"] / 1_000_000

    r = 0.22

    x_box = [
        x_center-r, x_center+r, x_center+r, x_center-r,
        x_center-r, x_center+r, x_center+r, x_center-r
    ]

    y_box = [
        y_center-r, y_center-r, y_center+r, y_center+r,
        y_center-r, y_center-r, y_center+r, y_center+r
    ]

    z_box = [
        0, 0, 0, 0,
        z_height, z_height, z_height, z_height
    ]

    i_idx = [7,0,0,0,4,4,6,6,4,0,3,2]
    j_idx = [3,4,1,2,5,6,5,2,1,5,2,6]
    k_idx = [0,7,2,3,6,7,1,1,5,4,7,7]

    hover_text = (
        f"<b>Chi nhánh:</b> {row['ChiNhanh']}<br>"
        f"<b>Quý:</b> {row['Quy']}<br>"
        f"<b>Sản phẩm:</b> {row['SanPham']}<br>"
        f"<b>Doanh thu:</b> {row['DoanhThu']:,.0f} VNĐ<br>"
        f"<b>Xếp hạng:</b> {row['XepHang']}<br>"
        f"<b>Số quý Top10:</b> {row['SoQuyTop10']}<br>"
        f"<b>Chiến lược:</b> {row['ChienLuoc']}<br>"
        f"<b>Khuyến nghị:</b> {row['KhuyenNghi']}"
    )

    show_leg = row["ChiNhanh"] not in legend_added

    if show_leg:
        legend_added.add(row["ChiNhanh"])

    fig.add_trace(
        go.Mesh3d(
            x=x_box,
            y=y_box,
            z=z_box,
            i=i_idx,
            j=j_idx,
            k=k_idx,
            color=branch_color[row["ChiNhanh"]],
            opacity=0.95,
            flatshading=True,
            lighting=dict(
                ambient=0.7,
                diffuse=0.8,
                roughness=0.1,
                specular=0.2
            ),
            name=row["ChiNhanh"],
            text=hover_text,
            hoverinfo="text",
            showlegend=show_leg
        )
    )

# ==========================
# GIAO DIỆN
# ==========================

fig.update_layout(
    template="plotly_dark",

    title={
        "text": "TOP 10 SẢN PHẨM THEO DOANH THU NĂM 2025",
        "x": 0.5,
        "y": 0.96
    },

    scene=dict(

        xaxis=dict(
            title="QUÝ",
            tickmode="array",
            tickvals=list(range(len(quarters))),
            ticktext=quarters
        ),

        yaxis=dict(
            title="SẢN PHẨM | CHI NHÁNH",
            tickmode="array",
            tickvals=list(range(len(products_branch))),
            ticktext=products_branch,

            # nếu muốn hiện hết tên thì đổi thành True
            showticklabels=False
        ),

        zaxis=dict(
            title="DOANH THU (TRIỆU VNĐ)"
        ),

        aspectmode="manual",

        aspectratio=dict(
            x=1.8,
            y=4.0,
            z=1.2
        ),

        camera=dict(
            projection=dict(type="orthographic"),
            eye=dict(
                x=1.8,
                y=2.3,
                z=1.4
            )
        )
    ),

    height=1000
)

# ==========================
# XUẤT HTML
# ==========================

output_file = r"C:\Users\ADMIN\Desktop\Top10SanPham2025_3D.html"

fig.write_html(output_file)

print(f"\nĐã xuất file tại:\n{output_file}")

if os.path.exists(output_file):
    os.startfile(output_file)