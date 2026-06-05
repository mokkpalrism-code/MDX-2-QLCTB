import plotly.graph_objects as go
import pandas as pd
import win32com.client
import os

server_name = r'MERINGUE\ADMIN'
database_name = 'MultidimensionalProject4'
conn_str = f"Provider=MSOLAP;Data Source={server_name};Initial Catalog={database_name};"

mdx_query = """
WITH 
MEMBER [Measures].[BienDo_DaoDong_Gia] AS 
    [Measures].[Maximum Gia Nhap] - [Measures].[Minimum Gia Nhap]
MEMBER [Measures].[Gia_Trung_Binh] AS 
    ([Measures].[Maximum Gia Nhap] + [Measures].[Minimum Gia Nhap]) / 2

SELECT 
    {  
        [Measures].[Maximum Gia Nhap], 
        [Measures].[Minimum Gia Nhap],
        [Measures].[BienDo_DaoDong_Gia],
        [Measures].[Gia_Trung_Binh]
    } ON COLUMNS,
    NON EMPTY 
    [DIM NGUYEN LIEU].[Ten NL].[Ten NL].Members *
    [DIM NHA CUNG CAP].[Ten NCC].[Ten NCC].Members *
    [DIM NGAY].[Quy].[Quy].Members 
    ON ROWS
FROM [QLCTB]
"""

print("--- Đang gửi lệnh MDX vào Cube...")

try:
    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open(conn_str)
    rs = win32com.client.Dispatch("ADODB.Recordset")
    rs.Open(mdx_query, conn)

    print(f"Số cột trả về: {rs.Fields.Count}")  

    data = []
    while not rs.EOF:
        nguyen_lieu = rs.Fields(0).Value
        nha_cung_cap = rs.Fields(1).Value
        quy = rs.Fields(2).Value
        max_gia = rs.Fields(3).Value
        min_gia = rs.Fields(4).Value
        bien_do = rs.Fields(5).Value
        gia_tb = rs.Fields(6).Value

        data.append([nguyen_lieu, nha_cung_cap, quy, gia_tb, max_gia, min_gia, bien_do])
        rs.MoveNext()

    rs.Close()
    conn.Close()

    df = pd.DataFrame(data, columns=["NguyenLieu", "NhaCungCap", "ThoiGian",
                                     "AvgPrice", "MaxPrice", "MinPrice", "BienDo"])
    print(f"Quao tui đã lấy được {len(df)} dòng.")

except Exception as e:
    print(f"Lỗi kết nối hoặc truy vấn: {e}")
    df = pd.DataFrame()

if not df.empty:
    # Chuyển đổi số (xử lý dấu phẩy)
    for col in ["AvgPrice", "MaxPrice", "MinPrice", "BienDo"]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')

    # Đơn vị K VND
    df["BienDo_K"] = df["BienDo"] / 1000
    df["AvgPrice_K"] = df["AvgPrice"] / 1000
    df = df.dropna(subset=["BienDo_K"])

    # In top 10 để kiểm tra
    top10 = df[["NguyenLieu", "BienDo_K"]].drop_duplicates().sort_values("BienDo_K", ascending=False).head(10)
    print("\n--- TOP 10 NGUYÊN LIỆU CÓ BIÊN ĐỘ DAO ĐỘNG GIÁ NHẬP LỚN NHẤT NĂM 2025 (K VND) ---")
    print(top10.to_string(index=False))

    all_time = sorted(df["ThoiGian"].unique())
    all_nl = sorted(df["NguyenLieu"].unique())

    df["Nhom"] = df["NguyenLieu"] + " - " + df["NhaCungCap"]

    df = df.groupby(["Nhom", "NguyenLieu", "NhaCungCap", "ThoiGian"], as_index=False).agg({
        "BienDo_K": "mean",
        "AvgPrice_K": "mean",
        "MaxPrice": "max",
        "MinPrice": "min"
    })

    nhom_bien_do = df.groupby("Nhom")["BienDo_K"].max().sort_values(ascending=False)
    nhom_list_sorted = nhom_bien_do.index.tolist()

    colors = ['#3366cc', '#dc3912', '#ff9900', '#109618', '#990099', '#0099c6', '#dd4477', '#66aa00',
              '#b82e2e', '#316395', '#994499', '#22aa99', '#aaaa11', '#6633cc', '#e67300', '#8b0707']
    nhom_color_map = {nhom: colors[i % len(colors)] for i, nhom in enumerate(nhom_list_sorted)}

    fig = go.Figure()
    added_legend = set()

    print("\n--- Tui đang vất vả xây biểu đồ 3D cho bạn nè...")

    for nhom in nhom_list_sorted:
        df_nhom = df[df["Nhom"] == nhom]
        for _, row in df_nhom.iterrows():
            x_idx = all_time.index(row["ThoiGian"])
            y_idx = all_nl.index(row["NguyenLieu"])
            z_val = row["BienDo_K"]

            r = 0.15
            x_box = [x_idx-r, x_idx+r, x_idx+r, x_idx-r, x_idx-r, x_idx+r, x_idx+r, x_idx-r]
            y_box = [y_idx-r, y_idx-r, y_idx+r, y_idx+r, y_idx-r, y_idx-r, y_idx+r, y_idx+r]
            z_box = [0, 0, 0, 0, z_val, z_val, z_val, z_val]

            i_idx = [7, 0, 0, 0, 4, 4, 6, 6, 4, 0, 3, 2]
            j_idx = [3, 4, 1, 2, 5, 6, 5, 2, 1, 5, 2, 6]
            k_idx = [0, 7, 2, 3, 6, 7, 1, 1, 5, 4, 7, 7]

            hover_text = (
                f"🌾 <b>Nguyên liệu:</b> {row['NguyenLieu']}<br>"
                f"📦 <b>Nhà cung cấp:</b> {row['NhaCungCap']}<br>"
                f"⏳ <b>Quý:</b> {row['ThoiGian']}<br>"
                f"🔥 <b>BIÊN ĐỘ:</b> {z_val:.2f} K VND<br>"
                f"💰 <b>Giá TB:</b> {row['AvgPrice_K']:.2f} K VND"
            )

            show_leg = nhom not in added_legend
            if show_leg:
                added_legend.add(nhom)

            fig.add_trace(go.Mesh3d(
                x=x_box, y=y_box, z=z_box,
                i=i_idx, j=j_idx, k=k_idx,
                color=nhom_color_map[nhom],
                name=nhom,
                text=hover_text,
                hoverinfo="text",
                flatshading=True,
                lighting=dict(ambient=0.7, diffuse=0.8, roughness=0.1, specular=0.2),
                showlegend=show_leg
            ))

    fig.update_layout(
        template="plotly_dark",
        title={
            'text': "BIỂU ĐỒ BIẾN ĐỘNG GIÁ NHẬP NGUYÊN VẬT LIỆU NĂM 2025 CỦA CHUỖI TIỆM BÁNH THE 350F",
            'y': 0.96, 'x': 0.5, 'xanchor': 'center', 'yanchor': 'top'
        },
        scene=dict(
            xaxis=dict(
                title=dict(text="QUÝ"),
                tickmode='array', tickvals=list(range(len(all_time))), ticktext=all_time,
                backgroundcolor="rgb(20,20,20)", gridcolor="rgb(50,50,50)", showbackground=True
            ),
            yaxis=dict(
                title=dict(text="NGUYÊN LIỆU"),
                tickmode='array', tickvals=list(range(len(all_nl))), ticktext=all_nl,
                showticklabels=False,    # Ẩn tên nguyên liệu trên trục y
                tickangle=45,
                backgroundcolor="rgb(25,25,25)", gridcolor="rgb(50,50,50)", showbackground=True
            ),
            zaxis=dict(
                title=dict(text="Biên độ dao động giá (K VND)"),
                backgroundcolor="rgb(30,30,30)", gridcolor="rgb(50,50,50)", showbackground=True
            ),
            aspectmode='manual',
            aspectratio=dict(x=2.0, y=1.8, z=1.0),
            camera=dict(
                projection=dict(type="orthographic"),
                eye=dict(x=1.8, y=1.8, z=1.2)
            )
        ),
        margin=dict(l=30, r=30, b=30, t=70),
        height=950
    )

    output_file = r"C:\Users\ADMIN\Desktop\MDX2.html"
try:
    fig.write_html(output_file)
    print(f"Chúc mừng bạn đã ghi file thành công tại: {output_file}")
    import os
    if os.path.exists(output_file):
        print(f"\n --- Đang bưng bạn đến chỗ biểu đồ nà...")
        os.startfile(output_file)   
    else:
        print("File không tồn tại dù đã ghi, làm ơn kiểm tra quyền thư mục giùm.")
except Exception as e:
    print(f"Lỗi khi ghi file: {e}")