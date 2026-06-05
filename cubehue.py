import plotly.graph_objects as go
import pandas as pd
import win32com.client

server_name = r'MERINGUE\ADMIN'
database_name = 'MultidimensionalProject4'

conn_str = f"Provider=MSOLAP;Data Source={server_name};Initial Catalog={database_name};"

mdx_query = """
WITH 
MEMBER [Measures].[GiaNhap_TrungBinh] AS 
    AVG(EXISTING [DIM NGAY].[Ngay].[Ngay].Members, [Measures].[Gia Nhap])
MEMBER [Measures].[GiaNhap_CaoNhat] AS 
    MAX(EXISTING [DIM NGAY].[Ngay].[Ngay].Members, [Measures].[Gia Nhap])
MEMBER [Measures].[GiaNhap_ThapNhat] AS 
    MIN(EXISTING [DIM NGAY].[Ngay].[Ngay].Members, [Measures].[Gia Nhap])
MEMBER [Measures].[BienDo_DaoDong_Gia] AS 
    [Measures].[GiaNhap_CaoNhat] - [Measures].[GiaNhap_ThapNhat]

SELECT 
    {  
        [Measures].[GiaNhap_TrungBinh], 
        [Measures].[GiaNhap_CaoNhat], 
        [Measures].[GiaNhap_ThapNhat],
        [Measures].[BienDo_DaoDong_Gia]
    } ON COLUMNS,
    NON EMPTY (
        [DIM CHI NHANH].[Ten CN].[Ten CN].Members *
        [DIM NHA CUNG CAP].[Ten NCC].[Ten NCC].Members *
        [DIM NGUYEN LIEU].[Ten NL].[Ten NL].Members *
        [DIM NGAY].[Quy].[Quy].Members 
    ) ON ROWS
FROM [QLCTB]
"""

print(" --- Đang gửi lệnh MDX vào SSAS Cube...")

try:
    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open(conn_str)
    
    rs = win32com.client.Dispatch("ADODB.Recordset")
    rs.Open(mdx_query, conn)
    
    data = []
    while not rs.EOF:
        chi_nhanh = rs.Fields(0).Value
        nha_cung_cap = rs.Fields(1).Value
        nguyen_lieu = rs.Fields(2).Value
        thang = rs.Fields(3).Value
        avg_price = rs.Fields(4).Value
        max_price = rs.Fields(5).Value
        min_price = rs.Fields(6).Value
        bien_do = rs.Fields(7).Value
        
        data.append([chi_nhanh, nha_cung_cap, nguyen_lieu, thang, avg_price, max_price, min_price, bien_do])
        rs.MoveNext()
        
    rs.Close()
    conn.Close()
    
    df = pd.DataFrame(data, columns=["ChiNhanh", "NhaCungCap", "NguyenLieu", "ThoiGian", 
                                     "AvgPrice", "MaxPrice", "MinPrice", "BienDo"])
    print(" --- Đã trích xuất dữ liệu thành công <3 !")

except Exception as e:
    print(f" --- Chúc mừng bạn nhé's, bạn đã dính lỗi kết nối hoặc lỗi truy vấn MDX: {e}")
    df = pd.DataFrame()


if not df.empty:
    # Chuyển sang đơn vị K VND
    df["AvgPrice_K"] = pd.to_numeric(df["AvgPrice"], errors='coerce') / 1000
    df["MaxPrice_K"] = pd.to_numeric(df["MaxPrice"], errors='coerce') / 1000
    df["MinPrice_K"] = pd.to_numeric(df["MinPrice"], errors='coerce') / 1000
    df["BienDo_K"] = pd.to_numeric(df["BienDo"], errors='coerce') / 1000
    
    # Loại bỏ dữ liệu trống dựa trên cột Biên Độ
    df = df.dropna(subset=["BienDo_K"])
    
    all_time = sorted(df["ThoiGian"].unique())
    all_nl = list(df["NguyenLieu"].unique())
    
    # Phân loại màu theo cặp Chi nhánh + Nhà cung cấp
    df["Nhom"] = df["ChiNhanh"] + " - " + df["NhaCungCap"]
    nhom_list = df["Nhom"].unique()
    colors = ['#3366cc', '#dc3912', '#ff9900', '#109618', '#990099', '#0099c6', '#dd4477', '#66aa00', 
              '#b82e2e', '#316395', '#994499', '#22aa99', '#aaaa11', '#6633cc', '#e67300', '#8b0707']
    nhom_color_map = {nhom: colors[i % len(colors)] for i, nhom in enumerate(nhom_list)}
    
    fig = go.Figure()
    added_legend = set()
    
    print(" --- Mình đang xây biểu đồ cho bạn nè, bạn đợi mình tí nhé...")
    
    for nhom in nhom_list:
        df_nhom = df[df["Nhom"] == nhom]
        chi_nhanh_display = nhom.split(" - ")[0]
        ncc_display = nhom.split(" - ")[1]
        
        for _, row in df_nhom.iterrows():
            x_idx = all_time.index(row["ThoiGian"])
            y_idx = all_nl.index(row["NguyenLieu"])
            
            # ĐÃ ĐỔI: Chiều cao trục Z thể hiện BIÊN ĐỘ DAO ĐỘNG
            z_val = row["BienDo_K"]  
            
            r = 0.15
            x_box = [x_idx-r, x_idx+r, x_idx+r, x_idx-r, x_idx-r, x_idx+r, x_idx+r, x_idx-r]
            y_box = [y_idx-r, y_idx-r, y_idx+r, y_idx+r, y_idx-r, y_idx-r, y_idx+r, y_idx+r]
            z_box = [0, 0, 0, 0, z_val, z_val, z_val, z_val]
            
            i_idx = [7, 0, 0, 0, 4, 4, 6, 6, 4, 0, 3, 2]
            j_idx = [3, 4, 1, 2, 5, 6, 5, 2, 1, 5, 2, 6]
            k_idx = [0, 7, 2, 3, 6, 7, 1, 1, 5, 4, 7, 7]
            
            # Sắp xếp lại bảng Hover thông tin hiển thị khi di chuột vào
            hover_text = (
                f"🏢 <b>Chi nhánh:</b> {chi_nhanh_display}<br>"
                f"📦 <b>Nhà cung cấp:</b> {ncc_display}<br>"
                f"🌾 <b>Nguyên liệu:</b> {row['NguyenLieu']}<br>"
                f"⏳ <b>Thời gian:</b> {row['ThoiGian']}<br>"
                f"🔥 <b>BIÊN ĐỘ DAO ĐỘNG:</b> {z_val:.2f} K VND<br>"
                f"💰 <b>Giá nhập trung bình:</b> {row['AvgPrice_K']:.2f} K VND<br>"
                f"📈 <b>Giá cao nhất:</b> {row['MaxPrice_K']:.2f} K VND<br>"
                f"📉 <b>Giá thấp nhất:</b> {row['MinPrice_K']:.2f} K VND"
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
            'text': "BIỂU ĐỒ TRỰC QUAN HÓA DỮ LIỆU BIẾN ĐỘNG GIÁ NHẬP NGUYÊN VẬT LIỆU NĂM 2025 CỦA CHUỖI TIỆM BÁNH THE 350F",
            'y': 0.96, 'x': 0.5, 'xanchor': 'center', 'yanchor': 'top'
        },
        dragmode="turntable",
        scene=dict(
            xaxis=dict(
                title=dict(text="QUÝ"),
                tickmode='array', tickvals=list(range(len(all_time))), ticktext=all_time,
                backgroundcolor="rgb(20,20,20)", gridcolor="rgb(50,50,50)", showbackground=True
            ),
            yaxis=dict(
                title=dict(text="NGUYÊN LIỆU (vì lý do thẩm mỹ nên đã lược bớt)"),
                tickmode='array', tickvals=list(range(len(all_nl))), ticktext=all_nl,
                showticklabels=False,  # Nếu muốn hiện tên nguyên liệu trên trục Y, đổi thành True
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
                eye=dict(x=1.8, y=1.8, z=1.2)  # Đã loại bỏ toán tử lỗi gán lồng := ở đây
            )
        ),
        margin=dict(l=30, r=30, b=30, t=70),
        height=950
    )
    
    output_file = "index.html"
    
    fig.write_html(output_file)
    fig.show()
    print(f"\n --- Đang bưng bạn đến chỗ biểu đồ nà...")
else:
    print(" --- ERROR #404. Hãy kiểm tra lại kết nối Cube hoặc cấu trúc câu lệnh MDX hoặc tất cả.")