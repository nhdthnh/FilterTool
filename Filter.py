from pkgutil import get_data
import re
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import Tk, Label, Button, filedialog, Entry, Frame, Toplevel, Canvas, Scrollbar, BOTH, LEFT, RIGHT, X, Y
from tkinter.messagebox import showinfo
from tkinter.ttk import Combobox, Treeview  # Loại bỏ Scrollbar và TopLevel từ ttk
from openpyxl import Workbook
from pyexcel_xls import get_data
from tkinter import ttk  # Import ttk để dùng ComboBox
# Biến toàn cục
file_path = ""
df = pd.DataFrame()
current_filtered_df = pd.DataFrame()  # Khởi tạo biến toàn cục

# Khởi tạo giao diện Tkinter
root = Tk()
root.title("Lọc dữ liệu Excel")
root.geometry("1000x600")
root.state('zoomed')  # Mở cửa sổ ở trạng thái maximize

# Hàm tải file Excel
# Hàm tải file Excel

def load_file():
    global file_path, df, current_filtered_df, is_right_visible_TonKho
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    try:
        if file_path.endswith(".xls"):
            # Kiểm tra file có phải định dạng .xls thực sự
            if not os.path.isfile(file_path):
                showinfo("Lỗi", "File không tồn tại!")
                return
            try:
                data = get_data(file_path)  # Đọc dữ liệu từ file .xls
                sheet_name = list(data.keys())[0]
                temp_file = os.path.splitext(file_path)[0] + "_converted.xlsx"
                
                # Ghi file tạm thời
                wb = Workbook()
                ws = wb.active
                ws.title = sheet_name
                for row in data[sheet_name]:
                    ws.append(row)
                wb.save(temp_file)

                # Đọc file .xlsx tạm thời
                df = pd.read_excel(temp_file, engine="openpyxl")
                os.remove(temp_file)  # Xóa file tạm thời sau khi đọc xong
            except Exception as e:
                showinfo("Lỗi", f"Không thể đọc file .xls: {e}")
                return
        elif file_path.endswith(".xlsx"):
            # Kiểm tra file có phải định dạng .xlsx thực sự
            try:
                df = pd.read_excel(file_path, engine="openpyxl")
            except Exception as e:
                showinfo("Lỗi", f"Không thể đọc file .xlsx: {e}")
                return
        else:
            showinfo("Lỗi", "Định dạng file không được hỗ trợ!")
            return

        if df.empty:
            showinfo("Lỗi", "File Excel không chứa dữ liệu!")
            return

        # Hiển thị dữ liệu
        headers = list(df.columns)
        if not headers:
            showinfo("Lỗi", "Không có cột nào được tìm thấy trong file!")
            return

        if is_right_visible_TonKho:
            combobox['values'] = headers
            combobox.current(0)
            update_treeview(df)
        else:
            combobox1['values'] = headers
            combobox1.current(0)
            update_treeview_DoanhThu(df)

    except Exception as e:
        showinfo("Lỗi", f"Không thể đọc file Excel: {e}")


def filter_data():
    global df, current_filtered_df, is_right_visible_TonKho
    if file_path == "":
        showinfo("Lỗi", "Vui lòng tải file Excel trước.")
        return
    if is_right_visible_TonKho:
        column = combobox.get()
        keyword = entry.get()
    else:
        column = combobox1.get()
        keyword = entry_Doanh_thu.get()
    
    if column == "" or keyword == "":
        showinfo("Lỗi", "Vui lòng chọn cột và nhập từ khóa.")
        return
    
    try:
        filtered_df = df[df[column].astype(str).str.contains(keyword, na=False, case=False)]
        current_filtered_df = filtered_df.copy()  # Lưu lại dữ liệu đã lọc vào biến toàn cục
        if is_right_visible_TonKho:
            update_treeview(filtered_df)  # Show "Tồn kho" columns
        elif is_right_visible_DoanhThu:
            update_treeview_DoanhThu(filtered_df)  # Show "Doanh thu" columns
    except Exception as e:
        showinfo("Lỗi", f"Không thể lọc dữ liệu: {e}")


def divide_stock():
    global current_filtered_df
    try:
        # Lấy giá trị chia từ ô nhập
        divisor = float(entry_divisor.get())
        if divisor == 0:
            showinfo("Lỗi", "Không thể chia cho 0.")
            return

        # Lấy đơn vị từ ComboBox
        unit = combobox_unit.get()

        # Chuyển đổi "In Stock" sang kg nếu đơn vị là g
        if unit == "g":
            divisor = divisor/1000

        # Tính giá trị Packs và làm tròn đến 3 chữ số thập phân
        current_filtered_df["Packs"] = (current_filtered_df["In Stock"] / divisor).round(3)

        # Đưa cột "Packs" vào ngay sau cột "In Stock"
        cols = list(current_filtered_df.columns)
        if "Packs" in cols:
            cols.insert(cols.index("In Stock") + 1, cols.pop(cols.index("Packs")))
        current_filtered_df = current_filtered_df[cols]

        # Cập nhật lại Treeview
        update_treeview(current_filtered_df)
    except ValueError:
        showinfo("Lỗi", "Vui lòng nhập một số hợp lệ để chia.")
    except Exception as e:
        showinfo("Lỗi", f"Không thể thực hiện chia: {e}")


def update_treeview(filtered_df):
    global current_filtered_df
    current_filtered_df = filtered_df.copy()  # Cập nhật dữ liệu hiện tại vào biến toàn cục

    # Xóa cột "STT" cũ nếu có
    if "STT" in filtered_df.columns:
        filtered_df = filtered_df.drop(columns=["STT"])

    # Thêm cột "STT" mới với giá trị tăng dần từ 1
    filtered_df.insert(0, "STT", range(1, len(filtered_df) + 1))  # Thêm cột STT

    # Loại bỏ các cột có tất cả giá trị là NaN
    filtered_df = filtered_df.dropna(axis=1, how='all')

    # Thay thế các giá trị NaN trong các ô có dữ liệu khác thành chuỗi rỗng
    filtered_df = filtered_df.fillna('')

    # Xóa các hàng và cột cũ trong Treeview
    for item in treeview.get_children():
        treeview.delete(item)

    # Kiểm tra nếu DataFrame rỗng
    if filtered_df.empty or filtered_df.shape[1] == 0:
        showinfo("Thông báo", "Không có cột nào có dữ liệu để hiển thị.")
        return

    treeview["columns"] = list(filtered_df.columns)
    treeview["show"] = "headings"

    # Cập nhật tiêu đề và kích thước cột
    for col in filtered_df.columns:
        max_width = max(filtered_df[col].astype(str).map(len).max(), len(col))
        calculated_width = max_width * 5  # Đặt chiều rộng cột lớn hơn
        treeview.heading(col, text=col)
        treeview.column(col, width=calculated_width, stretch=True)  # stretch=True giúp cuộn ngang hoạt động

    # Thêm dữ liệu vào Treeview
    for _, row in filtered_df.iterrows():
        treeview.insert("", "end", values=list(row))

def update_treeview_DoanhThu(filtered_df):
    global current_filtered_df
    current_filtered_df = filtered_df.copy()  # Cập nhật dữ liệu hiện tại vào biến toàn cục

    # Xóa cột "STT" cũ nếu có
    if "STT" in filtered_df.columns:
        filtered_df = filtered_df.drop(columns=["STT"])

    # Thêm cột "STT" mới với giá trị tăng dần từ 1
    filtered_df.insert(0, "STT", range(1, len(filtered_df) + 1))  # Thêm cột STT

    # Loại bỏ các cột có tất cả giá trị là NaN
    filtered_df = filtered_df.dropna(axis=1, how='all')

    # Thay thế các giá trị NaN trong các ô có dữ liệu khác thành 0
    filtered_df = filtered_df.fillna(0).astype({col: 'int' for col in filtered_df.select_dtypes(include=['float64']).columns})


    # Xác định các cột A đến E (thay thế bằng tên cột thực tế của bạn)
    columns_A_to_E = ['#', 'Item No.', 'Item Description', 'Annual Total']  # Thay bằng tên cột thực tế

    # Kiểm tra và lấy các cột A đến E có tồn tại trong DataFrame
    columns_A_to_E = [col for col in columns_A_to_E if col in filtered_df.columns]

    # Lọc các cột có chứa "Quantity" hoặc "Sales Amount"
    quantity_sales_columns = [
        col for col in filtered_df.columns 
        if re.search(r"(quantity|sales amount)", col, re.IGNORECASE)
    ]

    # Kết hợp các cột A-E với các cột Quantity và Sales Amount, loại bỏ trùng lặp
    filtered_columns = columns_A_to_E + [col for col in quantity_sales_columns if col not in columns_A_to_E]

    # Loại bỏ các cột có chứa "(Currency)" trong tên
    filtered_columns = [col for col in filtered_columns if "(currency)" not in col]

    # Thêm cột "STT" vào đầu danh sách
    filtered_columns.insert(0, "STT")

    # Kiểm tra nếu không có cột nào để hiển thị
    if not filtered_columns:
        showinfo("Thông báo", "Không có cột nào phù hợp để hiển thị.")
        return

    # Cài đặt các cột cho Treeview
    treeview["columns"] = filtered_columns
    treeview["show"] = "headings"

    # Cập nhật tiêu đề và kích thước cột
    for col in filtered_columns:
        if col not in filtered_df.columns and col != "STT":
            continue
        max_width = max(filtered_df[col].astype(str).map(len).max(), len(col))
        calculated_width = max_width * 5  # Điều chỉnh độ rộng theo nhu cầu
        treeview.heading(col, text=col)
        treeview.column(col, width=calculated_width, stretch=True)

    # Xóa dữ liệu cũ trong Treeview
    for item in treeview.get_children():
        treeview.delete(item)

    # Thêm dữ liệu vào Treeview chỉ cho các cột đã lọc
    for _, row in filtered_df[filtered_columns].iterrows():
        treeview.insert("", "end", values=list(row))


def filter_specific(value):
    global current_filtered_df, is_right_visible_TonKho, is_right_visible_DoanhThu
    if current_filtered_df.empty:
        showinfo("Lỗi", "Không có dữ liệu để lọc. Vui lòng thực hiện lọc dữ liệu ban đầu trước.")
        return
    try:
        exclude_value = ""
        if is_right_visible_TonKho:
            exclude_value = exclude_textbox.get().strip()
        if is_right_visible_DoanhThu:
            exclude_value = exclude_textbox_DoanhThu.get().strip()

        # Lọc theo giá trị chính (value)
        filtered_df = current_filtered_df[
            current_filtered_df['Item Description'].astype(str).str.contains(value, na=False, case=False)
        ]

        # Lọc theo ngoại trừ nếu có
        if exclude_value:
            # Tách các giá trị ngoại trừ bằng dấu phẩy
            exclude_values = [val.strip() for val in exclude_value.split(',') if val.strip()]
            for ex_val in exclude_values:
                filtered_df = filtered_df[
                    ~filtered_df['Item Description'].astype(str).str.contains(r'\b' + re.escape(ex_val) + r'\b', na=False, case=False)
                ]

        # Cập nhật giao diện
        if is_right_visible_TonKho:
            update_treeview(filtered_df)  # Cập nhật Tồn kho
        elif is_right_visible_DoanhThu:
            update_treeview_DoanhThu(filtered_df)  # Cập nhật Doanh thu

    except Exception as e:
        showinfo("Lỗi", f"Không thể lọc dữ liệu: {e}")



def get_treeview_data():
    # Trích xuất dữ liệu từ Treeview
    tree_data = []
    for item in treeview.get_children():
        tree_data.append(treeview.item(item)["values"])
    
    # Chuyển dữ liệu thành DataFrame
    columns = treeview["columns"]
    tree_df = pd.DataFrame(tree_data, columns=columns)
    return tree_df

def show_scrollable_charts():
    global root
    # Lấy dữ liệu từ Treeview
    tree_df = get_treeview_data()

    if tree_df.empty:
        showinfo("Lỗi", "Không có dữ liệu trong Treeview để hiển thị biểu đồ.")
        return

    try:
        # Chuyển các cột thành kiểu số
        tree_df['In Stock'] = pd.to_numeric(tree_df['In Stock'], errors='coerce')
        tree_df['Packs'] = pd.to_numeric(tree_df.get('Packs', pd.Series()), errors='coerce')

        # Tính tổng hàng tồn và Packs cho mỗi Item Description
        grouped_df = tree_df.groupby('Item Description').agg({'In Stock': 'sum', 'Packs': 'sum'})

        # Tạo cửa sổ mới cho biểu đồ
        chart_window = Toplevel(root)
        chart_window.title("Biểu đồ tồn kho và Packs")
        chart_window.state('zoomed')  # Mở cửa sổ ở trạng thái maximize

        # Canvas và Scrollbar
        canvas = Canvas(chart_window)
        scrollbar = Scrollbar(chart_window, orient="vertical", command=canvas.yview)
        scrollable_frame = Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Biểu đồ đầu tiên: In Stock
        fig1, ax1 = plt.subplots(figsize=(14, 8))
        grouped_df['In Stock'].plot(kind='bar', ax=ax1, color='skyblue')
        ax1.set_title("Biểu đồ Tồn kho (In Stock)", fontsize=14)
        ax1.set_ylabel("Số lượng Tồn kho", fontsize=12)
        ax1.set_xticklabels(grouped_df.index, rotation=45, ha='right', fontsize=10)
        ax1.bar_label(ax1.containers[0], fmt='%.2f', fontsize=8)
        plt.tight_layout()

        # Hiển thị biểu đồ In Stock trên Tkinter
        chart_canvas1 = FigureCanvasTkAgg(fig1, scrollable_frame)
        chart_canvas1.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig1)

        # Biểu đồ thứ hai: Packs (nếu có)
        if 'Packs' in grouped_df.columns:
            fig2, ax2 = plt.subplots(figsize=(14, 8))
            grouped_df['Packs'].plot(kind='bar', ax=ax2, color='orange')
            ax2.set_title("Biểu đồ Số gói (Packs)", fontsize=14)
            ax2.set_ylabel("Số lượng Packs", fontsize=12)
            ax2.set_xticklabels(grouped_df.index, rotation=45, ha='right', fontsize=10)
            ax2.bar_label(ax2.containers[0], fmt='%.2f', fontsize=8)
            plt.tight_layout()

            # Hiển thị biểu đồ Packs trên Tkinter
            chart_canvas2 = FigureCanvasTkAgg(fig2, scrollable_frame)
            chart_canvas2.get_tk_widget().pack(fill="both", expand=True)
            plt.close(fig2)

    except Exception as e:
        showinfo("Lỗi", f"Không thể hiển thị biểu đồ: {e}")


def show_scrollable_charts_DoanhThu():
    global root
    # Lấy dữ liệu từ Treeview
    tree_df = get_treeview_data()

    if tree_df.empty:
        showinfo("Lỗi", "Không có dữ liệu trong Treeview để hiển thị biểu đồ.")
        return

    try:
        # Bước 1: Trích xuất thông tin các cột liên quan đến Quantity và Sales Amount
        month_columns = {}
        for col in tree_df.columns:
            match = re.match(r'^(.*?)\s*-\s*(Quantity|Sales Amount)$', col)
            if match:
                month = match.group(1).strip()
                measure = match.group(2).strip()

                if month not in month_columns:
                    month_columns[month] = {}

                month_columns[month][measure] = col

        if not month_columns:
            showinfo("Lỗi", "Không tìm thấy cột 'Quantity' và 'Sales Amount' theo tháng.")
            return

        # Bước 2: Lặp qua từng sản phẩm và lấy dữ liệu
        chart_window = Toplevel(root)
        chart_window.title("Doanh Thu và Số Lượng Theo Tháng")
        chart_window.state('zoomed')  # Mở cửa sổ ở trạng thái maximize

        # Bước 3: Tạo Canvas và Scrollbar cho cửa sổ chính
        canvas = Canvas(chart_window)
        scrollbar = Scrollbar(chart_window, orient="vertical", command=canvas.yview)
        scrollable_frame = Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Thêm sự kiện cuộn chuột vào Canvas
        def on_mouse_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        # Bước 4: Lặp qua từng sản phẩm và vẽ biểu đồ
        for product_name in tree_df['Item Description'].unique():
            data = []
            for month, measures in month_columns.items():
                quantity_col = measures.get('Quantity')
                sales_amount_col = measures.get('Sales Amount')

                if not quantity_col or not sales_amount_col:
                    continue

                product_row = tree_df[tree_df['Item Description'] == product_name]

                for index, row in product_row.iterrows():
                    product_quantity = pd.to_numeric(row[quantity_col], errors='coerce')
                    product_sales = pd.to_numeric(row[sales_amount_col], errors='coerce')

                    if pd.notna(product_quantity) and pd.notna(product_sales):
                        data.append({
                            'Month': month,
                            'Quantity': product_quantity,
                            'Sales Amount': product_sales
                        })

            if not data:
                continue

            product_df = pd.DataFrame(data)

            product_df['Month_dt'] = pd.to_datetime(product_df['Month'], format='%B (%Y)', errors='coerce')
            invalid_dates = product_df['Month_dt'].isna()
            if invalid_dates.any():
                showinfo("Lỗi", f"Định dạng tháng không hợp lệ cho các giá trị: {product_df.loc[invalid_dates, 'Month'].tolist()}")
                product_df = product_df.dropna(subset=['Month_dt'])

            if product_df.empty:
                continue

            product_df = product_df.sort_values('Month_dt')

            fig, ax1 = plt.subplots(figsize=(14, 7))
            ax1.plot(product_df['Month_dt'], product_df['Quantity'], marker='o', color='blue', label='Quantity')
            ax1.set_xlabel('Tháng')
            ax1.set_ylabel('Quantity', color='blue')
            ax1.tick_params(axis='y', labelcolor='blue')

            for i, row in product_df.iterrows():
                ax1.text(row['Month_dt'], row['Quantity'], f"{row['Quantity']:.2f}", color='blue', ha='center', va='bottom', fontsize=8)

            ax1.xaxis.set_major_formatter(plt.FixedFormatter(product_df['Month_dt'].dt.strftime('%b %Y')))
            plt.setp(ax1.get_xticklabels(), rotation=45, ha='right', rotation_mode='anchor')

            ax2 = ax1.twinx()
            ax2.plot(product_df['Month_dt'], product_df['Sales Amount'], marker='o', color='red', label='Sales Amount')
            ax2.set_ylabel('Sales Amount', color='red')
            ax2.tick_params(axis='y', labelcolor='red')

            for i, row in product_df.iterrows():
                ax2.text(row['Month_dt'], row['Sales Amount'], f"${row['Sales Amount']:.2f}", color='red', ha='center', va='bottom', fontsize=10)

            plt.title(f"Quantity và Sales Amount theo Tháng - {product_name}")

            lines_1, labels_1 = ax1.get_legend_handles_labels()
            lines_2, labels_2 = ax2.get_legend_handles_labels()
            ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper left')
            plt.tight_layout()

            chart_canvas = FigureCanvasTkAgg(fig, scrollable_frame)
            chart_canvas.draw()
            chart_canvas.get_tk_widget().pack(fill=BOTH, expand=True)
            plt.close(fig)

    except Exception as e:
        showinfo("Lỗi", f"Không thể hiển thị biểu đồ: {e}")




def export_to_excel():
    # Lấy dữ liệu từ Treeview
    tree_df = get_treeview_data()

    if tree_df.empty:
        showinfo("Lỗi", "Không có dữ liệu trong Treeview để xuất.")
        return

    try:
        # Hộp thoại lưu file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu file Excel"
        )
        
        if not file_path:  # Người dùng hủy lưu file
            return

        # Xuất dữ liệu ra file Excel
        tree_df.to_excel(file_path, index=False, engine="openpyxl")
        showinfo("Thông báo", f"Xuất file Excel thành công!\nLưu tại: {file_path}")
    except Exception as e:
        showinfo("Lỗi", f"Không thể xuất file Excel: {e}")

# Variable to track the visibility state
is_right_visible_TonKho = False
is_right_visible_DoanhThu = False


# Function to toggle the visibility of the "Tồn kho" section
def toggle_right_section():
    global is_right_visible_TonKho, is_right_visible_DoanhThu
    if is_right_visible_TonKho:
        right_header_TonKho.pack_forget()  # Hide the right section
        toggle_button.config(text="Tồn kho")
    else:
        # Hide "Doanh thu" section if it's visible
        if is_right_visible_DoanhThu:
            right_header_DoanhThu.pack_forget()
            revenue_button.config(text="Doanh thu")
            is_right_visible_DoanhThu = False
        
        # Show "Tồn kho" section
        right_header_TonKho.pack(side="right", padx=10, fill="x", expand=True)
        toggle_button.config(text="Hide")
    
    # Update visibility flag
    is_right_visible_TonKho = not is_right_visible_TonKho

# Function to toggle the visibility of the "Doanh thu" section
def toggle_revenue_section():
    global is_right_visible_DoanhThu, is_right_visible_TonKho
    if is_right_visible_DoanhThu:
        right_header_DoanhThu.pack_forget()  # Hide the right section
        revenue_button.config(text="Doanh thu")
    else:
        # Hide "Tồn kho" section if it's visible
        if is_right_visible_TonKho:
            right_header_TonKho.pack_forget()
            toggle_button.config(text="Tồn kho")
            is_right_visible_TonKho = False
        
        # Show "Doanh thu" section
        right_header_DoanhThu.pack(side="right", padx=10, fill="x", expand=True)
        revenue_button.config(text="Hide")
    
    # Update visibility flag
    is_right_visible_DoanhThu = not is_right_visible_DoanhThu


# Header frame divided into two parts
header_frame = Frame(root)
header_frame.pack(fill="x", pady=5)

# Left section for the toggle buttons
left_header = Frame(header_frame)
left_header.pack(side="left", padx=10, fill="x", expand=True)

toggle_button = Button(left_header, text="Tồn kho", command=toggle_right_section)
toggle_button.pack(side="left", padx=5, pady=5)

# New "Doanh thu" button with an empty command for now
revenue_button = Button(left_header, text="Doanh thu", command=toggle_revenue_section)
revenue_button.pack(side="left", padx=5, pady=5)

# Right section for buttons and input fields using grid
right_header_TonKho = Frame(header_frame)
Label(right_header_TonKho, text="Tồn kho").grid(row=0, column=2, padx=5)
Label(right_header_TonKho, text="Chọn cột cần lọc").grid(row=1, column=0, padx=5)
combobox = Combobox(right_header_TonKho, state="readonly", width=30)
combobox.grid(row=1, column=1, padx=5)
Button(right_header_TonKho, text="Tải file Excel", command=load_file).grid(row=1, column=2, padx=5)
Label(right_header_TonKho, text="Nhập từ khóa cần lọc").grid(row=2, column=0, padx=5)
entry = Entry(right_header_TonKho, width=30)
entry.grid(row=2, column=1, padx=5)
Button(right_header_TonKho, text="Lọc dữ liệu", command=filter_data).grid(row=2, column=2, padx=5)
filter_frame = Frame(root)
filter_frame.pack(pady=10)
Label(right_header_TonKho, text="Bộ lọc phụ").grid(row=3, column=0, padx=5)
entry1 = Entry(right_header_TonKho, width=30)
entry1.grid(row=3, column=1, padx=5)
exclude_label = Label(right_header_TonKho, text="Ngoại trừ: ")
exclude_label.grid(row=3, column=2, padx=5)
exclude_textbox = Entry(right_header_TonKho, width=30)
exclude_textbox.grid(row=3, column=3, padx=5)
Button(right_header_TonKho, text="Lọc dữ liệu", command=lambda:filter_specific(entry1.get()) ).grid(row=3, column=4, padx=5)
Label(right_header_TonKho, text="Nhập số cần chia").grid(row=4, column=0, padx=5)
entry_divisor = Entry(right_header_TonKho, width=30)
entry_divisor.grid(row=4, column=1, padx=5)
combobox_unit = ttk.Combobox(right_header_TonKho, width=3, state="readonly")
combobox_unit['values'] = ("kg", "g")  # Các đơn vị
combobox_unit.grid(row=4, column=2, padx=5)
combobox_unit.current(0)  # Mặc định chọn "kg"
Button(right_header_TonKho, text="Chia số gói", command=divide_stock).grid(row=4, column=3, padx=5)
Button(right_header_TonKho, text="Hiển thị biểu đồ", command=lambda: show_scrollable_charts()).grid(row=5, column=0, padx=5)
Button(right_header_TonKho, text="Xuất file Excel", command=export_to_excel).grid(row=5, column=2, padx=5)
right_header_TonKho.pack_forget()


right_header_DoanhThu = Frame(header_frame)
Label(right_header_DoanhThu, text="Doanh thu").grid(row=0, column=2, padx=5)
Label(right_header_DoanhThu, text="Chọn cột cần lọc").grid(row=1, column=0, padx=5)
combobox1 = Combobox(right_header_DoanhThu, state="readonly", width=30)
combobox1.grid(row=1, column=1, padx=5)
Button(right_header_DoanhThu, text="Tải file Excel", command=load_file).grid(row=1, column=2, padx=5)
Label(right_header_DoanhThu, text="Nhập từ khóa cần lọc").grid(row=2, column=0, padx=5)
entry_Doanh_thu = Entry(right_header_DoanhThu, width=30)
entry_Doanh_thu.grid(row=2, column=1, padx=5)
Button(right_header_DoanhThu, text="Lọc dữ liệu", command=filter_data).grid(row=2, column=2, padx=5)
filter_frame_DoanhThu = Frame(root)
filter_frame_DoanhThu.pack(pady=10)
Label(right_header_DoanhThu, text="Bộ lọc phụ").grid(row=3, column=0, padx=5)
entry1_Doanh_Thu = Entry(right_header_DoanhThu, width=30)
entry1_Doanh_Thu.grid(row=3, column=1, padx=5)
exclude_label_DoanhThu = Label(right_header_DoanhThu, text="Ngoại trừ: ")
exclude_label_DoanhThu.grid(row=3, column=2, padx=5)
exclude_textbox_DoanhThu = Entry(right_header_DoanhThu, width=30)
exclude_textbox_DoanhThu.grid(row=3, column=3, padx=5)
Button(right_header_DoanhThu, text="Lọc dữ liệu", command=lambda:filter_specific(entry1_Doanh_Thu.get()) ).grid(row=3, column=4, padx=5)
Button(right_header_DoanhThu, text="Hiển thị biểu đồ", command=lambda: show_scrollable_charts_DoanhThu()).grid(row=4, column=0, padx=5)
Button(right_header_DoanhThu, text="Xuất file Excel", command=export_to_excel).grid(row=4, column=2, padx=5)
right_header_DoanhThu.pack_forget()

# Treeview below the header
tree_frame = Frame(root)
tree_frame.pack(expand=True, fill="both")

# Vertical scrollbar
tree_vertical = Scrollbar(tree_frame, orient="vertical")
tree_vertical.pack(side="right", fill="y")

# Horizontal scrollbar
tree_horizontal = Scrollbar(tree_frame, orient="horizontal")
tree_horizontal.pack(side="bottom", fill="x")

# Treeview with scrollbars
treeview = Treeview(
    tree_frame,
    yscrollcommand=tree_vertical.set,
    xscrollcommand=tree_horizontal.set,
)

treeview.pack(side="left", expand=True, fill="both")
tree_vertical.config(command=treeview.yview)
tree_horizontal.config(command=treeview.xview)





# Khởi động giao diện
root.mainloop()
