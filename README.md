
# Tính sơ bộ dây dẫn, cột, móng đường dây 110kV–500kV

Ứng dụng Python giao diện Streamlit để tính nhanh:
- Dòng tải và chọn dây dẫn sơ bộ.
- Tổn thất điện áp sơ bộ.
- Số khoảng cột, số cột sơ bộ.
- Khối lượng thép cột.
- Khối lượng móng, bê tông, cốt thép, bu lông neo, tiếp địa.
- BOQ tổng hợp và xuất Excel.

## Cách chạy trên Mac

### 1. Cài Python
Khuyến nghị Python 3.10 trở lên.

Kiểm tra:
```bash
python3 --version
```

### 2. Tạo môi trường ảo
```bash
cd duong_day_calc_app
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Cài thư viện
```bash
pip install -r requirements.txt
```

### 4. Chạy app
```bash
streamlit run app.py
```

Sau đó trình duyệt sẽ mở giao diện tại:
```text
http://localhost:8501
```

## Lưu ý kỹ thuật

Đây là công cụ tính sơ bộ/BOQ/check logic. Không thay thế:
- PLS-CADD, PLS-TOWER, PLS-POLE.
- Phụ lục tính toán cơ lý dây.
- Tính toán kết cấu cột/móng chính thức.
- Hồ sơ thiết kế TKKT/TKBVTC được phê duyệt.
