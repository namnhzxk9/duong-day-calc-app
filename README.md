# BOQ sơ bộ đường dây 110kV–500kV - Streamlit App v3

## Cải tiến bản v3

- Sửa lỗi export Excel không đúng với BOQ đang hiển thị.
- Nút export đặt sau khi BOQ được tính xong, xuất đúng sheet `06_BOQ_Tong_Hop`.
- Thêm chế độ tính số cột: theo khoảng cột trung bình hoặc nhập tổng số cột thực tế.
- Phân bổ cột bằng thuật toán largest remainder để tổng số cột không bị lệch do làm tròn.
- Với transit/cắt xen, dây BOQ ưu tiên theo dây hiện trạng.
- Thêm nguồn số liệu/Assumption để hạn chế dùng nhầm dữ liệu mặc định.

## Cách chạy trên Mac

```bash
cd duong_day_calc_app_v3
chmod +x run_mac.sh
./run_mac.sh
```

Hoặc:

```bash
cd duong_day_calc_app_v3
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy Streamlit Cloud

Upload/replace vào GitHub:

- app.py
- requirements.txt
- README.md
- run_mac.sh

Main file path: `app.py`.

## Lưu ý

Đây là công cụ estimate/check BOQ sơ bộ, không thay thế phụ lục tính toán thiết kế chính thức.
