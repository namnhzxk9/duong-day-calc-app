
# Tính sơ bộ dây dẫn, cột, móng đường dây 110kV–500kV

Bản cập nhật có thêm logic **đấu nối transit / cắt xen trên đường dây hiện trạng**.

## Điểm mới

- Có trường chọn kiểu đấu nối:
  - Đấu nối mới
  - Transit / cắt xen trên ĐD hiện trạng
- Có vùng nhập đường dây hiện trạng:
  - Cấp điện áp hiện trạng
  - Dây hiện trạng
  - Số mạch hiện trạng
  - Số dây/bó/pha hiện trạng
- Nếu chọn Transit, BOQ sẽ ưu tiên dùng dây hiện trạng thay vì dây chọn tự động.
- Có tab `00. Transit` để kiểm dòng tải, tiết diện và đồng bộ cấu hình.
- Có RFI riêng cho phương án transit.

## Cách chạy trên Mac

```bash
cd duong_day_calc_app_transit
chmod +x run_mac.sh
./run_mac.sh
```

Hoặc:

```bash
cd duong_day_calc_app_transit
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy Streamlit Cloud

Upload các file sau lên GitHub repository:

- app.py
- requirements.txt
- README.md
- run_mac.sh

Sau đó deploy trên Streamlit Cloud với `Main file path = app.py`.

## Lưu ý

Đây là công cụ estimate/check BOQ sơ bộ. Không thay thế hồ sơ thiết kế chính thức.
