# Add-in Chuẩn hóa Thể thức Văn bản (Microsoft Word)

Chuẩn hóa thể thức văn bản theo **Nghị định 30/2020/NĐ-CP** (văn bản hành chính), thể thức **văn bản của Đảng**, khối Công an, Quân đội, MTTQ và **văn bản quy phạm pháp luật**.

- Người duy trì: **Vương Đình Anh Tú**
- Dựa trên add-in gốc của **Văn phòng Đảng ủy xã Thạnh Phú** (phiên bản 2026.11.07.5), phát hành lại với sự đồng ý của tác giả.

## Cài đặt (đồng nghiệp)

1. Tải **[Bo_Cai_Dat_AddIn.zip](https://github.com/anhtu604/addin-the-thuc-van-ban/raw/main/Bo_Cai_Dat_AddIn.zip)** rồi giải nén.
2. Lưu và đóng hết Word.
3. Nhấp đúp `CaiDat.bat`.
4. Mở Word. Tab **Home** có nhóm **THỂ THỨC VĂN BẢN**.

Chi tiết xem `HuongDan.txt` trong gói cài.

## Cập nhật

Add-in tự đọc [`version.json`](version.json) mỗi ngày một lần và hỏi trước khi cài bản mới. File tải về được kiểm tra SHA-256. Có thể bấm nút **Cập nhật** trên Ribbon bất cứ lúc nào.

## Điểm khác so với bản gốc

- Sửa lỗi Quyết định bị gom phần căn cứ và các Điều vào ô "Nơi nhận".
- Có đường kẻ dưới trích yếu; dòng thẩm quyền ban hành in hoa, đậm, canh giữa.
- Tự nhận đúng khối Đảng hoặc Nhà nước kể cả khi đã chọn cố định khối.
- Hỗ trợ văn bản quy phạm pháp luật (Phần, Chương, Mục, Điều, khoản, điểm).
- Chữ ký không đứng một mình ở trang mới, không bị cắt đôi; tự ép trang trong khung NĐ 30.
- Sửa lỗi dấu cách khi dán; bảng số liệu lặp dòng tiêu đề khi sang trang.
- Một lần Ctrl+Z hoàn tác toàn bộ lần chuẩn hóa.
- Không còn đăng ký máy với máy chủ bản quyền của bản gốc.

## Phát hành bản mới (người duy trì)

Mã nguồn nằm trong `ma_nguon/`. `Module1.bas` được tạo từ mã gốc qua các bản vá đánh số:

```powershell
cd ma_nguon
python va_1_the_thuc.py   Module1_ban_goc_2026.11.07.5.bas a.bas
python va_2_ban_rieng.py  a.bas b.bas
python va_3_tien_ich.py   b.bas c.bas
python va_4_phat_hanh.py  c.bas Module1.bas anhtu604/addin-the-thuc-van-ban <PHIEN_BAN_MOI>
# Word > Trust Center > bật "Trust access to the VBA project object model" rồi:
.\nap_ma_vao_dotm.ps1 -Src khung_goc_2026.11.07.5.dotm -Bas Module1.bas -Out tam.dotm
python va_5_ribbon.py tam.dotm ..\ChuanHoa_VanBan_Rieng.dotm
cd ..
.\phat_hanh.ps1 -Dotm .\ChuanHoa_VanBan_Rieng.dotm -Version <PHIEN_BAN_MOI> -Notes "Noi dung thay doi (khong dau)"
```

`phat_hanh.ps1` chép file vào `dist/` và `bo_cai_dat/`, đóng gói lại `Bo_Cai_Dat_AddIn.zip`, tính SHA-256, ghi `version.json` rồi commit và push. Số phiên bản phải tăng dần, dạng `2026.11.09.3`.
