EXERCISE SUMARISE  
**PHẦN 1: HOÀN THIỆN SRS BẰNG HEADING STYLES VÀ AUTO TOC**  
[1. Giới thiệu dự án1  
   
 ](#anchor-1 "#anchor-1")[1.1.Mục tiêu1  
   
 ](#anchor-2 "#anchor-2")[1.2. Đối tượng sử dụng1  
   
 ](#anchor-3 "#anchor-3")[2. Chức năng hệ thống1  
   
 ](#anchor-4 "#anchor-4")[2.1. Quản lý mượn sách 1  
   
 ](#anchor-5 "#anchor-5")[2.1.1 Luồng mượn sách qua thẻ sinh viên1  
   
 ](#anchor-6 "#anchor-6")[2.2. Quản lý trả sách 1  
   
 ](#anchor-7 "#anchor-7")[3. Yêu cầu bảo mật2  
   
    
   
    
   
 ](#anchor-8 "#anchor-8")1. Giới thiệu dự án  
1.1.Mục tiêu  
1.2. Đối tượng sử dụng  
2. Chức năng hệ thống  
- Hệ thống cho phép nhân viên thư viện tìm kiếm, tra cứu thông tin sách và khởi tạo phiếu mượn/trả cho sinh viên theo mã thẻ. Đồng thời, hệ thống tự động tính toán số ngày mượn, cảnh báo các trường hợp quá hạn và tự động cập nhật số lượng sách khả dụng trong kho.  
2.1. Quản lý mượn sách   
2.1.1 Luồng mượn sách qua thẻ sinh viên  
**[2.2. Quản lý trả sách**  
 ]()  
3. Yêu cầu bảo mật  
- Toàn bộ mật khẩu tài khoản người dùng và dữ liệu giao dịch mượn trả phải được mã hóa theo chuẩn SHA-256 trước khi lưu trữ vào CSDL. Hệ thống phải phân quyền truy cập chặt chẽ (Role-Based Access Control), chỉ cho phép Nhân viên có thẩm quyền thực hiện thao tác duyệt phiếu mượn và cập nhật kho sách.  
**Phần 2:Chuẩn hoá dữ liệu và Auto Filler**  
Chính xác 3 lỗi:  
- Lỗi 1: Không merge ô ở dòng đầu tiên( Cụ thể: ôA1:E1 đang bị Merge làm tiêu đề lớn)  
- Lỗi 2: Không để trống ô trong vùng dữ liệu ( Cụ thể là Hàng 5)  
- Lỗi 3: Không để tổng/Averge trong vùng dữ liệu ( Cụ thể là Hàng 6)  
   
 Bảng dữ liệu sau khi sửa  
   
    
   
 | | | | | | |  
   
 |-|-|-|-|-|-|  
   
 | Hàng 1 | MSSV | Họ tên  | Mã Sách | Ngày Mượn | Tình trạng |  
   
 | Hàng 2 | SV001 | Nguyễn Văn An | S001 | 01/10/2026 | Đang mượn |  
   
 | Hàng 3 | SV002 | Trần Thị Bình | S002 | 03/10/2026 | Quá hạn |  
   
 | Hang 4 | SV003 | Lê Văn Cường | S003 | 10/10/2026 | Đã trả |  
   
 | Hàng 5 | SV004 | Phạm Thu Hà | S099 | 12/10/2026 | Đang mượn |  
   
    
   
 Phân biệt Record và Field:  
- 1 Record (Bản ghi/ Dòng): Là toàn bộ thông tin đối tượng duy nhất trên 1 hàng nằm trong CSDL:  
   
 Ví dụ: Hàng 2 | SV001 | Nguyễn Văn An | S001 | 01/10/2026 | Đang mượn  
- 1 Field (Trường/Cột): Là thuộc dữ liệu chưá cùng một kiểu thông tin của tất cả đối tượng.  
   
 Ví dụ: Cột MSSV (chưa tất cả mã số sinh viên:SV001 –SV002 –SV003 –SV004)  
**Phần 3: NỐI DỮ LIỆU BẲNG VLOOKUP:**  
1. Bảng Phiếu Mượn sau khi chèn cột “Tên Sách” và áp dụng công thức  
   
    
   
 | | | | | | | |  
   
 |-|-|-|-|-|-|-|  
   
 | Hàng 1 | MSSV | Họ tên  | Mã Sách | Tên Sách(Cột mới) | Ngày Mượn | Tình trạng |  
   
 | Hàng 2 | SV001 | Nguyễn Văn An | S001 | =VLOOKUP(C2, DanhMucSach!$A$2:$C$4, 2, FALSE) → Lập Trình C Căn Bản | 01/10/2026 | Đang mượn |  
   
 | Hàng 3 | SV002 | Trần Thị Bình | S002 | =VLOOKUP(C3, DanhMucSach!$A$2:$C$4, 2, FALSE) → Giải Tích 1 | 03/10/2026 | Quá hạn |  
   
 | Hang 4 | SV003 | Lê Văn Cường | S003 | =VLOOKUP(C4, DanhMucSach!$A$2:$C$4, 2, FALSE) → Kỹ Năng Mềm | 10/10/2026 | Đã trả |  
   
 | Hàng 5 | SV004 | Phạm Thu Hà | S099 | =VLOOKUP(C5, DanhMucSach!$A$2:$C$4, 2, FALSE) → #N/A | 12/10/2026 | Đang mượn |  
2. Giải thích lỗi #N/A ở dòng S099:  
- Lỗi trả về: #N/A (Not Available).  
- Nguyên nhân: Tham số thứ 4 của VLOOKUP là FALSE (yêu cầu dò tìm chính xác 100%). Giá trị S099 tìm kiếm không tồn tại trong cột đầu tiên (Mã Sách) của bảng tra cứu DanhMucSach, do đó Excel thông báo không tìm thấy dữ liệu phù hợp.  
1. Viết lại công thức kết hớp IFERROR:  
   
 Excel  
   
 =IFERROR(VLOOKUP(C5, DanhMucSach!$A$2:$C$4, 2, FALSE), "Không tìm thấy sách")  
   
 —> Kết quả hiển thị tại dòng SV004: Không tìm thấy sách.  
**PHẦN 4 – SLIDE BÁO CÁO TIẾN ĐỘ**  
1. Thiết kế nội dung 1 Slide duy nhất:  
- Assertion Title (Tiêu đề dạng kết luận):  
   
 Hệ thống hiện ghi nhận 4 phiếu mượn, trong đó có 25% quá hạn và 25% mã sách lỗi cần rà soát.  
- Content Area (Số liệu & Kết luận nổi bật):  
  - Con số nổi bật (Big Stat Callout):  
    - 25% (1/4 phiếu) rơi vào trạng thái Quá hạn.  
    - 1 trường hợp mã sách không hợp lệ (S099).  
- Lựa chọn Biểu đồ: Donut Chart (Biểu đồ Vành khăn).  
  - Lý do chọn: Phù hợp nhất để thể hiện cơ cấu/tỷ trọng phần trăm các trạng thái phiếu mượn (2 Đang mượn, 1 Quá hạn, 1 Đã trả) trên tổng thể 100% tài nguyên.  
1. Áp dụng Rule 10/20/30:  
- Tổng thể quy tắc:  
  - 10 slides: Tổng số slide toàn bài presentation.  
  - 20 phút: Tổng thời gian trình bày bài báo cáo.  
  - 30pt: Cỡ chữ tối thiểu cho nội dung text để đảm bảo người nghe dễ quan sát.  
- Áp dụng cho 1 slide báo cáo tiến độ này:  
  - Slide này chiếm 1 slide trong tổng số 10 slides của bài báo cáo.  
  - Thời gian trình bày dự kiến cho slide này: khoảng 1.5 - 2 phút (trên tổng số 20 phút trình bày của toàn nhóm).  
   
    
