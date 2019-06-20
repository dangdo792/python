# View existing remotes:
git remote -v

$ git branch -a


Kiểm tra trạng thái thay đổi
git status

Đưa những file vào danh sách trước khi commit (bỏ qua nếu không có file mới được tạo)
git add -A  (-A Tất cả files)

Commit những thay đổi trước khi push lên server
git commit -a -m “Thông tin ngắn gọn về thay đổi này” (-a Tất cả thay đổi)

Tải lên server nhánh master
git push -u origin master

TẠO NHÁNH (BRANCH) RIÊNG
Xem toàn bộ các nhánh đang có
git branch -a

Tạo nhánh mới
git branch <tên mới>

Chuyển nhánh
git checkout <tên nhánh>

Xóa nhánh
git branch -d <Tên branch>

NHẬP NHÁNH CON VÀO NHÁNH HIỆN TẠI
git merge <Tên nhánh>

LIỆT KÊ MỘT SỐ LỆNH HAY DÙNG
touch “tên file”  (Tạo 1 file)
echo “nội dung” > “tên file” (viết nội dung vào trong file)
git remote add origin git@github.com:<tên repo>  (Tạo kết nối tới Github server)
git pull git@github.com:<tên repo> <tên branch>  (Cập nhật dữ liệu từ server)
git status (Xem trạng thái hiện tại)
git add -A (Cộng tất cả nhưng file đã thay đổi vào danh sách để đưa lên server)
git commit -a -m “Thông tin note”  (Đưa tất cả những file vào danh sách để chuẩn bị push)
git push origin master  (Gửi toàn bộ file đã commit lên server vào nhánh master)