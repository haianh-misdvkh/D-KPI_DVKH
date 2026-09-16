# Đưa Dashboard lên Internet miễn phí — Streamlit Community Cloud

Đây là cách **free + đơn giản nhất** để mọi người truy cập dashboard qua 1 link internet, không cần bạn bật máy tính.

## ⚠️ Lưu ý quan trọng trước khi làm

1. **Dùng Repository PRIVATE trên GitHub, không dùng Public.**
   App của bạn có đăng nhập bảo vệ, nhưng nếu repo GitHub để Public thì bất kỳ ai cũng tải thẳng được file `DATA_KPI.xlsb` từ GitHub — bỏ qua luôn màn hình đăng nhập. Repo Private vẫn deploy free bình thường trên Streamlit Cloud.
2. **File `user_accounts.json` sẽ mất khi app khởi động lại.**
   Streamlit Cloud không lưu file bạn ghi ra khi chạy (ổ đĩa tạm thời). Nghĩa là nếu Admin đổi mật khẩu / thêm user trên app đang public, thay đổi đó **có thể mất** khi app tự khởi động lại (ngủ do không ai dùng, hoặc bạn deploy lại). Cách khắc phục đơn giản: sau khi chỉnh sửa tài khoản, bấm nút **"Tải file user_accounts.json (backup)"** trong trang Admin, rồi thêm file đó vào repo GitHub (đè lên file cũ) để lần deploy sau giữ được thay đổi.

## Bước 1 — Tạo tài khoản GitHub (free)

Vào https://github.com/signup nếu chưa có tài khoản.

## Bước 2 — Tạo Repository mới

1. Vào https://github.com/new
2. Đặt tên repo, ví dụ `kpi-dashboard`
3. Chọn **Private**
4. Bấm **Create repository**

## Bước 3 — Đưa code lên GitHub

Giải nén file zip tôi gửi, **thêm file dữ liệu thật của bạn vào đó** với tên `DATA_KPI.xlsb`, rồi mở Terminal/Command Prompt tại thư mục đó và chạy:

```bash
git init
git add .
git commit -m "Dashboard KPI ban dau"
git branch -M main
git remote add origin https://github.com/<TEN_GITHUB_CUA_BAN>/kpi-dashboard.git
git push -u origin main
```

(Thay `<TEN_GITHUB_CUA_BAN>` bằng username GitHub của bạn. Lần đầu push, GitHub sẽ hỏi đăng nhập — làm theo hướng dẫn trên màn hình.)

## Bước 4 — Deploy lên Streamlit Community Cloud (free)

1. Vào https://share.streamlit.io
2. Đăng nhập bằng tài khoản GitHub (nút "Sign in with GitHub")
3. Bấm **"New app"** (hoặc "Create app")
4. Chọn:
   - Repository: `<TEN_GITHUB_CUA_BAN>/kpi-dashboard`
   - Branch: `main`
   - Main file path: `kpi_streamlit_app.py`
5. Bấm **Deploy**

Đợi 1-2 phút, app sẽ có link công khai dạng:
```
https://kpi-dashboard-xxxxx.streamlit.app
```

Gửi link này cho toàn bộ cán bộ — ai cũng vào được, mỗi người vẫn đăng nhập bằng user/pass riêng như cũ.

## Cập nhật dữ liệu sau này

Muốn thay `DATA_KPI.xlsb` mới: vào repo GitHub → upload file mới đè lên file cũ (hoặc `git add` + `git commit` + `git push` lại từ máy bạn) → app trên Streamlit Cloud **tự động deploy lại** trong 1-2 phút, không cần làm gì thêm ở bước Streamlit Cloud.

## Nếu app "ngủ"

Với tài khoản free, app sẽ tạm ngưng nếu không ai truy cập trong vài ngày. Chỉ cần ai đó mở lại link, app tự khởi động lại sau khoảng 30-60 giây — không mất dữ liệu Excel, chỉ có `user_accounts.json` là cần lưu ý như phần Lưu ý ở trên.

made by Dungpv-QC_DVKH
