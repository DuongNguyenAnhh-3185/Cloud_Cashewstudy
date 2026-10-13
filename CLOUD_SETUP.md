# Thiết lập Cloud cho CashewDoc

Ứng dụng dùng Firebase Authentication làm nguồn danh tính, Firebase Storage giữ
tệp vật lý và Supabase Postgres giữ `users`, `documents`, `permissions`. RLS trên
Supabase kiểm tra trực tiếp Firebase ID token. Firestore chỉ chứa bản sao ACL tối
thiểu để Firebase Storage Rules có thể kiểm tra quyền đọc tệp; không lưu metadata
nghiệp vụ trong Firestore.

## 1. Firebase

Project đã được cấu hình trong mã nguồn là `device-streaming-33cac12e`.

1. Trong Firebase Console, bật **Authentication > Sign-in method > Google**.
2. Thêm SHA-1 và SHA-256 của debug/release keystore cho ứng dụng Android.
3. Tạo Cloud Firestore ở Production mode và bật Firebase Storage.
4. Triển khai Rules:

   ```bash
   firebase login
   firebase use device-streaming-33cac12e
   firebase deploy --only firestore:rules,storage
   ```

Không dùng Firebase download URL công khai. Ứng dụng đọc tệp qua Firebase SDK,
để `storage.rules` luôn kiểm tra người dùng hiện tại.

## 2. Firebase JWT và Postgres role

Firebase ID token mặc định không có custom claim `role`. Supabase ánh xạ request
này vào Postgres role `anon`; migration cấp quyền tối thiểu cho role đó và RLS
restrictive tiếp tục bắt buộc `iss`, `aud` và `sub` thuộc đúng Firebase project.
Request chỉ có publishable key nhưng không có Firebase JWT sẽ bị từ chối.

Vì vậy không cần triển khai thư mục `functions` hay bật Identity Platform chỉ để
thêm claim `role=authenticated`.

## 3. Supabase

1. Tạo một Supabase project.
2. Vào **Authentication > Third-Party Auth > Firebase**, thêm Firebase Project ID
   `device-streaming-33cac12e`.
3. Chạy file migration
   `supabase/migrations/202610130001_cloud_documents.sql` trong SQL Editor, hoặc:

   ```bash
   supabase link --project-ref <project-ref>
   supabase db push
   ```

Migration tạo ba bảng, khóa ngoại, index, RPC chia sẻ/thu hồi và toàn bộ RLS.
Publishable key có thể đặt ở client; tuyệt đối không đưa `service_role` key vào app.

## 4. Chạy ứng dụng

```bash
flutter pub get
flutter run \
  --dart-define=SUPABASE_URL=https://<project-ref>.supabase.co \
  --dart-define=SUPABASE_PUBLISHABLE_KEY=<publishable-key>
```

Trên PowerShell:

```powershell
flutter run `
  --dart-define=SUPABASE_URL=https://<project-ref>.supabase.co `
  --dart-define=SUPABASE_PUBLISHABLE_KEY=<publishable-key>
```

## 5. Kiểm thử ba tài khoản

Bộ kiểm thử RLS tự động nằm tại
`supabase/tests/database/rls_3_users_test.sql`:

```bash
supabase start
supabase db reset
supabase test db
```

Kiểm thử thực tế sau khi deploy:

| Tài khoản | Thao tác | Kết quả mong đợi |
|---|---|---|
| Alice | Tải tài liệu và chia sẻ `editor` cho Bob | Alice thấy quyền của Bob |
| Bob | Tìm, tải và sửa tiêu đề tài liệu Alice | Thành công; không thể chia sẻ/xóa |
| Carol | Mở tài liệu trước khi được chia sẻ | Không thấy metadata và không tải được tệp |
| Alice | Chia sẻ `viewer` cho Carol | Carol thấy và tải được nhưng không sửa được |
| Alice | Thu hồi quyền Carol | Carol mất quyền metadata lẫn tệp ngay lập tức |

Chạy thêm kiểm thử Flutter:

```bash
flutter test
flutter analyze
```

## Lưu ý nền tảng

Google Sign-In chính thức hỗ trợ Android, iOS, macOS và Web. Với Windows, chạy
bản Web (`flutter run -d chrome`) hoặc bổ sung OAuth desktop broker riêng; không
nhúng OAuth client secret vào ứng dụng Flutter.
