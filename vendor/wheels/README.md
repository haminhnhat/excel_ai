# Offline Package Cache

Thư mục này dùng để chứa các file Python package `.whl` hoặc source package `.tar.gz` đã tải sẵn.

Mục tiêu: user nội bộ có thể chạy `Setup_First_Time.bat` mà không cần tải package từ internet.

## Dành cho admin/IT

Trên một máy có internet, chạy:

```text
Build_Offline_Package.bat
```

Script sẽ tải các package trong `requirements.txt` vào:

```text
vendor/wheels/
```

Sau đó copy toàn bộ thư mục project sang máy user.

## Dành cho user

User chỉ cần chạy:

```text
Setup_First_Time.bat
```

Nếu thư mục `vendor/wheels/` có package, setup sẽ tự cài offline từ thư mục này.

## Lưu ý

Wheel cache nên được build bằng cùng version Python với máy user, ví dụ Python 3.11 hoặc 3.12 trên Windows 64-bit.
