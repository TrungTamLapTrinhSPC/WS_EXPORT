# WS_EXPORT

**WS_EXPORT** là một repository chứa một giải pháp C# (Solution `MOU_CSV`) phục vụ mục đích xuất/convert dữ liệu sang CSV/Excel. README này là một bản giới thiệu tổng quan và hướng dẫn nhanh để người khác có thể hiểu, build và chạy project trên GitHub.

> ⚠️ Ghi chú: tôi đã đọc cấu trúc dự án (có file/ thư mục chính: `MOU_CSV`, `MOU_CSV.sln`, và một vài thư mục cấu hình IDE). Vì repo hiện chưa có mô tả chi tiết, README này viết theo cách tổng quát — nếu bạn muốn tôi tùy chỉnh README dựa trên file cụ thể trong repo (ví dụ: `Program.cs`, `app.config`, hoặc mô-đun export), gửi tên file hoặc nội dung và tôi sẽ cập nhật.

---

## Mô tả

Project `MOU_CSV` là một ứng dụng C# (Windows/.NET) dùng để chuyển đổi dữ liệu (từ database, từ file nguồn hoặc từ dịch vụ nội bộ) thành file CSV/Excel phục vụ báo cáo hoặc trao đổi dữ liệu.

Các tính năng chính (ước lượng):

* Đọc dữ liệu từ nguồn (DB / file / webservice).
* Map/format dữ liệu theo cấu trúc MOU / bảng mẫu.
* Xuất ra file CSV (hoặc Excel) tương thích với hệ thống tiêu thụ.
* Cấu hình tham số (đường dẫn, connection string, template) qua file cấu hình.

---

## Yêu cầu

* Windows 10/11 hoặc máy phát triển có .NET Framework / .NET SDK phù hợp.
* Visual Studio 2019 / 2022 (hoặc `dotnet` CLI nếu project là .NET Core/.NET 5+).

> **Gợi ý**: mở file `MOU_CSV.sln` trong Visual Studio để biết chính xác target framework. Nếu repo dùng .NET Core / .NET 5+, bạn có thể dùng `dotnet build`.

---

## Cách build và chạy

### 1) Mở bằng Visual Studio

1. Clone repository:

```bash
git clone https://github.com/TrungTamLapTrinhSPC/WS_EXPORT.git
cd WS_EXPORT
```

2. Mở `MOU_CSV.sln` bằng Visual Studio.
3. Chọn configuration `Release` hoặc `Debug` và `Build Solution`.
4. Chạy project chính (Set as Startup Project nếu có nhiều project).

### 2) Dùng dotnet CLI (nếu là .NET Core / .NET 5+)

```bash
cd MOU_CSV
dotnet restore
dotnet build
dotnet run --project ./Path/To/Project.csproj
```

> Nếu project target .NET Framework (ví dụ `net48`) thì phải dùng Visual Studio để build.

---

## Cấu hình (ví dụ)

Project thường có phần cấu hình để đặt: connection string, thư mục đầu vào/đầu ra, tham số map cột. Có thể nằm trong `app.config`, `web.config`, hoặc file JSON `appsettings.json`.

Ví dụ `appsettings.json` (mẫu):

```json
{
  "InputFolder": "./input",
  "OutputFolder": "./output",
  "ConnectionString": "Server=...;Database=...;User Id=...;Password=...;",
  "CsvOptions": {
    "Delimiter": ",",
    "Encoding": "utf-8",
    "IncludeHeader": true
  }
}
```

---

## Cách sử dụng (ví dụ)

1. Chuẩn bị nguồn dữ liệu: file đầu vào hoặc đảm bảo DB/Service sẵn sàng.
2. Cập nhật cấu hình (đường dẫn, connection string, template mapping).
3. Chạy ứng dụng. File CSV sẽ được tạo trong thư mục `output` (hoặc theo cấu hình).

 
## Kiểm thử

* Nếu repo có unit tests, mở project test và chạy qua Test Explorer của Visual Studio hoặc `dotnet test`.
* Với các công cụ export, test bằng dataset mẫu để kiểm tra định dạng CSV, encoding, và ordering của cột.

---

## Triển khai

* Dùng bản build `.exe` (Windows) để chạy định kỳ (Task Scheduler) hoặc đóng gói thành Windows Service nếu cần chạy lâu dài.
* Với `.NET Core`, có thể publish dưới dạng single-file hoặc containerize (Docker) nếu cần.

--- 

## Liên hệ

Nếu bạn muốn mình cập nhật README cho chính xác với nội dung các file trong repo (ví dụ: nội dung `Program.cs`, cấu trúc `Mappers`, `Exporters`), gửi tên những file cụ thể hoặc cho phép mình mở những file đó, mình sẽ chỉnh README để chính xác 100% theo code.

Xin chào,
*SPC Assistant*
