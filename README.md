# Local Word AI

Trên ý tưởng Claude for word và với mong muốn bảo mật dữ liệu (hoàn toàn chạy offline), sự hỗ trợ của claude và kilo code, tôi đã hoàn thành Add-in cho Microsoft Word sử dụng LLMs để kiểm tra văn bản, kết nối LM Studio, chạy hoàn toàn offline.

---

## Yêu cầu hệ thống

- Windows 10/11
- Microsoft Word 2019, 2021 LTSC, hoặc Microsoft 365
- .NET Framework 4.8
- VSTO Runtime (Visual Studio 2010 Tools for Office Runtime)
- LM Studio ([lmstudio.ai](https://lmstudio.ai)) với model hỗ trợ tiếng Việt

---

## Cài đặt (Developer)

### 1. Build từ source

```powershell
# Cần Visual Studio 2019/2022 hoặc Build Tools
.\build.ps1 -Configuration Release
```

### 2. Restore NuGet packages thủ công (nếu cần)

```powershell
nuget restore LocalWordAI.sln
```

Các package cần thiết:
- `Newtonsoft.Json` 13.0.3
- `Polly` 7.2.4
- `DiffPlex` 1.7.2
- `FuzzySharp` 2.0.2

### 3. Đăng ký add-in (manual, khi chưa có installer)

Chạy file `register-addin.reg` hoặc thêm vào registry:

```
HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\LocalWordAI
  Description = Local AI Assistant for Microsoft Word - Offline
  FriendlyName = Local Word AI
  LoadBehavior = 3 (DWORD)
  Manifest = file:///C:\path\to\LocalWordAI.vsto|vstolocal
```

### 4. Cài với Inno Setup

```powershell
.\build.ps1 -Configuration Release -Package
```

---

## Cấu hình LM Studio

1. Mở LM Studio.
2. Tải model có hỗ trợ tiếng Việt (ví dụ: Qwen2.5, Gemma2, SeaLLM).
3. Vào **Local Server** → **Start Server**.
4. Mặc định endpoint: `http://localhost:1234`
5. Trong Word, mở **Local AI** tab → **Cài đặt** → nhập URL và test kết nối.

### Model khuyến nghị

| Model | Context | Tiếng Việt | JSON |
|-------|---------|------------|------|
| Qwen2.5-7B-Instruct | 32k | Tốt | Tốt |
| Gemma-2-9b-it | 8k | Khá | Khá |
| SeaLLM-7B | 4k | Tốt | Khá |

Cấu hình LM Studio:
- Temperature: 0.1–0.3
- Max tokens: 2048–4096
- Context: 8192+

---

## Sử dụng

### Ribbon

Tab **Local AI** trên Ribbon Word:
- **AI Sidebar**: Bật/tắt panel bên phải
- **Review**: Kiểm tra vùng chọn
- **Rewrite**: Viết lại vùng chọn
- **Redline**: Tạo redline
- **Comments**: Phân tích comments
- **Track Changes**: Tóm tắt tracked changes
- **Thuật ngữ**: Kiểm tra nhất quán
- **Final Check**: Kiểm tra toàn diện

### Sidebar

1. **Context buttons**: Chọn nguồn dữ liệu (vùng chọn / tài liệu / comments / revisions)
2. **Quick actions**: Các nút tác vụ nhanh
3. **Chat**: Trò chuyện trực tiếp với AI
4. **Gợi ý**: Danh sách vấn đề được phát hiện với các nút:
   - **Track Change**: Áp dụng sửa đổi bằng Track Changes
   - **Comment**: Thêm comment giải thích
   - **Highlight**: Tô màu đoạn cần xem xét
   - **Bỏ qua**: Bỏ qua gợi ý này
5. **Skills**: Chạy workflow có sẵn

### Skills

Skills là các workflow tái sử dụng được lưu trong `%APPDATA%\LocalWordAI\Skills\`.

Skills mẫu:
- `nda-review.json` — Rà soát NDA
- `sale-contract-review.json` — Rà soát hợp đồng mua bán
- `official-letter-review.json` — Rà soát công văn
- `final-consistency-check.json` — Final check trước phát hành

Để tạo skill mới, copy một file JSON và chỉnh sửa.

---

## Kiến trúc

```
LocalWordAI/
├── ThisAddIn.cs          # Entry point VSTO
├── Ribbon.cs             # Custom Ribbon XML
├── AI/                   # LM Studio client, Prompt builder
├── WordIntegration/      # Word Interop layer
├── Rules/                # Rule-based checkers
├── Skills/               # Skill manager
├── UI/                   # WinForms Task Pane
├── Models/               # Data models
└── Persistence/          # Settings, cache
```

---

## Bảo mật & Privacy

- **Không gửi dữ liệu ra ngoài** — mọi xử lý trong máy.
- URL chỉ cho phép `localhost` / `127.0.0.1`.
- Log mặc định chỉ ghi lỗi kỹ thuật, không ghi nội dung tài liệu.
- Không có telemetry, analytics, hoặc auto-update online.

---

## Troubleshooting

| Vấn đề | Giải pháp |
|--------|-----------|
| Add-in không hiện trong Word | Kiểm tra Trust Center → Add-ins; chạy lại installer |
| "Mất kết nối" | Kiểm tra LM Studio đang chạy, URL đúng trong Cài đặt |
| JSON lỗi / không parse được | Giảm temperature LM Studio về 0.1; thử lại |
| Không tìm thấy đoạn văn khi áp dụng | Đoạn văn có thể đã thay đổi; refresh và thử lại |
| Word bị chậm | Tắt realtime check; dùng manual review |

---

## License

Dự án nội bộ — không phân phối thương mại.
