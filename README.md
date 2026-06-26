# LocalDocAI — Offline AI Document Reviewer for Microsoft Word

[🇻🇳 Tiếng Việt](#vn) | [🇬🇧 English](#en) | [🇫🇷 Français](#fr)

---

<a name="vn"></a>

# 🇻🇳 Tiếng Việt

## Giới thiệu

Dựa trên ý tưởng Claude for Word và mong muốn bảo mật dữ liệu hoàn toàn offline, tôi đã phát triển Add-in Microsoft Word sử dụng LLMs cục bộ qua LM Studio để kiểm tra, viết lại và chỉnh sửa tài liệu — tất cả chạy hoàn toàn trên máy của bạn.

<!--more-->

## Yêu cầu hệ thống

- Windows 10 / 11
- Microsoft Word 2019, 2021 LTSC, hoặc Microsoft 365
- .NET Framework 4.8
- VSTO Runtime (Visual Studio 2010 Tools for Office Runtime)
- LM Studio ([lmstudio.ai](https://lmstudio.ai)) với model hỗ trợ tiếng Việt

## Cài đặt (Developer)

### 1. Build từ source

```powershell
.\build.ps1 -Configuration Release
```

### 2. Restore NuGet packages (nếu cần)

```powershell
nuget restore LocalDocAI.sln
```

Package cần thiết:

- `Newtonsoft.Json` 13.0.3
- `Polly` 7.2.4
- `DiffPlex` 1.7.2
- `FuzzySharp` 2.0.2

### 3. Đăng ký add-in (manual)

Chạy `register-addin.reg` hoặc thêm vào registry:

```
HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\LocalDocAI
  Description = Local AI Assistant for Word — Offline
  FriendlyName = Local Word AI
  LoadBehavior = 3 (DWORD)
  Manifest = file:///C:\path\to\LocalDocAI.vsto|vstolocal
```

### 4. Cài với Inno Setup

```powershell
.\build.ps1 -Configuration Release -Package
```

## Cấu hình LM Studio

1. Mở LM Studio, tải model có hỗ trợ tiếng Việt (ví dụ: Qwen2.5, Gemma2, SeaLLM).
2. Vào **Local Server** → **Start Server**.
3. Mặc định endpoint: `http://localhost:1234`
4. Trong Word: **Local AI** tab → **Cài đặt** → nhập URL và test kết nối.

### Model khuyến nghị

| Model | Context | Tiếng Việt | JSON |
|-------|---------|------------|------|
| Qwen2.5-7B-Instruct | 32k | Tốt | Tốt |
| Gemma-2-9b-it | 8k | Khá | Khá |
| SeaLLM-7B | 4k | Tốt | Khá |

- Temperature: 0.1 – 0.3
- Max tokens: 2048 – 4096
- Context: 8192+

## Sử dụng

### Ribbon

Tab **Local AI** trên Ribbon Word:

| Nút | Chức năng |
|-----|-----------|
| AI Sidebar | Bật/tắt panel bên phải |
| Review | Kiểm tra vùng chọn |
| Rewrite | Viết lại vùng chọn |
| Redline | Tạo tracked changes |
| Comments | Phân tích comments |
| Track Changes | Tóm tắt tracked changes |
| Thuật ngữ | Kiểm tra nhất quán |
| Final Check | Kiểm tra toàn diện trước phát hành |

### Sidebar

1. **Context buttons** — chọn nguồn (vùng chọn / tài liệu / comments / revisions)
2. **Quick actions** — các nút tác vụ nhanh
3. **Chat** — trò chuyện trực tiếp với AI
4. **Gợi ý** — danh sách vấn đề phát hiện:
   - **Track Change**: áp dụng sửa đổi bằng Track Changes
   - **Comment**: thêm comment giải thích
   - **Highlight**: tô màu đoạn cần xem xét
   - **Bỏ qua**: bỏ qua gợi ý này
5. **Skills** — chạy workflow có sẵn

### Skills

Skills là workflow tái sử dụng, lưu trong `%APPDATA%\LocalDocAI\Skills\`.

Skill mẫu:

- `nda-review.json` — Rà soát NDA
- `sale-contract-review.json` — Rà soát hợp đồng mua bán
- `official-letter-review.json` — Rà soát công văn
- `final-consistency-check.json` — Final check trước phát hành

## Kiến trúc

```
LocalDocAI/
├── ThisAddIn.cs           # Entry point VSTO
├── Ribbon.cs              # Custom Ribbon XML
├── AI/                    # LM Studio client, Prompt builder
├── WordIntegration/       # Word Interop layer
├── Rules/                 # Rule-based checkers
├── Skills/                # Skill manager
├── UI/                    # WinForms Task Pane
├── Models/                # Data models
└── Persistence/           # Settings, cache
```

## Bảo mật & Quyền riêng tư

- **Không gửi dữ liệu ra ngoài** — mọi xử lý trong máy.
- URL chỉ cho phép `localhost` / `127.0.0.1`.
- Log chỉ ghi lỗi kỹ thuật, không ghi nội dung tài liệu.
- Không có telemetry, analytics, hoặc auto-update online.

## Xử lý sự cố

| Vấn đề | Giải pháp |
|--------|-----------|
| Add-in không hiện trong Word | Kiểm tra Trust Center → Add-ins; chạy lại installer |
| "Mất kết nối" | Kiểm tra LM Studio đang chạy, URL đúng trong Cài đặt |
| JSON lỗi / không parse được | Giảm temperature LM Studio về 0.1; thử lại |
| Không tìm thấy đoạn văn khi áp dụng | Đoạn văn có thể đã thay đổi; refresh và thử lại |
| Word bị chậm | Tắt realtime check; dùng manual review |

## License

Dự án nội bộ — không phân phối thương mại.

---

<a name="en"></a>

# 🇬🇧 English

## Overview

Add-in for Microsoft Word that performs AI-powered document review, rewriting, and editing **entirely offline** using a local LLM via [LM Studio](https://lmstudio.ai). Designed with Vietnamese-language document workflows in mind, but works with any language supported by your local model.

## System Requirements

- Windows 10 / 11
- Microsoft Word 2019, 2021 LTSC, or Microsoft 365
- .NET Framework 4.8
- VSTO Runtime (Visual Studio 2010 Tools for Office Runtime)
- LM Studio ([lmstudio.ai](https://lmstudio.ai)) with a model that supports your target language

## Installation (Developer)

### 1. Build from source

```powershell
.\build.ps1 -Configuration Release
```

### 2. Restore NuGet packages (if needed)

```powershell
nuget restore LocalDocAI.sln
```

Required packages:

- `Newtonsoft.Json` 13.0.3
- `Polly` 7.2.4
- `DiffPlex` 1.7.2
- `FuzzySharp` 2.0.2

### 3. Register the add-in (manual)

Run `register-addin.reg` or add the following registry entries:

```
HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\LocalDocAI
  Description = Local AI Assistant for Word — Offline
  FriendlyName = Local Word AI
  LoadBehavior = 3 (DWORD)
  Manifest = file:///C:\path\to\LocalDocAI.vsto|vstolocal
```

### 4. Install with Inno Setup

```powershell
.\build.ps1 -Configuration Release -Package
```

## LM Studio Configuration

1. Open LM Studio.
2. Download a model with strong multilingual support (e.g., Qwen2.5, Gemma2, SeaLLM).
3. Go to **Local Server** → **Start Server**.
4. Default endpoint: `http://localhost:1234`
5. In Word: **Local AI** tab → **Settings** → enter URL and test connection.

### Recommended models

| Model | Context | Vietnamese | JSON |
|-------|---------|------------|------|
| Qwen2.5-7B-Instruct | 32k | Good | Good |
| Gemma-2-9b-it | 8k | Fair | Fair |
| SeaLLM-7B | 4k | Good | Fair |

- Temperature: 0.1 – 0.3
- Max tokens: 2048 – 4096
- Context window: 8192+

## Usage

### Ribbon

The **Local AI** tab in the Word Ribbon provides:

| Button | Action |
|--------|--------|
| AI Sidebar | Toggle the right-side panel |
| Review | Check selected text |
| Rewrite | Rewrite selected text |
| Redline | Create tracked changes |
| Comments | Analyze document comments |
| Track Changes | Summarize tracked revisions |
| Terms | Check terminology consistency |
| Final Check | Comprehensive pre-release check |

### Sidebar

1. **Context buttons** — choose data source (selection / full document / comments / revisions)
2. **Quick actions** — common tasks
3. **Chat** — direct conversation with the AI
4. **Suggestions** — detected issues with actions:
   - **Track Change**: apply via Track Changes
   - **Comment**: add an explanatory comment
   - **Highlight**: highlight the section
   - **Skip**: dismiss this suggestion
5. **Skills** — run built-in workflows

### Skills

Skills are reusable workflows stored in `%APPDATA%\LocalDocAI\Skills\`.

Sample skills:

- `nda-review.json` — NDA review
- `sale-contract-review.json` — Sales contract review
- `official-letter-review.json` — Official letter review
- `final-consistency-check.json` — Final consistency check before release

## Architecture

```
LocalDocAI/
├── ThisAddIn.cs           # VSTO entry point
├── Ribbon.cs              # Custom Ribbon XML
├── AI/                    # LM Studio client, Prompt builder
├── WordIntegration/       # Word Interop layer
├── Rules/                 # Rule-based checkers
├── Skills/                # Skill manager
├── UI/                    # WinForms Task Pane
├── Models/                # Data models
└── Persistence/           # Settings, cache
```

## Security & Privacy

- **No data leaves your machine** — all processing is local.
- Only `localhost` / `127.0.0.1` URLs are permitted.
- Default logging covers technical errors only; document content is not logged.
- No telemetry, analytics, or online auto-updates.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Add-in missing from Word | Check Trust Center → Add-ins; re-run installer |
| "Connection lost" | Verify LM Studio is running and URL is correct in Settings |
| JSON errors / parsing failures | Lower LM Studio temperature to 0.1; retry |
| Paragraph not found on apply | Paragraph may have changed; refresh and retry |
| Word is slow | Disable realtime check; use manual review |

## License

Internal project — not for commercial distribution.

---

<a name="fr"></a>

# 🇫🇷 Français

## Aperçu

Complément (add-in) pour Microsoft Word qui effectue la révision, la réécriture et l'édition de documents **entièrement hors ligne** grâce à un LLM local via [LM Studio](https://lmstudio.ai). Conçu principalement pour les workflows de documents en vietnamien, mais fonctionne avec toute langue prise en charge par votre modèle local.

## Configuration requise

- Windows 10 / 11
- Microsoft Word 2019, 2021 LTSC ou Microsoft 365
- .NET Framework 4.8
- VSTO Runtime (Visual Studio 2010 Tools for Office Runtime)
- LM Studio ([lmstudio.ai](https://lmstudio.ai)) avec un modèle prenant en charge votre langue cible

## Installation (Développeur)

### 1. Construire depuis les sources

```powershell
.\build.ps1 -Configuration Release
```

### 2. Restaurer les packages NuGet (si nécessaire)

```powershell
nuget restore LocalDocAI.sln
```

Packages requis :

- `Newtonsoft.Json` 13.0.3
- `Polly` 7.2.4
- `DiffPlex` 1.7.2
- `FuzzySharp` 2.0.2

### 3. Enregistrer le complément (manuellement)

Exécuter `register-addin.reg` ou ajouter les entrées de registre suivantes :

```
HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\LocalDocAI
  Description = Local AI Assistant for Word — Offline
  FriendlyName = Local Word AI
  LoadBehavior = 3 (DWORD)
  Manifest = file:///C:\path\to\LocalDocAI.vsto|vstolocal
```

### 4. Installer avec Inno Setup

```powershell
.\build.ps1 -Configuration Release -Package
```

## Configuration de LM Studio

1. Ouvrez LM Studio.
2. Téléchargez un modèle avec une bonne prise en charge multilingue (ex. : Qwen2.5, Gemma2, SeaLLM).
3. Allez dans **Local Server** → **Start Server**.
4. Point d'accès par défaut : `http://localhost:1234`
5. Dans Word : onglet **Local AI** → **Paramètres** → saisissez l'URL et testez la connexion.

### Modèles recommandés

| Modèle | Contexte | Vietnamien | JSON |
|--------|----------|------------|------|
| Qwen2.5-7B-Instruct | 32k | Bon | Bon |
| Gemma-2-9b-it | 8k | Correct | Correct |
| SeaLLM-7B | 4k | Bon | Correct |

- Température : 0.1 – 0.3
- Tokens max : 2048 – 4096
- Fenêtre de contexte : 8192+

## Utilisation

### Ruban

L'onglet **Local AI** dans le ruban Word propose :

| Bouton | Action |
|--------|--------|
| AI Sidebar | Afficher/masquer le panneau latéral |
| Review | Vérifier la sélection |
| Rewrite | Réécrire la sélection |
| Redline | Créer des modifications suivies |
| Comments | Analyser les commentaires |
| Track Changes | Résumer les révisions suivies |
| Terms | Vérifier la cohérence terminologique |
| Final Check | Vérification complète avant publication |

### Panneau latéral

1. **Context buttons** — choisir la source (sélection / document entier / commentaires / révisions)
2. **Quick actions** — tâches courantes
3. **Chat** — conversation directe avec l'IA
4. **Suggestions** — problèmes détectés avec actions :
   - **Track Change** : appliquer via le suivi des modifications
   - **Comment** : ajouter un commentaire explicatif
   - **Highlight** : surligner la section
   - **Skip** : ignorer cette suggestion
5. **Skills** — exécuter des workflows prédéfinis

### Skills

Les Skills sont des workflows réutilisables, stockés dans `%APPDATA%\LocalDocAI\Skills\`.

Skills d'exemple :

- `nda-review.json` — Révision NDA
- `sale-contract-review.json` — Révision de contrat de vente
- `official-letter-review.json` — Révision de lettre officielle
- `final-consistency-check.json` — Vérification finale avant publication

## Architecture

```
LocalDocAI/
├── ThisAddIn.cs           # Point d'entrée VSTO
├── Ribbon.cs              # XML du ruban personnalisé
├── AI/                    # Client LM Studio, générateur de prompts
├── WordIntegration/       # Couche Interop Word
├── Rules/                 # Vérifications par règles
├── Skills/                # Gestionnaire de skills
├── UI/                    # Volet de tâches WinForms
├── Models/                # Modèles de données
└── Persistence/           # Paramètres, cache
```

## Sécurité & Vie privée

- **Aucune donnée ne quitte votre machine** — tout le traitement est local.
- Seules les URL `localhost` / `127.0.0.1` sont autorisées.
- La journalisation couvre uniquement les erreurs techniques ; le contenu des documents n'est pas enregistré.
- Aucune télémétrie, analytique ou mise à jour automatique en ligne.

## Dépannage

| Problème | Solution |
|----------|----------|
| Le complément n'apparaît pas dans Word | Vérifier Centre de confiance → Compléments ; réexécuter l'installateur |
| "Connexion perdue" | Vérifier que LM Studio est en cours d'exécution et que l'URL est correcte dans Paramètres |
| Erreurs JSON / échec d'analyse | Baisser la température de LM Studio à 0.1 ; réessayer |
| Paragraphe introuvable lors de l'application | Le paragraphe a pu être modifié ; actualiser et réessayer |
| Word est lent | Désactiver la vérification en temps réel ; utiliser la révision manuelle |

## Licence

Projet interne — pas de distribution commerciale.
