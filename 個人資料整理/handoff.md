# 工作交接 — 2026-05-10

## 已完成
- 將 `個人基本資料/D_歷年考核` 改名為 `D_考核與敘薪`，並同步更新索引
- 建立個人資料整理系統 v2 完整架構（_inbox/、_待確認/、分類規則.md、PROFILE.md、專案紀錄/、CLAUDE.md v2 觸發指令、INDEX.md）
- 修復 git index 損壞物件，設定 gc.auto=0 停用 auto-gc
- 建立專案紀錄：啟發潛能之星申請（115年度國語）、115典範社群
- 刪除 `D:\test ch\114啟發潛能之星申請\`（使用者確認可刪）
- 整合 Obsidian：
  - PROFILE.md 第10節改為指向 Obsidian Dashboard，不再重複維護專案清單
  - 在 Obsidian `artifacts/個人資料.md` 建立指標筆記
- 建立三個 GitHub 私人 repo 備份：
  - `lesson-search`：D:\test ch\（含個人資料整理、字音形系統等）
  - `beike-ai`：D:\備課ai\ 程式碼（媒體/文件交 Google Drive）
  - `obsidian-vault`：D:\test ch\obsidian\

## 確定的系統分工

| 系統 | 負責什麼 |
|------|---------|
| Obsidian Dashboard | 所有專案的完整清單（技術細節、部署連結、狀態） |
| PROFILE.md | 申請用自介素材、有成效數字的亮點摘要 |
| 專案紀錄/_專案摘要.md | 申請用完整描述（成效、教學特色、創新亮點） |
| _inbox/ | 所有新文件的唯一入口 |

## 備份架構

| 資料夾 | Google Drive | GitHub |
|--------|-------------|--------|
| D:\test ch\ | ✅ | ✅ lesson-search |
| D:\備課ai\ 程式碼 | ✅ | ✅ beike-ai |
| D:\test ch\obsidian\ | ✅ | ✅ obsidian-vault |

**換電腦回復流程：**
1. 安裝 Git、Python、Node.js、clasp
2. 等 Google Drive 同步完成
3. git clone 三個 repo
4. 重跑 Google OAuth 授權（token.json）

## 使用情境對照

- **開發新專案** → 在 備課ai/ 建 CLAUDE.md，Obsidian Dashboard 自動出現
- **申請獎項** → 說「寫自介」，需要專案清單看 Obsidian Dashboard
- **有新文件** → 丟 _inbox/，說「入庫」
- **專案有新成效** → 說「更新專案」
- **寫懶人包/腳本** → 在 Obsidian 用 /write-partner，引用 artifacts/個人資料.md

## 未完成／待確認
- PROFILE.md 第3–7節（教學成果、輔導紀錄、行政服務、進修研習、得獎榮譽）均標示「待補」
- PROFILE.md 第8節「等第摘要」需翻閱實際考核通知書後填入
- 素材庫尚未產出任何自介版本
- `G_獎懲紀錄` 中「第04頁」與「第04頁續」需確認是否重拍

## 下一步
1. 說「寫自介 講師簡介50字」測試系統能否正常產出
2. 開始入庫：把進修研習時數證明、教學成果文件丟進 `_inbox/`，說「入庫」
3. 每次工作結束 git push 保持備份最新

## 注意事項
- git repo 根目錄是 `D:\test ch\`，commit 時加 `-c gc.auto=0`
- D:\備課ai\.gitignore 已排除大型媒體檔（pdf、mp4、docx、xlsx 等）和 AI擺渡人/、claude_backup/、youtube影片整理/、研習講義/
- obsidian-vault repo 排除了 備課ai Junction（它是指向 D:\備課ai 的捷徑，由 beike-ai repo 另行備份）
- PROFILE.md 第10節專案清單已改為指向 Obsidian，不再自行維護完整清單
