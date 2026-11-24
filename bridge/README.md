# ACT / PF bridge worker

這是獨立的同步工作程序。臨床資料只透過執行中的 Overlay HTTP API 存取；它不直接開啟 DuckDB、DBF 或 HIS 檔案，也不需要重新啟動 Overlay、Tray 或 TK Panel。

截至 2026-11-24 15:23（Asia/Taipei），**Google 擁有者初始化與手機前端發布已完成；worker 私人設定、安裝及啟動尚待完成**。15:22 已請使用者明確同意將同步金鑰保存於這台電腦，回覆前不保存或啟動。請依 [部署與回復手冊](../repair-delivery/20261124-history/DEPLOY_AND_ROLLBACK_ZH.md) 完成本機設定。前端已發布或程式檔已存在，不代表 bridge 已執行或正式雙向同步已驗收。

## 資料方向

| 方向 | 路徑 | 成功依據 |
| --- | --- | --- |
| 桌面輸入 → 手機 | Overlay API 的有效 PF 列 → `pf_sync` → `PF Records` → 手機 `nurse_history` | 雲端回執的事件、病歷、日期、數值及 Overlay ID 一致 |
| 手機 → Overlay | `pf_record` → Google `PF Records` pending → bridge 病歷核對 → Overlay PF 佇列 → 精確讀回 → `pf_ack` | 精確讀回已存在的臨床列後，才可把 Google 狀態改成 synced |

Overlay 接受排入佇列不代表資料已寫入。worker 以穩定 `entered_by=act_pf_sync:<eventId>` 查回病歷、日期、值、來源及實際列 ID，避免把排隊回應或逾時誤認為完成。病歷不存在時送出 `rejected_patient`，不創造臨床病患。

## 本輪範圍

- 病歷號標準化為 7 位數字，不以姓名猜配；PF 為 50–900 L/min 整數。
- 手機可查已存在於 ACT / PF 雲端資料中的過去日期，每次最多 31 天；查歷史必須填精確病歷號，空病歷號只允許今天全部記錄。
- 桌面向外同步只選取 `assessment_type=PF`、`source=manual`、`entered_by` 為 `tray_panel` 或 `dr_tsui`、未被 supersede，且登錄時間 `created_at` 不早於固定 `bootstrapSince` 的遞增 ID。首次啟動用當時往前 31 天建立此起點並存入 state；之後不隨日期滑動。若 2026-11-24 首次啟動，起點為 2026-10-25。實際量測日期可以更早；這不是全歷史匯出。
- `dr_tsui` 是既有 Overlay PF API 的預設輸入者標記；它與 `tray_panel` 都必須同時符合 `source=manual` 及未被 supersede。Google Sheets 來源與失效／被 supersede 的列不匯出。
- `source=manual_edit` 的更正紀錄不在上述來源白名單；本輪沒有實作更正版本的自動回傳。不要把它標成已同步，也不要把舊事件移植到新來源。
- 獨立 PF 是獨立量測，不會自動回填 ACT Records 的 P 欄，不會製造空分數 ACT，不會改題目或寄信。
- 手機事件 ID 對同一重試保持不變；同一天已有量測時，必須明確確認新增另一筆。跨來源數值不自動覆寫。

## 程序與檔案

原始 worker 為 `bridge/pf_bridge.cjs`。私人執行目錄預設為 `%LOCALAPPDATA%\ClinicACTPFBridge`；安裝腳本管理已核對雜湊的 worker 副本 與隱藏啟動器。

| 檔案 | 用途 |
| --- | --- |
| `config.json` | `overlayBase`、既有 Apps Script `/exec` URL、私人 `syncKey`；不要提交到 Git |
| `state.json` | 已獲雲端確認的桌面向外同步 checkpoint `outCursor` 與首次固定 capture 起點 `bootstrapSince` |
| `health.json` | 最後一輪開始／完成時間、匯出、匯入、拒絕、pending 數量及錯誤代碼 |
| `install.json` | 已安裝 node／worker／launcher／Startup 捷徑路徑與安裝時間 |
| `start-hidden.vbs` | 隱藏啟動器；只啟動本 worker |
| `startup-error.txt` | 啟動失敗提示；不包含病歷或金鑰 |
| `worker.pid` | worker PID；停止前仍須驗證程序路徑與 command line |
| 已核對雜湊的 worker 副本、launcher | 安裝器產生的本機執行檔與隱藏啟動器 |

worker 使用 `127.0.0.1:18769` 作為單一工作程序鎖，Overlay 目標固定為 `http://127.0.0.1:8080/`。每輪完成後等 30 秒再進下一輪；一次網路請求上限 20 秒，批次與讀回可能使整輪長於 30 秒。健康檔時間不是固定每 30 秒更新的保證。`pending` 是本輪拉取的待辦數，不是本輪結束後尚未處理的完整佇列總量。

私人設定格式：

```json
{
  "overlayBase": "http://127.0.0.1:8080",
  "scriptUrl": "https://script.google.com/macros/s/EXISTING_DEPLOYMENT_ID/exec",
  "syncKey": "COPY_PRIVATE_ACT_PF_SYNC_KEY_HERE"
}
```

上述內容為格式範例，不是可用憑證。`ACT_PF_SYNC_KEY` 最少 32 字元；工作人員管理 PIN `ACT_NURSE_PIN` 最少 4 字元且不能全為空白，後端只接受此 Script Property 明確配置的值；缺少設定不會回退到任何歷史相容 PIN。這是工作人員登入 PIN，與服務同步金鑰分開管理。worker 不需要管理 PIN。

## 安裝、執行及停止

在專案根目錄的 PowerShell 中：

```powershell
.\bridge\install.ps1 -PrepareOnly
```

私人目錄及 ACL 已在準備階段建立；擁有者已於 15:20:53 完成 Google 初始化。下一步須取得使用者對本機金鑰保存的明確同意，再使用本機設定頁將同步金鑰寫入私人 `config.json`。未設定完整前不啟動正式同步。

```powershell
.\bridge\install.ps1 -StartIfConfigured
```

安裝器建立目前使用者的 Startup 捷徑，以本機隱藏 launcher 啟動已安裝 worker；它不是 Windows 服務，也不是 Overlay 的一部分。正式是否安裝、是否啟動，必須以實際腳本結果、捷徑與健康檔記錄確認。安裝器可能回傳 `ProcessVerified=true`、`HealthPending=true`：這代表程序身份已核對，但首輪仍在處理；應繼續等候新的 `health.json`，不能直接宣稱資料同步完成，也不要重複啟動。

```powershell
.\bridge\stop.ps1
.\bridge\stop.ps1 -RemoveStartup
```

第一個指令只停止經 PID、執行檔與已安裝 script 路徑核對的本 worker；第二個另移除本 worker 的 Startup 捷徑。兩者保留設定、checkpoint 與健康資料。不得用通殺 `node.exe`、`python.exe` 或重新啟動 Overlay 代替。

如需單輪正式觀測，`node .\bridge\pf_bridge.cjs --once` 會使用私人設定且可能執行真正同步；應在配對部署與設定確認後操作。離線測試請使用 `tests` 下的假 adapter，不要對正式病歷做測試寫入。

## 維護界線

- 不刪除 `state.json` 當作通用修復；先保留檔案並查明 checkpoint 與已確認事件。
- 更正已匯出的舊桌面列、補齊早於首次 `bootstrapSince` 的桌面歷史、變更事件來源，都需要獨立的版本／對帳方案。
- 停止 worker 可暫停跨端傳遞，但不撤銷 Google 或 Overlay 已保存的紀錄。
- 真實患者是否顯示、是否寫回正確日期，需要正式事件讀回與現場驗收；合成測試或 QA ACK 不足以證明。

## 既有 poller 的已知交互作用

bridge 新增的是獨立 manual PF，本身不覆寫舊量測。但既有 Sheets poller 仍可能把「同病歷、同日且數值等於 ACT P 欄」的 manual PF 標為 superseded；稽核列保留，最新視圖可能隱藏它。ACT P 留空的獨立 PF 路徑不受這個同值條件影響。本輪沒有修改或重新載入既有 Overlay poller，不能宣稱此交互作用已消除。

## 可重現的隔離驗證

在專案根目錄執行：

```powershell
node --test tests/backend-pf.cjs tests/backend-history-contract.cjs tests/frontend-nurse.test.cjs tests/frontend-history.test.cjs tests/integration-nurse.cjs tests/bridge-adversarial.cjs
```

接受明確配置的 4 碼管理 PIN 後，候選完整六檔驗證回報 108 項通過、0 失敗。所有 adapter 為合成替身，不連 Google、Overlay 或任何資料庫。

## 使用者指定管理 PIN 的後續修訂

候選 `2026-11-24-pf-history2` 接受 `ACT_NURSE_PIN` 已明確配置的至少 4 字元管理 PIN，不設前端 PIN 值、不在缺少 property 時使用預設回退。服務端 `ACT_PF_SYNC_KEY` 仍需至少 32 字元，沒有放寬。候選測試已驗證指定 4 碼 PIN 可登入、缺少屬性仍拒絕；Google 屬性的實際輸入及正式登入確認仍由使用者／主代理完成，不因離線測試而宣稱已變更。
