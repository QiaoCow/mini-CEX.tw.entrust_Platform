# 架構說明：程式碼公開，資料與權限不公開

## 先講一個無法迴避的事實

**前端程式碼藏不住。** 只要網頁在瀏覽器裡跑，程式碼就已經在使用者電腦上，按 F12 就看得到。壓縮、混淆、改副檔名都只是增加閱讀難度，擋不住有心人。這不是這個專案的限制，是所有網頁應用的共同前提。

所以正確的問題不是「怎麼把前端藏起來」，而是：

> **就算有人把前端程式碼看光、任意竄改，他能拿到不該拿的資料嗎？**

答案應該是「不能」。做到這件事，程式碼公開在 GitHub 就完全沒有問題，就像銀行 ATM 的操作介面設計公開不影響金庫安全。

---

## 目前的問題

改版前的程式有兩個致命點：

**一是身分由前端決定。** 讀取資料時，角色是用網址參數傳的：

```
.../exec?role=deptAdmin&staffId=M70001&scopeDept=內科部
```

把 `role` 改成 `hospital`，後端就照全院範圍回傳資料。不需要任何技術，改個網址就行。

**二是資料全部送到前端才過濾。** 後端把十二張表整包吐出，前端再依角色顯示。這表示範圍外的資料其實已經到了使用者電腦上，只是介面沒畫出來。打開開發者工具的網路分頁就能看到完整內容。

這兩點加起來，等於實際上沒有權限控制。

---

## 改版後的架構

### 身分由後端取得

```javascript
var email = Session.getActiveUser().getEmail();   // Google 登入狀態，前端無法偽造
```

拿到 email 後回查 Users 表決定角色與負責科部。前端送來的任何角色參數一律忽略。

### 資料在回傳前就裁切

```javascript
payload.Evaluations = scopeEvaluations(payload.Evaluations, me.roles, me.staffId, me.scopeDept);
payload.DeptStats   = deptStatsFor(me.roles, me.scopeDept);
```

科部承辦人拿到的 JSON 裡就只有自己科的資料，範圍外的內容從未離開伺服器。稽核紀錄只有系統管理者的回應中才有。

### 寫入權限也在後端判斷

```javascript
function canWrite(roles, sheet) {
  if (sheet === 'AuditLogs' || sheet === 'AuditChanges') return false;   // 任何人都不可改稽核
  if (has(roles, 'sysAdmin')) return true;
  if (sheet === 'Evaluations') return has(roles, 'teacher');
  ...
}
```

被拒絕的寫入同樣留下稽核紀錄，記下是誰試圖做什麼。

---

## 為什麼前端要改由 Apps Script 提供

這是這次架構調整的核心，也是最容易被忽略的一點。

要取得 `Session.getActiveUser().getEmail()`，Apps Script 必須部署為「執行身分：存取應用程式的使用者」。但這種部署會要求 Google 登入，而登入是透過網頁跳轉完成的。

如果前端放在 GitHub Pages（`qiaocow.github.io`），它去呼叫 `script.google.com` 屬於跨來源請求。瀏覽器的 `fetch()` 無法完成 Google 的登入跳轉，請求會失敗。這不是設定問題，是瀏覽器的安全模型使然。

解法是把前端也交給 Apps Script 提供。使用者開啟的網址就是 Apps Script 的網址，登入由 Google 處理，前後端同源，身分自然帶得到。

---

## 兩個部署目標，一份原始碼

| | GitHub Pages | Apps Script |
|---|---|---|
| 用途 | 公開展示、設計討論 | 實際使用 |
| 網址 | `qiaocow.github.io/...` | `script.google.com/macros/s/.../exec` |
| 資料 | 內建虛構示範資料 | Google 試算表真實資料 |
| 登入 | 點選示範帳號 | Google 帳號 |
| 權限 | 前端模擬，僅供展示 | 後端強制執行 |
| 誰能開 | 任何人 | 名冊內的院內帳號 |

GitHub 儲存庫存放原始碼與展示版，讓人看得到設計、看得到修改紀錄。真實資料完全不經過 GitHub。

這樣做還有一個好處：展示版永遠可以給人看，不必擔心洩漏任何真實資訊，因為裡面本來就只有虛構資料。

---

## 部署步驟

### 展示版（GitHub Pages）

沿用現行做法，把 `index.html` 上傳到儲存庫根目錄即可。設定區的 `window.__SHEETS_API__` 保持註解狀態，讓它使用內建示範資料。

### 正式版（Apps Script）

1. 在 Apps Script 專案中新增一個 HTML 檔案，命名為 `index`
2. 把打包好的 `index.html` 內容整份貼進去
3. 確認 `Code.gs` 最上方的 `SPREADSHEET_ID` 已填入試算表 ID
4. 部署 → 新增部署作業 → 網頁應用程式
   - 執行身分：**存取應用程式的使用者**（這一項與展示版不同，務必確認）
   - 具有存取權的使用者：**機構內的任何使用者**
5. 在 Users 表的 `email` 欄填入每個人的院內信箱，這是身分比對的依據

部署後，未登入者會先看到 Google 登入頁；登入後若信箱不在 Users 表中，會收到明確的錯誤訊息而非空白畫面。

---

## 仍然存在的限制

**Apps Script 取不到呼叫端 IP。** 稽核紀錄的 IP 欄位一律記為「(未取得)」。若稽核要求必須包含來源位址，需要改用真正的伺服器。

**Users 表本身可被有試算表編輯權的人修改。** 有人若能編輯試算表，就能把自己的信箱改成 sysAdmin。試算表的共用權限必須嚴格控制，且異動應納入稽核。

**稽核紀錄存在試算表中，可被直接竄改。** 正式上線應寫入僅可新增的儲存體，並定期異地備份。

**Apps Script 有配額限制。** 使用人數多時可能觸及執行時間與呼叫次數上限，屆時需評估遷移到正式後端。

這幾點都不是這次改版能解決的，但把它們寫下來，日後決策時才不會誤以為已經處理好了。
