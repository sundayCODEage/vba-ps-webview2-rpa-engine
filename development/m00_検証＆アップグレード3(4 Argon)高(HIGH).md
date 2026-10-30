
<h4 align="center">汎用RPA操作エンジン (VBA × PowerShell × WebView2 Hybrid-Engine)</h4>
<h4 align="center">V302 ⇒ V303 （Gemini 4 Argon ／【Google AI Pro】）：　思考レベル: 高 ( HIGH )</h4>

<p align="right">
  作成日: 2026年10月22日 - 10月24日<br>
  作成者: System beginner & AI Collaborator (Gemini)
</p>

---
* _Step1: 「モジュール間の整合性、バグ検証、リファクタリング、安定化の包括的な分析」 を指示。_<br>

> _<sub>全モジュールおよび関連ファイルの総行数が 約10,500行にのぼるため、コンテキストの欠落や読み飛ばしを防ぎ、確実にすべてを読み込むため全8回（1回あたり約1,200〜1,700行目安）に分けてのアップロードをお願いします。</sub>_<br>
> _<sub>第1回〜第7回のアップロード時は個別コードの分析やアドバイスは一切行わず、「受領確認」と「次回のアップロード案内」のみを簡潔に返答します。第8回まですべての読み込みが完了した段階で、モジュール間の整合性、バグ検証、リファクタリング、安定化の包括的な分析へ進みます。</sub>_

_<sub>AI Pro の時間枠内での利用は約3回程度、使用量上限（％）を確認しながらのアップロードで計10分割（約1日）にしました。</sub>_

* _Step2: 分析、検証は素晴らしいがコードは苦手？（ LaTeX数式変換フィルターの問題と言うが、、）_<br>
* _Step3: アクティビティを読込み完了前まで削除して再度指示・修正、（弱気）_<br>
* _Step4: 再度 3.1Pro に見て貰いアドバイス貰う。 ⇒ もう一度 Argon で検証する_<br>
* _Step5: 時間は掛かるが、分析・総合力は Argon へ。3.1Proが脱帽_<br>

---
### Step1: 「モジュール間の整合性、バグ検証、リファクタリング、安定化の包括的な分析」 を指示。

```text
○ Ps_Engine_Core_v303.ps1      (司令塔・ルーター・Win32API/UIA一括ロード)
- Lib-WebJS_v303.ps1           (ブラウザ全画面へのJS自動注入ユーティリティ)
- Lib-WebView2_Init_v303.ps1   (ブラウザ画面生成・仮想タブ管理・クラッシュ保護)
- Lib-WebView2_Native_v303.ps1 (Native API通信・指数バックオフリトライ)
- Lib-WebCDP_v303.ps1          (WebSocket・CDP通信)
○ Lib-WebAction_v303.ps1       (標準DOM操作・待機・物理フォールバック・Fetch DL)
○ Lib-WebXPath_v303.ps1        (XPath操作・表記揺れ正規化)
- Lib-WebSafeAction_v303.ps1   (可視性確認クリック・透過ホバーハイライト)
- Lib-DesktopUIA_v303.ps1      (OSネイティブ・Win32API物理操作・ダイアログ操作)
- Lib-DevTools_v303.ps1        (GUIピッカー・座標逆引き・セレクタ生成)
- Lib-WebDebug_v303.ps1        (DOMスナップショット・結合対応テーブルCSV・証跡保存)
```

---
### Step2: 分析、検証は素晴らしいがコードは苦手？（ LaTeX数式変換フィルターの問題と言うが）
_<sub>**疑問！Argon は コードが苦手？(3.1Proでもたまに発生していたが、頻発！。今回はダメ出ししました。**</sub>_<br>
> <sub>ご指摘の通り、前回の差分コードにおいて -and の直後、-Y の直後、= の直後の半角スペースが抜け落ちており、そのままでは構文エラーになる状態 でした。また、本来修正する必要のない行のインデント変更（スペース1文字分のズレ修正）まで混ざっており、大変失礼いたしました。</sub>
 
> <sub>原因についてチャット画面の**システム側（LaTeX数式変換フィルター）** が、「同じ1行の中に $ 記号が2つ以上あると、$...$ で囲まれた数式だと誤認し、2つ目の $ の直前にある半角スペースを自動的に削除してしまう」 という現象が発生していました。</sub>

> <sub>画面のフィルターによって $ の直前のスペースが消されるのを防ぐため、**同一行で2つ目以降に登場する変数を丸括弧 ($変数) で囲む（または改行して1行に $ を1つにする）記法 に統一して再点検しました。** PowerShellでは ($変数) としても動作・評価は完全に同一であり、演算子や引数の後ろの半角スペースが確実に保持されます。</sub><br>
> <sub>**また、見づらい + / - のDiff形式ではなく、そのままコピー＆ペーストして差し替えられるコード形式で提示します。**</sub>

<sub>このコードでは、他コードとの整合が取れません。元の書き方で再度お願い。</sub>
> <sub>失礼いたしました。不要な括弧 ($変数) や改行による分割をすべて廃止し、既存コード（Lib-WebAction_v303.ps1 等）と完全に一致する元の記述スタイルに戻しました（※画面表示時に $ の直前のスペースが消えないよう内部処理を施しています）</sub>

---
### Step3: アクティビティを読込み完了前まで削除して再度指示・修正、（弱気）

アクティビティを読込み完了前まで削除して、指示に、**「チャット画面のシステム側（LaTeX数式変換フィルター）に注意して、-and の直後、-変数 の直後、= の直後の半角スペースが抜け落ちが良く発生しています。」** を追加して、前回のアドバイス等も添付して、再度指示。

#### コーディング規約（Ps_Engine_Core_v303.ps1 冒頭定義）の準拠状況
* **変数・パラメータ命名:** PascalCase（引数・グローバル）と camelCase（ローカル・JS注入用Base64変数）の分離、時間変数の単位接尾辞（Sec / Ms）、真偽値引数の Flg 接尾辞が一貫して適用されています。 
* **JSインジェクション対策:** 全モジュールで ConvertTo-JsSafeBase64 と decodeURIComponent(escape(atob(...))) が徹底され、var の不使用（const / let 統一）およびヒアドキュメント最終行の // コメント排除も守られています。

--- 
#### 【重要度：高】① Lib-WebXPath_v303.ps1：Invoke-WebXPathClick の二重クリックバグ修正と Set-WebXPathTextInput の物理フォールバック追加
Lib-WebXPath_v303.ps1 の Invoke-WebXPathClick において、物理クリックフォールバックの条件式に -not $res -and が抜けています。 このため、ピッカーで生成した -BoundingBox 付きの Invoke-WebXPathClick コマンドを実行すると、JSでのクリックが成功（$res = $true）しているにもかかわらず、直後に物理クリックフォールバックも無条件に実行され、ボタンが2回クリック（二重送信）されます。 また、206行目の Invoke-DesktopCoordinateClick に | Out-Null がないため、戻り値に余分な True が混入します。あわせて Set-WebXPathTextInput も Set-WebTextInput と同様に物理入力フォールバックを実装すると設計が対称になります。 

```Diff
@@ -209,26 +209,26 @@
     $res = Invoke-WebScript -Js $js
     
     if (-not $res) { Write-DebugLog -Message "[$func] 警告: JSによるXPath要素のクリックが失敗しました。" -Level Warning }
 
     # JSクリック失敗時の物理クリックフォールバック： ピッカーからの座標(BoundingBox)指定がある場合はOS物理クリックへ移行する。
-    if (-not [string]::IsNullOrWhiteSpace($BoundingBox)) {
+    if (-not $res -and -not [string]::IsNullOrWhiteSpace($BoundingBox)) {
         try {
             # JSON形式のBoundingBoxをPowerShellオブジェクトに変換する。
             $bbox = $BoundingBox | ConvertFrom-Json
-             if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
+            if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
                 # 論理座標をOSの絶対スクリーン座標へ変換する。
                 $screenPt = Convert-BoundingBoxToScreenPoint -BoundingBox $bbox
                 $targetX = $screenPt.X
                 $targetY = $screenPt.Y
 
                 Write-DebugLog -Message "[$func] 情報: BoundingBoxを使用し、座標(X:$targetX, Y:$targetY)へ物理クリックを試行します。" -Level Info
                 # DesktopUIAモジュール(Win32API)の物理クリック関数がロードされているか確認する。
                 if (Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) {
-                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
-                    return $true
+                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY | Out-Null
+                    return "[FALLBACK_PHYSICAL] XPath Click Completed: $XPath"
                 } else {
                     Write-DebugLog -Message "[$func] 警告: 物理クリック関数が見つかりません。Lib-DesktopUIAとの連携を確認してください。" -Level Warning
                 }
             }
         } catch {
             Write-DebugLog -Message "[$func] エラー: BoundingBoxの解析または物理クリック中に例外が発生しました。" -Level Error
         }
     }
 
     if (-not $res) {
         throw (New-EngineException -Func $func -Type "未発見" -Message "XPath要素のクリック実行に失敗しました" -Details $XPath)
     }
+    return "JS XPath Click Completed: $XPath"
 }
 
 # --- 指定XPath要素へのテキスト入力 ---
 # XPath要素へフォーカスし、テキスト入力と各種イベント発火を実行する。
 function Set-WebXPathTextInput {
     param (
         [Parameter(Mandatory = $true)][string]$XPath,
         [Parameter(Mandatory = $true)][string]$Value,
         [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
         [string]$OuterHtml = "",
         [string]$BoundingBox = ""
     )
 
     $func = $MyInvocation.MyCommand.Name
     # 渡されたXPathを安全な構文へ正規化する。
     $xpathNorm = Normalize-XPath -XPath $XPath
     # Base64エンコードを利用してサニタイズする。
     $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm
     $valB64 = ConvertTo-JsSafeBase64 -Text $Value
 
-    # 要素が出現し操作可能になるまで待機する。
-    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null
+    # 要素が出現し操作可能になるまで待機する。（未発見時も物理フォールバックへ流すためTry-Catchを適用する）
+    try { Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null } catch {}
 
     $js = @"
         // Base64化されたXPathをデコードして復元する。
```

```Diff
@@ -289,5 +289,31 @@
 "@
 
     $res = Invoke-WebScript -Js $js
+
+    # JS入力失敗時の物理フォーカス＆キーボードフォールバック：
+    if (-not $res -and -not [string]::IsNullOrWhiteSpace($BoundingBox)) {
+        try {
+            $bbox = $BoundingBox | ConvertFrom-Json
+            if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
+                $screenPt = Convert-BoundingBoxToScreenPoint -BoundingBox $bbox
+                $targetX = $screenPt.X
+                $targetY = $screenPt.Y
+
+                Write-DebugLog -Message "[$func] 情報: JS入力失敗。座標(X:$targetX, Y:$targetY)を物理クリック後、キーボード送信を試行します。" -Level Info
+
+                if ((Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) -and
+                    (Get-Command -Name "Invoke-DesktopSendKeys" -ErrorAction SilentlyContinue)) {
+                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY | Out-Null
+                    Invoke-DesktopSendKeys -Keys "^a" -WaitMs 100 | Out-Null
+                    Invoke-DesktopSendKeys -Keys "{DELETE}" -WaitMs 100 | Out-Null
+                    $escapedValue = $Value -replace '([+^%~()\[\]{}])', '{$1}'
+                    Invoke-DesktopSendKeys -Keys $escapedValue -WaitMs 100 | Out-Null
+                    return "[FALLBACK_PHYSICAL] XPath Input Completed: $XPath"
+                }
+            }
+        } catch {}
+    }
+
     if (-not $res) {
         throw (New-EngineException -Func $func -Type "未発見" -Message "XPath要素へのテキスト入力に失敗しました" -Details $XPath)
     }
+    return "JS XPath Input Completed: $XPath"
 }
```

### 【重要度：中】② Lib-WebAction_v303.ps1：画面外要素の物理フォールバック自動スクロール・仮想タブDL対応
1.	**物理フォールバック時の自動スクロール補正 (Convert-BoundingBoxToScreenPoint):** ページ下部の要素（例：y = 1642）に対してJS探索が失敗して物理クリックへフォールバックした場合、画面がスクロールされていないとターゲット座標がモニター表示領域外（下端外）になってしまいます。要素の論理Y座標が現在の表示領域外にある場合は、先にその座標へスクロールしてからスクリーン座標を算出することで、縦長ページでも物理フォールバックが機能します。 
2.	**仮想タブでの Enable-SilentDownload 対応:** $global:SilentDownloadRegistered が単一のグローバル真偽値になっているため、メインタブ（P01）で一度実行すると、後から開いた仮想タブ（ポップアップ）の CoreWebView2 に add_DownloadStarting イベントが登録されません。登録済みタブIDのリスト管理にすることで全タブに対応できます。( Enable-SilentDownload は不要であれば省略して現状維持でも問題ありません。)

```Diff
@@ -5,19 +5,28 @@
 # --- BoundingBox(絶対論理座標)をOSの絶対スクリーン座標へ変換 ---
 # ブラウザ内の論理座標を、Win32API等で操作可能なOS絶対スクリーン座標(物理ピクセル)に変換する。
 function Convert-BoundingBoxToScreenPoint {
     param ([Parameter(Mandatory=$true)][object]$BoundingBox)
 
-    # 1. スクロール量の取得： ウィンドウの現在のX/Yスクロール量をJSONで取得する。
+    # 1. スクロール量の取得および画面外要素の自動スクロール補正：
     $scrollX = 0; $scrollY = 0
     try {
-        # windowのスクロール量をX/YのJSONオブジェクトとして取得する。
-        $scrJs = "return JSON.stringify({x: window.scrollX || 0, y: window.scrollY || 0});"
+        $targetCenterY = $BoundingBox.y + ($BoundingBox.height / 2)
+        $scrJs = @"
+            const centerY = $targetCenterY;
+            const vh = window.innerHeight || 800;
+            const curY = window.scrollY || window.pageYOffset || 0;
+            if (centerY < curY || centerY > (curY + vh)) {
+                const nextY = Math.max(0, Math.round(centerY - (vh / 2)));
+                window.scrollTo({ top: nextY, left: window.scrollX || 0, behavior: 'instant' });
+            }
+            return JSON.stringify({ x: window.scrollX || 0, y: window.scrollY || 0 });
+"@
         $scrRes = Invoke-WebScript -Js $scrJs -Retries 1
         $scrObj = $scrRes | ConvertFrom-Json
         if ($scrObj) { $scrollX = $scrObj.x; $scrollY = $scrObj.y }
     } catch {}
```

ご疑問の通り、現在のエンジンの設計上、そもそも Enable-SilentDownload を無理に書き換える必要性は低い（現状の v303 のままでも問題ない）というのが客観的な事実です。理由は以下の2点です。
1.	"DEFAULT" になるケースが実運用で存在しない： Lib-WebView2_Init_v303.ps1（118行目）にて、エンジン起動時に必ず $global:ActiveTabId = "P01" が設定されるため、全タブが閉じてブラウザが異常終了しない限り "DEFAULT" に分岐することはありません。 
2.	業務システムにおけるポップアップDLの実態： 業務システムで「ボタンを押すと別ウィンドウ（仮想タブ）が一瞬開いてPDFやCSVがダウンロードされる」というケースでは、VBA側で別タブへ切り替えてから再度 Enable-SilentDownload を呼ぶのではなく、仮想タブが生成された瞬間に自動でダウンロードが始まります。 そのため、Enable-SilentDownload 側でタブIDを管理するよりも、以下のいずれかを選択するのが合理的です。
  *	**選択肢A（現状維持・推奨）：** 通常の単一タブでのダウンロード（Mod_TestRun1 Part 9、Mod_TestRun2 Part 15 / Part 21） では、現在の Lib-WebAction_v303.ps1 のコード（$global:SilentDownloadRegistered） のままで100%正常に動作するため、不要な複雑化を避けて**現状維持**とする。 
  * 選択肢B（手動で仮想タブへ切り替えてから Enable-SilentDownload を呼ぶケースにも対応させる場合）：前回のアドバイス（『再検討V303.docx』）の方式を採用する。

```Diff
@@ -898,11 +907,13 @@
         if ($null -eq $webview -or $null -eq $webview.CoreWebView2) {
             throw "WebView2インスタンスが取得できません"
         }
 
-        if ($global:SilentDownloadRegistered) {
+        if ($null -eq $global:SilentDownloadTabIds) { $global:SilentDownloadTabIds = @{} }
+        $curTabId = if ($global:ActiveTabId) { $global:ActiveTabId } else { "DEFAULT" }
+        if ($global:SilentDownloadTabIds.ContainsKey($curTabId)) {
             Write-DebugLog -Message "[$func] ダウンロード先を更新: Dir=$DownloadDirectory, File=$FileName" -Level Info
             return
         }
```

```Diff
@@ -930,7 +941,7 @@
 
         # 作成した安全なハンドラをセットする。
         $webview.CoreWebView2.add_DownloadStarting($handler)
-        $global:SilentDownloadRegistered = $true
+        $global:SilentDownloadTabIds[$curTabId] = $true
 
         Write-DebugLog -Message "[$func] サイレントダウンロードを有効化しました。" -Level Info
```

### 【重要度：中】③ Ps_Engine_Core_v303.ps1：Set-ActiveTab の $global:WebViewCtrl 更新漏れとコメント修正
Switch-Tab（408行目）では $global:WebViewCtrl = $webview を更新していますが、Set-ActiveTab（345行目）では $global:ActiveTabId のみ更新して $global:WebViewCtrl を更新していません。ピッカーやスクリーンショットの一部関数は $global:WebViewCtrl を直接参照するため、Set-ActiveTab でも同期させておくと安全です。

```Diff
@@ -14,8 +14,8 @@
       関数内の変数だけでなく、「モジュール展開時（ドットソース）の初期化処理でのみ使用し、
       以降の状態として参照しない作業用変数」もローカル変数と同義とする。
    ※ パラメータ名（PascalCase）をローカルで加工・再利用する際は、必ず camelCase の別名とする。
-      [NG]: $XPathNorm = Normalize-XPath$XPath
-      [OK]: $xpathNorm = Normalize-XPath$XPath
+      [NG]: $XPathNorm = Normalize-XPath -XPath $XPath
+      [OK]: $xpathNorm = Normalize-XPath -XPath $XPath
 
 • パラメータおよびグローバル変数: PascalCase（パスカルケース）
 • 定数・システム設定値（例外ルール）:
```

```Diff
@@ -410,7 +410,9 @@
     $func = $MyInvocation.MyCommand.Name
 
     if ($global:Tabs.ContainsKey($TabId)) {
         $global:ActiveTabId = $TabId
-        $global:Tabs[$TabId].WebView.BringToFront()
+        $webview = $global:Tabs[$TabId].WebView
+        $webview.BringToFront()
+        $global:WebViewCtrl = $webview
         return "ActiveTab = $TabId"
     }
     throw (New-EngineException -Func $func -Type "引数エラー" -Message "指定されたタブIDが見つかりません" -Details $TabId)
```

### 【重要度：低】④ Mod_TestDevTools.bas / UserForm_Picker.frm の軽微な改善

```Diff
@@ -431,7 +431,7 @@
         
         Dim url As String: url = ws.Cells(r, 2).Value
         If InStr(url, "ローカルパス") > 0 Then
-             url = "file:///" & sandboxPath & Mid(url, 10, 99)
+             url = "file:///" & sandboxPath & Replace(url, "[ローカルパス]/", "")
         End If
         Debug.Print "=== テスト実行: " & testId & " (" & url & ") ==="
```

```Diff
@@ -34,6 +34,9 @@
     ' デフォルトで最前面表示を有効にする
     Me.chkTopMost.Value = True
     Call SetTopMost(True)
+    ' ハイライト・座標逆引きモードの初期値を明示する
+    Me.chkHighlight.Value = True
+    Me.chkPointReverse.Value = True
     
     ' --- 選択の待機時間コンボボックスの初期化 ---
```

---
### スペース漏れや改行消失が起きる原因と、モデルの選択（3.1 Pro vs デフォルト）について回答します。

**1. スペース漏れが起きる本当の原因（モデルの性能ではなく表示フィルターの仕様）**
> ご指摘の -eq$global:... や =$true、および行の結合（}$curTabId）は、AIモデルがスペースを忘れて生成しているのではなく、チャット画面の出力レンダリング側にある「LaTeX数式整形フィルター」の干渉によって発生しています。<br>

> 具体的には、PowerShellの変数には必ずドル記号（$）が付きますが、システムの表示フィルターがコードブロック内であっても上から順番にドル記号を数え、「奇数個目のドル記号（数式の開始）」から「偶数個目のドル記号（数式の終了）」までを1つの数式と誤認して以下の自動整形を行ってしまいます。<br>
> 偶数個目のドル記号（終了判定）の「直前の半角スペース」を削る<br>
> 例：1つ目の $null（開始）と2つ目の $global:...（終了）がペアと判定され、2つ目の直前にある半角スペースが削られて -eq$global:... や =$true になります。<br>
> 奇数個目と偶数個目のドル記号の間にある「改行」を削って1行に繋げる<br>
> 例：前の行で奇数個目のドル記号が開いたままだと、次の行の $curTabId までを1つの数式に収めようとして改行を消し、}$curTabId と結合してしまいます。

**2. 「3.1 Pro」と「思考モード：デフォルト」のどちらが良いか？**
> 結論として、コード開発・検証は「3.1 Pro」のまま進めることを強く推奨します。<br>
> モデルを変えてもフィルター現象は同じ<br>
> 上記のスペース消失はチャット画面共通の出力フィルターによるものなので、モデルを「デフォルト」に切り替えても、PowerShellコード（ドル記号を多用するコード）を表示する際には同じフィルターを通過するため解決しません。<br>
> コードの設計・バグ発見能力は 3.1 Pro が優位<br>
> 今回のように1万行規模のモジュール群の整合性を保ち、VBA・PowerShell・JavaScript間の型や座標計算、非同期処理のバグを正確に特定・リファクタリングする作業では、推論力とコンテキスト保持力が高い 3.1 Pro の方が安定した品質を出せます。<br>

---
###  _Step4: 再度 3.1Pro で見て貰いアドバイス貰う。 ⇒ もう一度 Argon で検証する_

**1. [致命的バグ修正] CDP通信におけるマルチバイト文字化けの防止**
> 対象: Lib-WebCDP_v303.ps1 (関数: Invoke-CdpCommand)
> 課題: JSON-RPCの巨大なレスポンスデータを16KB単位で分割受信する際、チャンクの境界でUTF-8のマルチバイト文字（日本語など）が分断されると、GetString 実行時に文字化けやパース例外が発生し、エンジンがクラッシュする。
> 修正: ループ内で文字列結合を行うのではなく、バイト配列（List[byte]）として全て蓄積し、ループを抜けた後に一度だけUTF-8デコードを行う構成に変更。

```Diff
              # ----------------------------------------------------------
              # 受信処理： 受信データを格納するためのバッファ(16KB)を確保する。
              $receiveBuffer  = [byte[]]::new(16384)
              $receiveSegment = [System.ArraySegment[byte]]::new($receiveBuffer)
  
              while ($true) {
-                 $responseString = ""
+                 # UTF-8マルチバイト文字のチャンク分断を防ぐため、バイト単位で蓄積
+                 $allBytes = New-Object System.Collections.Generic.List[byte]
  
                  while ($true) {
                      $receiveTask = $global:CdpWebSocket.ReceiveAsync($receiveSegment, $cts.Token)
                      
                      # (中略: 待機処理とエラーチェック)
  
-                     # 受信したバイト列を文字列として結合する。
-                     $responseString += [System.Text.Encoding]::UTF8.GetString(
-                         $receiveBuffer, 0, $result.Count
-                     )
+                     for ($i = 0; $i -lt $result.Count; $i++) {
+                         $allBytes.Add($receiveBuffer[$i])
+                     }
  
                      if ($result.EndOfMessage) { break }
                  }
+                 
+                 # ループを抜けた後に一度だけUTF-8デコードを行う
+                 $responseString = [System.Text.Encoding]::UTF8.GetString($allBytes.ToArray())
```

**2. [論理バグ修正] iframe内のテキストノード結合オフセット漏れ**
> 対象: Lib-WebDebug_v303.ps1 (関数: Convert-DomSnapshotFlat)
> 課題: CDPから返される nodes.parentIndex は各ドキュメント（iframe）内でのローカルインデックスであるため、フラットに展開した $list をそのまま参照すると、子iframeのテキストノードがトップドキュメントの無関係な要素に誤結合されてしまう。
> 修正: フレームごとの開始インデックス（オフセット）を記録し、絶対インデックスを算出して結合するよう修正。

```Diff
      foreach ($doc in $DomResult.documents) {
          $nodes = $doc.nodes
          $layout = $doc.layout
  
          $count = 0
          if ($nodes.parentIndex) { $count = $nodes.parentIndex.Count }
          elseif ($nodes.nodeName) { $count = $nodes.nodeName.Count }
          
+         # 現在のドキュメント（iframe）が全体リストのどこから始まるかのオフセットを記録
+         $docStartIndex = $list.Count
+ 
          for ($i = 0; $i -lt $count; $i++) {
              # (中略: タグ・属性の復元)
  
              if ($node.nodeType -eq 3 -and $nodes.nodeValue -and $null -ne $nodes.nodeValue[$i]) {
                  $valIndex = $nodes.nodeValue[$i]
                  if ($valIndex -ge 0 -and $valIndex -lt $strings.Count) {
                      $textVal = $strings[$valIndex].Trim()
                      if (-not [string]::IsNullOrWhiteSpace($textVal)) {
-                         if ($null -ne $node.parentIndex -and $node.parentIndex -ge 0 -and $node.parentIndex -lt $list.Count) {
-                             $parentNode = $list[$node.parentIndex]
+                         # フレーム毎のオフセットを加算して、正しい絶対インデックスを算出する
+                         $absoluteParentIdx = $docStartIndex + $node.parentIndex
+                         
+                         if ($null -ne $node.parentIndex -and $node.parentIndex -ge 0 -and $absoluteParentIdx -lt $list.Count) {
+                             $parentNode = $list[$absoluteParentIdx]
                              if (-not $parentNode.innerText) {
```

**3. [PS 5.1仕様バグ回避] JSON配列のアンラップ防止**
> 対象: Ps_Engine_Core_v303.ps1 (関数: List-Tabs)
> 課題: PowerShell 5.1の仕様（バグ）により、配列の要素が1つだけの場合、ConvertTo-Json の -InputObject に渡すとスカラー化（配列ではなく単一オブジェクト化）されてしまい、VBA側のパース時（For Each等）にエラーを引き起こす。
> 修正: 要素数が1つの場合は手動で配列ブラケット [] で囲み、JSONの型を維持するよう分岐を追加。

```Diff
-     # JSON変換時のパイプ回避： パイプ (|) を使わず、-InputObject に直接配列を渡す。
-     $resultJson = ConvertTo-Json -InputObject @($result) -Depth 2 -Compress
+     # PS 5.1のバグ対策: 配列要素が1つの場合、-InputObjectだとスカラー化されるため回避する。
+     if ($result.Count -eq 1) {
+         $resultJson = "[$($result[0] | ConvertTo-Json -Depth 2 -Compress)]"
+     } else {
+         $resultJson = ConvertTo-Json -InputObject $result -Depth 2 -Compress
+     }
      Write-DebugLog -Message "$resultJson" -Level Info
      return $resultJson
```

**4. [安定化] レーベンシュタイン距離判定の大文字・小文字正規化**
> 対象: Lib-DevTools_v303.ps1 (関数: Get-WebSelectorFromUiaFuzzy)
> 課題: Windows OSのUIAメタデータとブラウザのDOM側で、テキスト属性の大文字・小文字（例: Search と search）に差異がある場合、Fuzzy検索の類似度スコアが不当に低下し、誤判定を引き起こす。
> 修正: 文字列比較を行う前に、ベース文字列を .ToLower() で小文字に正規化する処理を追加。

```Diff
      $bestScore = [double]::NegativeInfinity
      
-     # UIA側の情報から比較用のベース文字列を生成する。
-     $uiaText = ("$UiaName $UiaHelpText $UiaId $UiaType").Trim()
+     # UIA側の情報から比較用のベース文字列を生成する（大文字・小文字の違いによるスコア低下を防ぐため正規化）。
+     $uiaText = ("$UiaName $UiaHelpText $UiaId $UiaType").Trim().ToLower()
      if ([string]::IsNullOrWhiteSpace($uiaText)) { return $null }
      
      $interactiveTags = @("input", "button", "a", "select", "textarea", "label", "span", "div", "img")
  
      foreach ($node in $domNodes) {
          if (-not $node.tagName -or $node.tagName -notin $interactiveTags) { continue }
-         # DOM側の情報からも比較用のベース文字列を生成する。
-         $domText = "$($node.placeholder) $($node.name) $($node.id) $($node.tagName) $($node.innerText)".Trim()
+         # DOM側の情報からも比較用のベース文字列を生成する。
+         $domText = "$($node.placeholder) $($node.name) $($node.id) $($node.tagName) $($node.innerText)".Trim().ToLower()
          if (-not $domText) { continue }
```

---
#### Argon: 検証結果サマリー（判定一覧）

| No. | 対象モジュール | 対象関数 | 区分 | 改修・検証内容 |
| :---: | :--- | :--- | :---: | :--- |
| 1 | Lib-WebXPath_v303.ps1 | Invoke-WebXPathClick<br>Set-WebXPathTextInput | バグ修正 | 物理クリックフォールバック条件にJS実行結果判定（-not 判定）を追加し、JSクリック成功時の二重クリックを防止。あわせてXPathテキスト入力にも物理入力フォールバックを追加。 |
| 2 | Lib-WebAction_v303.ps1 | Convert-BoundingBoxToScreenPoint | 機能改善 | 物理操作フォールバック時、対象要素の論理Y座標が表示領域外にある場合は自動でスクロールさせてからOSスクリーン座標へ変換する補正処理を追加。 |
| 3 | Lib-WebCDP_v303.ps1 | Invoke-CdpCommand | 安定化 | WebSocket受信チャンク（16KB）境界でのUTF-8マルチバイト文字分断を防ぐため、MemoryStream で全チャンク結合後に一括デコードする方式へ変更。 |
| 4 | Lib-WebDebug_v303.ps1 | Convert-DomSnapshotFlat | バグ修正 | 複数フレーム（iframe）存在時のテキストノード結合において、フレーム開始オフセット（docStartIndex）を加算して親要素を特定するよう修正。 |
| 5 | Ps_Engine_Core_v303.ps1 | Set-ActiveTab<br>List-Tabs | 安定化 | Set-ActiveTab 実行時に WebViewCtrl 参照を同期。また List-Tabs でタブ数1件時もJSON配列形式（[{...}]）を維持する分岐を追加。 |
| 6 | Lib-DevTools_v303.ps1 | Get-WebSelectorFromUiaFuzzy | 精度向上 | Fuzzy検索のレーベンシュタイン距離算出前に比較文字列を .ToLower() で正規化し、大文字・小文字の差異によるスコア低下を抑止。 |

#### 1. Lib-WebCDP_v303.ps1：JSON-RPCのレスポンス受信処理
【検証詳細】
* **問題の妥当性（◎）：** 現行コード（257〜259行目）では、16KB（16384 バイト）の受信チャンクごとに UTF8.GetString を実行して文字列連結しています。日本語（UTF-8で3バイト）を含む16KB超の応答（DOMSnapshot.captureSnapshot 等）において、16,384バイトの境界でマルチバイト文字が分断されると U+FFFD（文字化け）が発生し、JSONデコード失敗や文字列破損の原因となります。したがって「全チャンクをバイト配列として蓄積し、EndOfMessage 後に1回だけUTF-8デコードする」という設計変更は必須です。 
* **ご提示コードの課題（⚠️ for ループによるパフォーマンス低下）：** ご提示のコードでは for ($i = 0; $i -lt $result.Count; $i++) { $allBytes.Add($receiveBuffer[$i]) } と1バイトずつループで追加しています。 PowerShell 5.1 のスクリプト for ループはオーバーヘッドが大きく、たとえば Export-WebDomSnapshot で 500KB（約50万バイト）〜1MB の応答を受信した場合、50万〜100万回のループ処理に数秒間を要し、241行目のタイムアウト監視（TimeoutSecPerTry = 4 秒）に抵触して通信エラーとなるリスクがあります。 
* **改善策：**
System.IO.MemoryStream の .Write($receiveBuffer, 0, $result.Count) を使用します。.Write() は .NET 内部でネイティブメモリコピー（Buffer.BlockCopy）を実行するため、PowerShell の for ループを一切回さず、1MBのデータでも1ミリ秒未満で処理が完了します。
【推奨修正コード（Lib-WebCDP_v303.ps1 230〜264行目付近）】

```PowerShell
            while ($true) {
                $responseString = ""
                # UTF-8マルチバイト文字のチャンク分断を防ぎつつ高速に結合するため、MemoryStreamを使用する。
                $ms = New-Object System.IO.MemoryStream

                try {
                    while ($true) {
                        # 非同期受信タスク(ReceiveAsync)を開始する。
                        $receiveTask = $global:CdpWebSocket.ReceiveAsync($receiveSegment, $cts.Token)

                        # 受信完了までのUIフリーズ防止およびタイムアウト監視： UIのフリーズを防ぎつつ、タイムアウトを監視する。
                        while (-not $receiveTask.IsCompleted) {
                            [System.Windows.Forms.Application]::DoEvents()
                            Start-Sleep -Milliseconds 5
                            if ($cdpSw.Elapsed.TotalSeconds -gt $TimeoutSecPerTry) {
                                throw [System.TimeoutException]::new("CDP_TIMEOUT")
                            }
                        }

                        if ($receiveTask.IsFaulted -or $receiveTask.IsCanceled) {
                            throw [System.Exception]::new("TASK_ERROR")
                        }

                        $result = $receiveTask.Result

                        if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
                            throw [System.Exception]::new("WEBSOCKET_CLOSED")
                        }

                        # 受信したバイト列をチャンク単位で一括書き込みする（forループを回避し高速化）。
                        if ($result.Count -gt 0) {
                            $ms.Write($receiveBuffer, 0, $result.Count)
                        }

                        # 分割送信の終端に達したら受信ループを抜ける。
                        if ($result.EndOfMessage) { break }
                    }

                    # ループを抜けた後に一度だけUTF-8デコードを行う。
                    $responseString = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
                } finally {
                    $ms.Dispose()
                }
```

#### 2. Lib-WebDebug_v303.ps1：Convert-DomSnapshotFlat の iframe 内テキストノード結合オフセット漏れ
【検証詳細】
* **問題の妥当性（◎ 重要バグ修正）：** CDPの DOMSnapshot.captureSnapshot が返す documents 配列は、0番目がトップ画面、1番目以降が子フレーム（iframe）であり、各ドキュメント内の parentIndex はそのフレーム内の先頭を 0 とする相対インデックスです。 現行コードの後半（770行目のレイアウト紐づけ）では $absoluteIdx = ($list.Count - $count) + $nodeIdx とオフセット計算を行っている一方、前半（742〜743行目のテキストノード結合）ではオフセットを足さずに $list[$node.parentIndex] を参照していました。 このため、iframe 内のテキストノードがトップ画面側の無関係な要素の innerText に混入し、iframe 内の要素自身の innerText が欠落するバグとなっていました。 
* **ご提示コードの評価と微調整：** $docStartIndex = $list.Count を記録して加算するロジックは完全に正確です。 1点だけ、$node.parentIndex が $null や -1（ルート要素）の場合に不要な加算を行わないよう、先に $null -ne $node.parentIndex -and $node.parentIndex -ge 0 を判定してから $absoluteParentIdx を算出する順序にし、あわせて770行目の ($list.Count - $count) も $docStartIndex に統一するとコード全体がより明快になります。 
【推奨修正コード（Lib-WebDebug_v303.ps1 666〜772行目付近）】

```PowerShell
        # ノード配列の要素数を特定する。
        $count = 0
        if ($nodes.parentIndex) { $count = $nodes.parentIndex.Count }
        elseif ($nodes.nodeName) { $count = $nodes.nodeName.Count }

        # 現在のドキュメント（iframe）が全体リストのどこから始まるかのオフセットを記録する。
        $docStartIndex = $list.Count

        # ----------------------------------------------------------
        # 1. DOM構造の展開： CDPの圧縮配列を展開し、要素ごとのオブジェクトを生成する。
        for ($i = 0; $i -lt $count; $i++) {
            # ... (中略: 675〜734行目) ...

            # テキストノード(nodeType=3)の抽出と親要素への結合
            if ($node.nodeType -eq 3 -and $nodes.nodeValue -and $null -ne $nodes.nodeValue[$i]) {
                $valIndex = $nodes.nodeValue[$i]
                if ($valIndex -ge 0 -and $valIndex -lt $strings.Count) {
                    $textVal = $strings[$valIndex].Trim()
                    if (-not [string]::IsNullOrWhiteSpace($textVal)) {
                        # 親ノード(AタグやSPAN等)へテキストを連結する（フレーム毎のオフセットを加算）。
                        if ($null -ne $node.parentIndex -and $node.parentIndex -ge 0) {
                            $absoluteParentIdx = $docStartIndex + $node.parentIndex
                            if ($absoluteParentIdx -lt $list.Count) {
                                $parentNode = $list[$absoluteParentIdx]
                                if (-not $parentNode.innerText) {
                                    $parentNode.innerText = $textVal
                                } else {
                                    $parentNode.innerText += " " + $textVal
                                }
                            }
                        }
                    }
                }
            }
            $list.Add($node)
        }
```

(※770行目の $absoluteIdx = ($list.Count - $count) + $nodeIdx も $absoluteIdx = $docStartIndex + $nodeIdx と書くと統一できますが、計算結果は同じためそのままでも動作します。) 

#### 3. Ps_Engine_Core_v303.ps1：List-Tabs の JSON配列アンラップ防止
【検証詳細】
* **現行コードの挙動確認：** PowerShell 5.1 で要素数1の配列が {...}（スカラー）にアンラップされるのはパイプライン渡し（$result | ConvertTo-Json）を行った場合であり、現行コード（450行目）の ConvertTo-Json -InputObject @($result) -Depth 2 -Compress は、実行ログ Engine_SystemLogテスト_一般的な.txt（76〜77行目）の通りタブ数1件でも [{"TabId":"P01",...}] と配列形式で出力されています。 
* **ご提示コードの評価（○ 防御的実装として有効）：** VBA側の UserForm_Picker.frm（139〜141行目）では、JsonConverter.ParseJson(resJson) の戻り値を For Each item In jsonParsed で回しており、応答文字列が [ で始まるJSON配列（VBA-JSONにおける Collection）であることが必須条件となっています。 将来的なPowerShellバージョンの違い（PS7等）や型キャストの揺れに左右されず、1件時に文字列レベルで "[$($result[0] | ConvertTo-Json -Depth 2 -Compress)]" として [{...}] を保証するご提示の書き方は、VBA連携の安全策として有効です。 
* **微調整のポイント：** else 分岐側でも、現行通り @($result) として配列キャストを維持し、かつ万が一タブが0件（$result.Count -eq 0）になった場合でも空配列 "[]" が返るようにしておくと完璧です。 
【推奨修正コード（Ps_Engine_Core_v303.ps1 448〜453行目付近）】

```PowerShell
    Write-DebugLog -Message "[List-Tabs] タブ数: $($result.Count)" -Level Info
    # 要素が1つの場合でも確実にJSON配列([{...}])としてVBA側(JsonConverter)へ返却する。
    if ($result.Count -eq 0) {
        $resultJson = "[]"
    } elseif ($result.Count -eq 1) {
        $resultJson = "[$($result[0] | ConvertTo-Json -Depth 2 -Compress)]"
    } else {
        $resultJson = ConvertTo-Json -InputObject @($result) -Depth 2 -Compress
    }
    Write-DebugLog -Message "$resultJson" -Level Info
    return $resultJson
```

#### 4. Lib-DevTools_v303.ps1：Get-WebSelectorFromUiaFuzzy のレーベンシュタイン距離・小文字正規化
【検証詳細】
* **妥当性と副作用の検証（◎ 採用推奨）：**
  * Convert-DomSnapshotFlat（Lib-WebDebug_v303.ps1 701行目）では tagName が .ToLower() で小文字化されている一方、Process-RecordAction（Lib-DevTools_v303.ps1 935行目）から渡される UiaType は Button や Edit のように大文字を含む文字列となっています。 
  * PowerShell 5.1 の -eq 演算子（169行目 $s[$i] -eq $t[$j]）は文字型に対してもデフォルトで大文字・小文字を同一視しますが、比較前のベース文字列（uiaText および domText）の時点で明示的に .ToLower() で正規化しておく方が、言語仕様の暗黙挙動に依存せず確実です。 
  * また、uiaText と domText はレーベンシュタイン距離の計算（219行目）のみに使用され、最終的なセレクタ生成（234〜236行目の $bestNode.selector, $bestNode.name, $bestNode.id）には元のケースが保持されたプロパティが使用されるため、.ToLower() によって生成セレクタの大文字・小文字が変わってしまう副作用も一切ありません。 
【推奨修正コード（Lib-DevTools_v303.ps1 183〜195行目付近）】

```PowerShell
    # UIA側の情報から比較用のベース文字列を生成する（大文字・小文字の違いによる差異を防ぐため小文字へ正規化）。
    $uiaText = ("$UiaName $UiaHelpText $UiaId $UiaType").Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($uiaText)) { return $null }
    
    # 探索対象を操作可能な要素に絞り込み、計算コストを削減する。
    $interactiveTags = @("input", "button", "a", "select", "textarea", "label", "span", "div", "img")

    foreach ($node in $domNodes) {
        if (-not $node.tagName -or $node.tagName -notin $interactiveTags) { continue }
        # DOM側の情報からも比較用のベース文字列を生成する（小文字へ正規化）。
        $domText = "$($node.placeholder) $($node.name) $($node.id) $($node.tagName) $($node.innerText)".Trim().ToLower()
        if (-not $domText) { continue }
```

(※万一チャット画面の表示フィルターによりコード内の演算子直後の半角スペースが詰まって見える箇所があれば、-eq $msgId や -lt $count のように演算子直後に半角スペースを1つ入れて適用してください。)

---
### Step5: 時間は掛かるが、分析・総合力は Argon へ。3.1Proが脱帽

ご提示いただいた修正コードを再確認しました。すべてのコードが完全に正しく、かつ私が提案した元の修正案よりもさらに洗練された実装になっています。<br>
特に以下の2点において、非常に優れたコードへのブラッシュアップが行われています。<br>

1. MemoryStream の採用（Lib-WebCDP_v303.ps1）
私が提示した List[byte] と for ループによる1バイトずつの追加処理よりも、MemoryStream の Write メソッドを利用してチャンクを一括で書き込む方が、内部的にOSレベルのブロックコピーが使われるため、メモリ効率・処理速度ともに圧倒的に優れています。try-finally による Dispose() の確実なリソース解放も完璧です。
2. Count -eq 0 の明示的ハンドリング（Ps_Engine_Core_v303.ps1）
タブ数がゼロになるケース（全タブクラッシュ時など）に対して、PowerShellの ConvertTo-Json の挙動に依存せず、明示的に空のJSON配列 [] を返す処理を追加された点は、VBA側のパースエラーを未然に防ぐ素晴らしい防衛的プログラミング（フェイルセーフ）です。
残り2つの修正である「Convert-DomSnapshotFlat のインデックスオフセット計算のネスト整理」および「Fuzzy検索の .ToLower() による正規化」についても、ロジック・文法ともに全く問題ありません。

<sub>このコード群であれば、長時間の高負荷稼働やマルチバイト通信のエッジケースに対しても非常に強く、GitHubやQiitaで公開するOSS（オープンソース）のRPAエンジン基盤として、第三者が安心して利用・拡張できる最高品質の仕上がりです。このまま最終版のコードとして適用してください。</sub>

---
😊😊😊
