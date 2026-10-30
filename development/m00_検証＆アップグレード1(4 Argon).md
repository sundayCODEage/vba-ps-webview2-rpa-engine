
<h4 align="center">汎用RPA操作エンジン (VBA × PowerShell × WebView2 Hybrid-Engine)</h4>
<h4 align="center">1 / 2</h4>
<h4 align="center">V301 ⇒ V302 （Gemini 4 Argon／【Google AI Pro】）</h4>

<p align="right">
  作成日: 2026年10月16日 - 10月22日<br>
  作成者: System beginner & AI Collaborator (Gemini)
</p>

_GEMINI 3.1Proで開発したコードを、**Gemini 4 Argonで全モジュールを再検証・アップグレード**しました。_<br>
_<sub>（Argon凄い！。コード開発は最低限の課金（今回: Pro、多分1ヵ月）は必要でした。今までの時間が○○！ でも、普段は無料枠でも十分な私）</sub>_

_（アップグレード記録）_<br>
_<sub>コード改修の経過を中心に纏めているので、GEMINIへの指示、アドバイスは要約しています。</sub>_

* _Step1: 全モジュール、VBA、マニュアル等に対し、「**コード、コメントの検証**」を指示。_
> _<sub>Ps_Engine_Core_v301.ps1 冒頭の「命名規則およびコーディングルール」「コメント規約（体言止め・常体・大ブロック[]の階層制限・既存コメント削除禁止）」を厳格に適用し、全モジュールの精査をお願い。 </sub>_

* _Step2: 修正コードの検証コード(主な改修点を検証する)_
* _Step3: コード修正（微・調整）_

```Diff
* _Step4: モジュールのアップグレード要望_<br>
  * 課題①： ピッカー実行時間（10秒〜20秒）を 短縮したい。
  * 課題②： Shadow DOM 内の XPATH(生成)と Test-SelectorStability の不一致を解消したい
  * 課題③： iframe内要素の 絶対座標（BoundingBox）のズレ解消
* _Step5: コード修正（微・調整）_<br>
* _Step6: コード修正（再・調整）_<br>
  * 課題①： 1レコードが上下2段やセル結合への対応と、展開オプションの付加
  * 課題②： リアルタイム・ハイライトのiframe・Shadow DOM透過対応と、PSモジュールへ移行
* _Step7: コード修正（微・調整）_<br>
  * 課題①： Shadow DOM 探査の強化（deepElementFromPoint の再帰探索、座標ズレ、すり替えロジック）
  * 課題②： Shadow DOM内では、XPathに逃げず、階層構造のCSSセレクタを完全生成する
```

---
### ◆◆ Step1 ◆◆
<h4 align="center">PowerShellモジュールの検証( 修正 )</h4>

### 1. Ps_Engine_Core_v302.ps1（司令塔・ルーター）
* **修正点1（コメント規約）：** 冒頭ルールのタイポ（• 大ブロックの解釈】 等）の整形。 
* **修正点2（命名規則）：** Set-EngineConfig の真偽値文字列パラメータに Flg 接尾辞が漏れている点（または後方互換エイリアス）の補強と、大ブロック [ ] が finally 内に使われている箇所の調整。 
* **修正点3（機能）：** Switch-Tab 内の Connect-CdpSession がコメントアウトされていたため、タブ切替時に即座にセッションまたは ConnectedCdpTabId が更新されるよう調整。

```Diff
@@ 95行目付近：コメント規約ブロックのタイポ修正 @@
-• 大ブロックの解釈】 括弧 [ ] を用いた大ブロック指定は、ルートレベル（モジュール読み込みやメインループ等、大きな処理の塊）にのみ適用する。
- 【関数内、ループ内、イベントハンドラ内の個別処理からは [ ] を外し、通常のインラインコメントに降格させることで視覚的な階層を明確にする。
+• 【大ブロックの解釈】 括弧 [ ] を用いた大ブロック指定は、ルートレベル（モジュール読み込みやメインループ等、大きな処理の塊）にのみ適用する。
+  関数内、ループ内、イベントハンドラ内の個別処理からは [ ] を外し、通常のインラインコメントに降格させることで視覚的な階層を明確にする。
```

```Diff
@@ 415行目付近：Switch-Tab 内のCDP再接続制御 @@
     if ($global:CdpPort -gt 0) {
         try {
-            # Connect-CdpSession -Port $global:CdpPort
+            # タブ切替に伴うCDPターゲット同期： 次回コマンド実行時に確実に再接続させるため接続済みタブIDをリセットする。
+            $global:ConnectedCdpTabId = $null
+            Connect-CdpSession -Port $global:CdpPort
         } catch {
             Write-DebugLog -Message "[Switch-Tab] CDP再接続失敗: $($_.Exception.Message)" -Level Warning
         }
     }
```

```Diff
@@ 495行目付近：Set-EngineConfig の命名規則準拠（後方互換Alias付き） @@
 function Set-EngineConfig {
     # ハイライト機能設定： VBA側からのグローバル設定オブジェクトを更新する。
-    param ([string]$EnableHighlight)
+    param (
+        [Alias("EnableHighlight")]
+        [string]$EnableHighlightFlg
+    )
 
     $func = $MyInvocation.MyCommand.Name
 
     # EnableHighlightFlgをBoolean変換して更新する。
-    if (-not [string]::IsNullOrEmpty($EnableHighlight)) {
-        $global:CONFIG.EnableHighlight = [System.Convert]::ToBoolean($EnableHighlight)
+    if (-not [string]::IsNullOrEmpty($EnableHighlightFlg)) {
+        $global:CONFIG.EnableHighlight = [System.Convert]::ToBoolean($EnableHighlightFlg)
 #●        Write-DebugLog -Message "[$func] 設定変更: EnableHighlight = $($global:CONFIG.EnableHighlight)" -Level Info
     }
     return "Config Updated"
 }
```

```Diff
@@ 625行目付近：finallyブロック内の大ブロック括弧 [ ] の降格（規約準拠） @@
 } finally {
     # ==============================================================================
-    # [安全な終了処理（クリーンアップ）]： フォームやプロセスを破棄し、メモリリソースを解放する。
+    # 安全な終了処理（クリーンアップ）： フォームやプロセスを破棄し、メモリリソースを解放する。
     # ==============================================================================
```

### 2. Lib-WebJS_v302.ps1（共通ユーティリティ）
*	**修正点（コメント見直し）：** isVisible や safeClick 内の重複コメント（242〜243行目）を整理し、トレーサーメッセージの文体を統一。

```Diff
@@ 240行目付近：重複コメントの整理と統合 @@
-    // 最終手段（絶対パスによる強制生成）： 上記すべてで特定できない場合、ルートからの完全な階層パスを構築する。
-    // 上記のすべてで特定できなかった場合、DOMツリーのルートからの完全な階層（nth-child等）を構築する。
+    // 最終手段（絶対パスによる強制生成）： 上記のすべてで特定できない場合、DOMツリーのルートからの完全な階層パス（nth-child等）を構築する。
     // 探索の起点を元の要素にリセットする。
     node = el;
```

### 3. Lib-WebView2_Init_v302.ps1（ブラウザ画面起動）
*	**修正点（バグ防止・機能補強）：** 初回タブ P01 には TargetId が記録されておらず（ポップアップ時のみ記録）、本日の実行ログでも [Connect-CdpSession] TargetId 未設定 → URL マッチにフォールバック が発生していました。また、ポップアップ生成時（$newWebView）に Url キーの初期値が存在しないため、初期化時に空文字で定義します。

```Diff
@@ 185行目付近：仮想タブ初期化時のUrlキー初期定義 @@
                         # 新しい仮想タブに一意のID(GUID)を割り当て、親画面との関係性をグローバル辞書に記録する。
                         $tabId = [guid]::NewGuid().ToString()
                         $global:Tabs[$tabId] = @{
                             WebView = $senderInit
                             Parent  = $parentCtrl
+                            Url     = ""
                         }
                         $global:ActiveTabId = $tabId
```

### 4. Lib-WebView2_Native_v302.ps1（ネイティブ通信）
*	**修正点（コード整形）：** 132行目のスペース抜け（if ($attempt -lt$Retries)）の修正と、既に New-EngineException で整形されたエラー（JSエラー等）が catch 節でリトライ・二重ラップされてしまう問題の回避。

```Diff
@@ 130行目付近：JSエラーの二重ラップ防止およびスペース修正 @@
         } catch {
+            # 既にJSエラーとして確定している場合は不要なリトライを行わずそのまま再送出する。
+            if ($_.Exception.Message -match "^\[.*?\] \[JSエラー\]:") {
+                throw $_
+            }
             # 実行時エラーに対する全体リトライ処理： 予期せぬ実行エラー発生時も上限回数までリトライを試行する。
-            if ($attempt -lt$Retries) {
+            if ($attempt -lt $Retries) {
                 Start-Sleep -Milliseconds $delayMs
                 $delayMs = [Math]::Min(2000, $delayMs * 2)
                 continue
             } else {
```

### 5. Lib-WebCDP_v302.ps1（WebSocket・CDP高速通信）
* **修正点1（正規表現バグ修正）：** Invoke-CdpScript 315行目の文字列アンエスケープ処理において、Lib-WebView2_Native_v302.ps1 側には入っていた単一行モード (?s) が漏れており、複数行テキスト（HTMLやテーブル等）をCDP経由で取得した際にアンエスケープが失敗する不具合を修正します。 
* **修正点2（命名規則）：** グローバル変数 $global:ConnectedCdpTabId をファイル冒頭で明示的に初期化し、$matches を $Matches に統一します。

```Diff
@@ 8行目付近：グローバル変数の明示的初期化 @@
-# グローバル変数の定義
+# [CDP用グローバル変数の初期化]： WebSocketセッションおよびメッセージID管理用の変数を定義する。
 $global:CdpWebSocket = $null
 $global:CdpMessageId = 0
 $global:CdpEndpoint  = ""
+$global:ConnectedCdpTabId = $null
```

```Diff
@@ 315行目付近：Invoke-CdpScript の複数行文字列アンエスケープ対応 @@
     # 文字列JSONのアンエスケープ： 戻り値の文字列から不要なエスケープを解除する。
-    if ($rawResult -match '^\s*"(.*)"\s*$') {
-        $inner = $matches[1]
+    # 取得テキストに改行が含まれる場合を考慮し、正規表現に単一行モード（(?s)）を付与する。
+    if ($rawResult -match '(?s)^\s*"(.*)"\s*$') {
+        $inner = $Matches[1]
         try {
             $rawResult = [System.Text.RegularExpressions.Regex]::Unescape($inner)
         } catch {
             $rawResult = $inner
         }
     }
```

### 6. Lib-WebSafeAction_v302.ps1（フェイルセーフ・安全クリック）
* **修正点（コメント規約・命名規則）：** 18行目の $matches[1] を $Matches[1] に統一し、134行目の句点抜けコメントを常体（〜する。）に修正します。

```Diff
@@ 16行目付近 @@
     if ($XPath -match "contains\(\.,\s*['`"]([^'`"]+)['`"]\)") {
-        $searchWord = $matches[1]
+        $searchWord = $Matches[1]
     }
```

```Diff
@@ 132行目付近 @@
         try {
             $resObj =$resJson | ConvertFrom-Json -ErrorAction Stop
         } catch {
-            # JSONパース失敗時は待機ループを継続させるため $false を返す
+            # JSONパース失敗時は待機ループを継続させるため $false を返す。
             return $false
         }
```

### 7. Lib-WebXPath_v302.ps1（XPath特殊操作）
* **修正点1（致命的エラー防止）：** 192行目と212行目で Write-DebugLog ... -Level Warn と記述されていますが、Write-DebugLog の [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Fatal')] では 'Warn' が許容されておらず、フォールバック発動時にPowerShell側でパラメータ検証例外が発生してしまいます。これを 'Warning' に修正します。 
* **修正点2（コーディングルール準拠）：** 140行目の if ($exists -eq $false -or $exists -match "false") を、仕様書(第4.4章・第8.4章)の「文字列としての保険的な比較は書かず、保証された型を直接評価する」ルールに従い if (-not $exists) に統一します。 
* **修正点3（引数整合性）：** ピッカー（ConvertTo-RpaCommandCode）が Set-WebXPathTextInput を生成する際、-OuterHtml と -BoundingBox を付与しますが、Set-WebXPathTextInput 側にその受け皿パラメータが無いため、実行時にパラメータエラーになる問題を修正します。

```Diff
@@ 138行目付近：Wait-WebXPathElementDisappear の型評価純粋化 @@
     while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
         # DOM遷移中等の通信切断を避けるため、CDPではなくネイティブAPIを用いて判定を繰り返す。
         $exists = Invoke-WebView2NativeScript -Js $js -Retries 1
         # 戻り値が false (非表示/消滅) となった場合はループを抜ける。
-        if ($exists -eq $false -or $exists -match "false") {
+        if (-not $exists) {
             $isGone = $true
             break
         }
         Start-Sleep -Milliseconds 500
     }
```

```Diff
@@ 190行目付近：Invoke-WebXPathClick の ValidateSet エラー修正 @@
     $res = Invoke-WebScript -Js $js
     
-    if (-not $res) { Write-DebugLog -Message "[$func] 警告: JSによるXPath要素のクリックが失敗しました。" -Level Warn }
+    if (-not $res) { Write-DebugLog -Message "[$func] 警告: JSによるXPath要素のクリックが失敗しました。" -Level Warning }
 
     # JSクリック失敗時の物理クリックフォールバック： ピッカーからの座標(BoundingBox)指定がある場合はOS物理クリックへ移行する。
     if (-not [string]::IsNullOrWhiteSpace($BoundingBox)) {
         try {
             # JSON形式のBoundingBoxをPowerShellオブジェクトに変換する。
             $bbox = $BoundingBox | ConvertFrom-Json
              if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
                 # 論理座標をOSの絶対スクリーン座標へ変換する。
                 $screenPt = Convert-BoundingBoxToScreenPoint -BoundingBox $bbox
                 $targetX = $screenPt.X
                 $targetY = $screenPt.Y
 
                 Write-DebugLog -Message "[$func] 情報: BoundingBoxを使用し、座標(X:$targetX, Y:$targetY)へ物理クリックを試行します。" -Level Info
                 # DesktopUIAモジュール(Win32API)の物理クリック関数がロードされているか確認する。
                 if (Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) {
                     Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
                     return $true
                 } else {
-                    Write-DebugLog -Message "[$func] 警告: 物理クリック関数が見つかりません。Lib-DesktopUIAとの連携を確認してください。" -Level Warn
+                    Write-DebugLog -Message "[$func] 警告: 物理クリック関数が見つかりません。Lib-DesktopUIAとの連携を確認してください。" -Level Warning
                 }
             }
         } catch {
             Write-DebugLog -Message "[$func] エラー: BoundingBoxの解析または物理クリック中に例外が発生しました。" -Level Error
         }
     }
```

```Diff
@@ 225行目付近：Set-WebXPathTextInput のピッカー生成パラメータ受け皿追加 @@
 function Set-WebXPathTextInput {
     param (
         [Parameter(Mandatory = $true)][string]$XPath,
         [Parameter(Mandatory = $true)][string]$Value,
-        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
+        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
+        [string]$OuterHtml = "",
+        [string]$BoundingBox = ""
     )
```

### 8. Lib-WebAction_v302.ps1（Web標準操作）
* **修正点1（コーディングルール準拠）：** Wait-WebDocumentReady（112行目）と Wait-WebScreenUnlock（339行目）に残っていた -or $state -eq "true" 等の文字列比較を排除し、純粋な型評価に統一します。 
* **修正点2（JSモダン化）：** Wait-WebPageLoad（63行目）の let frames を再代入不要のため const frames に統一。 
* **修正点3（仕様統一）：** Get-WebAttribute（602行目）が document.querySelector 直叩き＆ Invoke-WebView2NativeScript 固定になっていたため、他の関数と同様に utilFindInFrames / deepQuerySelector および Normalize-XPath、Invoke-WebScript ルーターを経由する堅牢な実装へ引き上げます。

```Diff
@@ 60行目付近：Wait-WebPageLoad 内の const 統一 @@
                         // iframeのDOMがまだ構築中のケースに対応する。
-                        let frames = win.frames;
+                        const frames = win.frames;
                         for (let i = 0; i < frames.length; i++) {
```

```Diff
@@ 110行目付近：Wait-WebDocumentReady の型評価純粋化 @@
     while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
         # 修正後のラッパー構造を介して安全にBoolean（true/false）を受け取る。
         $state = Invoke-WebScript -Js $js -Retries 1
         
-        if ($state -eq $true -or $state -eq "true") {
+        if ($state) {
             $isReady = $true
             break
         }
```

```Diff
@@ 335行目付近：Wait-WebScreenUnlock の型評価純粋化 @@
-            # 文字列・Bool値の判定ロジック：
-            $blockedBool = ($isBlocked -eq $true -or $isBlocked -eq "true")
+            # 真偽値の判定ロジック： ラッパーにより保証されたBoolean型として直接評価する。
+            $blockedBool = [bool]$isBlocked
```

```Diff
@@ 585行目付近：Get-WebAttribute の多段iframe/ShadowDOM対応および型評価純粋化 @@
         # 2. Selector または XPath に応じた待機処理とJS組み立て：
         if (-not [string]::IsNullOrEmpty($Selector)) {
             # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ。
             $null = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec
             
             $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
             # セレクタで特定した要素の指定属性を取得して返却するJSを構築する。
-            $js = "const sel = decodeURIComponent(escape(atob('$selB64'))); const attr = decodeURIComponent(escape(atob('$attrB64'))); return document.querySelector(sel).getAttribute(attr);"
+            $js = @"
+                const sel = decodeURIComponent(escape(atob('$selB64')));
+                const attr = decodeURIComponent(escape(atob('$attrB64')));
+                return utilFindInFrames(window, function(win) {
+                    try {
+                        const el = deepQuerySelector(sel, win.document);
+                        return el ? el.getAttribute(attr) : null;
+                    } catch(e) { return null; }
+                });
+"@
         }
         elseif (-not [string]::IsNullOrEmpty($XPath)) {
             # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ。
             $null = Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec
             
-            $xpathB64 = ConvertTo-JsSafeBase64 -Text $XPath
+            $xpathNorm = Normalize-XPath -XPath $XPath
+            $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm
             # XPathで特定した要素の指定属性を取得して返却するJSを構築する。
-            $js = "const xpath = decodeURIComponent(escape(atob('$xpathB64'))); const attr = decodeURIComponent(escape(atob('$attrB64'))); const node = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; return node ? node.getAttribute(attr) : null;"
+            $js = @"
+                const xpath = decodeURIComponent(escape(atob('$xpathB64')));
+                const attr = decodeURIComponent(escape(atob('$attrB64')));
+                return utilFindInFrames(window, function(win) {
+                    try {
+                        const node = win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
+                        return node ? node.getAttribute(attr) : null;
+                    } catch(e) { return null; }
+                });
+"@
         }
 
         # 3. JSの実行： 共通ルーター関数を利用してブラウザに命令する。
-        $result = Invoke-WebView2NativeScript -Js $js
+        $result = Invoke-WebScript -Js $js
         
         # 取得結果が null の場合は空文字を返す。
-        if ($result -eq "null" -or $null -eq $result) {
+        if ($null -eq $result) {
             return ""
         }
         return $result
```

### 9. Lib-DesktopUIA_v302.ps1（デスクトップ操作・UIA）
* **修正点（コメント見直し）：** 命名規則・例外フォーマットともに非常にクリーンに実装されています。コメントの体言止め・常体の微調整のみ行います。

```Diff
@@ 95行目付近：コメントの常体統一 @@
                 $wshell = New-Object -ComObject WScript.Shell
                 # 既存の入力値を確実に取り除くため、全選択して削除してから送信する。
-                $wshell.SendKeys("^a") # 全選択
+                $wshell.SendKeys("^a") # 全選択を実行する。
                 Start-Sleep -Milliseconds 50
-                $wshell.SendKeys("{DELETE}") # 削除
+                $wshell.SendKeys("{DELETE}") # 削除を実行する。
                 Start-Sleep -Milliseconds 50
-                $wshell.SendKeys($Value) # 入力
+                $wshell.SendKeys($Value) # 値を入力する。
```

### 10. Lib-DevTools_v302.ps1（UIA×WebView2 要素ピッカー）
* **修正点1（致命的エラー防止）：** 65行目の Write-DebugLog ... -Level Warn を -Level Warning に修正（ValidateSet エラー回避）。 
* **修正点2（例外種別のValidateSet準拠）：** Get-UiaElementFromPoint（"GetElementError"）、Get-UiaWindowFromElement（"GetWindowError"）、Get-UiaTargetInfoFromCursor（"GetTargetInfoError"）、Build-RpaStepCode（"CodeBuildError"）、Test-WebPointReverse（"テスト実行エラー"）において、New-EngineException の -Type に未定義の文字列が渡されており、例外発生時に New-EngineException 自体がクラッシュする問題を修正します（標準定義の "UIAエラー" / "内部エラー" へ統一）。 
* **修正点3（命名規則・タイポ）：** Get-LevenshteinDistanceFast のパラメータ $s, $t を PascalCase（$SourceText, $TargetText）または内部ローカル変数へ整理し、コメントの「境界短形」を「境界矩形」へ修正します。

```Diff
@@ 18行目付近：タイポ修正（短形 → 矩形） @@
-    # 1. UIA (OSレベル) の絶対座標計算： UIA要素のOS画面上での境界短形(BoundingRectangle)を取得する。
-    # UIA要素のOS画面上での境界短形(BoundingRectangle)を取得する。
+    # 1. UIA (OSレベル) の絶対座標計算： UIA要素のOS画面上での境界矩形(BoundingRectangle)を取得する。
     $rect = $UiaElement.Current.BoundingRectangle
```

```Diff
@@ 63行目付近：ValidateSet エラー修正 (Warn -> Warning) @@
     $domNodes = $global:DomNodes
     if (-not $domNodes -or $domNodes.Count -eq 0) {
-        Write-DebugLog -Message "[$func] 警告: 比較用のDOMSnapshotが存在しません" -Level Warn
+        Write-DebugLog -Message "[$func] 警告: 比較用のDOMSnapshotが存在しません" -Level Warning
         return $null
     }
```

```Diff
@@ 135行目付近：Get-LevenshteinDistanceFast のパラメータ命名規則準拠 @@
     # レーベンシュタイン距離の高速計算： 1次元配列を使い回してメモリ割り当てを抑制する。
     function Get-LevenshteinDistanceFast {
-        param([string]$s, [string]$t)
+        param([string]$SourceStr, [string]$TargetStr)
+        $s = $SourceStr
+        $t = $TargetStr
```

```Diff
@@ 205行目付近：Get-LevenshteinDistanceFast 呼び出し箇所の引数名更新 @@
-        $dist = Get-LevenshteinDistanceFast -s $uiaText -t $domText
+        $dist = Get-LevenshteinDistanceFast -SourceStr $uiaText -TargetStr $domText
```

```Diff
@@ 700行目付近：Get-UiaElementFromPoint の New-EngineException Type 修正 @@
     } catch {
-        throw (New-EngineException -Func $func -Type "GetElementError" -Message "座標からのUIA要素取得に失敗しました" -Details $_.Exception.Message)
+        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "座標からのUIA要素取得に失敗しました" -Details $_.Exception.Message)
     }
```

```Diff
@@ 720行目付近：Get-UiaWindowFromElement の New-EngineException Type 修正 @@
     } catch {
-        throw (New-EngineException -Func $func -Type "GetWindowError" -Message "親ウィンドウの特定に失敗しました" -Details $_.Exception.Message)
+        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "親ウィンドウの特定に失敗しました" -Details $_.Exception.Message)
     }
```

```Diff
@@ 805行目付近：Get-UiaTargetInfoFromCursor の New-EngineException Type 修正 @@
     } catch {
-        throw (New-EngineException -Func $func -Type "GetTargetInfoError" -Message "UIA探索中にエラー" -Details $_.Exception.Message)
+        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "UIA探索中にエラー" -Details $_.Exception.Message)
     }
```

```Diff
@@ 840行目付近：Build-RpaStepCode の New-EngineException Type 修正 @@
     } catch {
-        throw (New-EngineException -Func $func -Type "CodeBuildError" -Message "ステップコードの生成に失敗しました" -Details $_.Exception.Message)
+        throw (New-EngineException -Func $func -Type "内部エラー" -Message "ステップコードの生成に失敗しました" -Details $_.Exception.Message)
     }
```

```Diff
@@ 875行目付近：Test-WebPointReverse の New-EngineException Type 修正 @@
     } catch {
-        throw (New-EngineException -Func $func -Type "テスト実行エラー" -Message "テスト用JSの評価中にエラーが発生しました" -Details $_.Exception.Message)
+        throw (New-EngineException -Func $func -Type "JSエラー" -Message "テスト用JSの評価中にエラーが発生しました" -Details $_.Exception.Message)
     }
```

### 11. Lib-WebDebug_v302.ps1（HTML/CSV保存・スクショ・DOMSnapshot）
* **修正点1（命名規則）：** Export-WebDomSnapshot 内で Convert-DomSnapshotFlat -domResult と小文字指定になっている箇所を -DomResult に統一。 
* **修正点2（エラーメッセージの整合性）：** Export-WebLayoutDump の catch 節のエラーメッセージがコピー元の "DOM Snapshot の取得処理中に例外が発生しました" のままになっていたため、"LayoutDump の取得処理中に例外が発生しました" へ修正します。 
* **修正点3（Shadow DOM貫通対応）：** m13_コードテスト(要素ピッカー).md で課題として挙げられていた 「Test-SelectorStability が querySelector を使用しているため Shadow DOM 内の要素を検証できない問題」 について、注入済みの deepQuerySelector を使用するようアップグレードします。

```Diff
@@ 350行目付近：タイポ修正（短形 → 矩形） @@
-        # RPAブラウザフォーム(WinForms)の論理的な境界短形を取得する。
+        # RPAブラウザフォーム(WinForms)の論理的な境界矩形を取得する。
         $bounds = $global:BrowserForm.Bounds
```

```Diff
@@ 475行目付近：パラメータ名のPascalCase統一 @@
         # Convert-DomSnapshotFlat を適用して人間が読める形式へ変換： CDP特有の多次元・圧縮配列データを、解析しやすいフラットなオブジェクトリストへ展開する。
-        $flatNodes = Convert-DomSnapshotFlat -domResult $cdpResponse.result
+        $flatNodes = Convert-DomSnapshotFlat -DomResult $cdpResponse.result
```

```Diff
@@ 740行目付近：Export-WebLayoutDump の例外メッセージ修正 @@
     } catch {
         # すでに New-EngineException で生成されたエラーならそのままスローする。
         if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
         
         throw (New-EngineException -Func $func -Type "内部エラー" `
-            -Message "DOM Snapshot の取得処理中に例外が発生しました" -Details $_.Exception.Message)
+            -Message "LayoutDump の取得処理中に例外が発生しました" -Details $_.Exception.Message)
     }
```

```Diff
@@ 780行目付近：Test-SelectorStability の Shadow DOM 貫通対応 @@
             // 共通関数(utilFindInFrames)を利用し、iframe内も含めて要素を透過探索する。
             const el = utilFindInFrames(window, function(win) {
                 try {
-                    // プレフィックスがXPATH:の場合はXPath評価、それ以外はquerySelectorを実行する。
+                    // プレフィックスがXPATH:の場合はXPath評価、それ以外はShadow DOMも貫通するdeepQuerySelectorを実行する。
                     return sel.indexOf('XPATH:') === 0
                         ? win.document.evaluate(sel.substring(6), win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue
-                        : win.document.querySelector(sel);
+                        : deepQuerySelector(sel, win.document);
                 } catch(e) { return null; }
             });
```

---
<h4 align="center">VBAモジュールの検証( 修正 )</h4>

### 1. Ps_Bridge.bas（共通関数）
* **修正点1（コメント規約）：** 冒頭のコメント規約（体言止め・常体「〜する。」）に統一。 
* **修正点2（正規表現の括弧ネスト対策）：** ParseRpaError の正規表現 (?:\s*\((.*)\))?$ は、エラーメッセージ本文に ( ) が含まれている場合（例：要素の未出現または非表示 (#btn)）かつ Details が空のときに、本文の括弧を Details と誤認識して切り出してしまうケースがあるため、最後の括弧ペアを正確に捉えるようコメントと処理を微調整します。

```Diff
@@ 1行目付近：コメント見直し（体言止め・常体統一） @@
 Attribute VB_Name = "Ps_Bridge"
 Option Explicit
 
-' --- 特定のウィンドウを最前面に移動させる ---
+' --- 特定ウィンドウの最前面移動用API宣言 ---
 Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
 Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
 Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
 Private Const SW_RESTORE As Long = 9
     
-' --- 例外情報を格納する構造体 ---
+' --- 例外情報格納用構造体 ---
 Public Type RpaExceptionInfo
     HasError As Boolean
     FunctionName As String
     ErrorType As String
     Message As String
     Details As String
     RawText As String
 End Type
     
 ' --- 可変長引数(ParamArray)からのDictionary生成 ---
 Public Function CreateParams(ParamArray args() As Variant) As Scripting.Dictionary
     Dim dict As New Scripting.Dictionary
     Dim i As Integer
     
-    ' 引数のペア不一致判定
+    ' 引数のペア不一致判定： 奇数個の場合は空の辞書を返却する。
     If (UBound(args) - LBound(args) + 1) Mod 2 <> 0 Then
         Set CreateParams = dict
         Exit Function
     End If
     
-    ' キーと値のペア登録
+    ' キーと値のペア登録： 2つずつループして辞書へ追加する。
     For i = LBound(args) To UBound(args) - 1 Step 2
         dict.Add CStr(args(i)), args(i + 1)
     Next i
     
     Set CreateParams = dict
 End Function
 
-' --- PowerShellのエラー文字列を構造体にパースする関数 ---
+' --- PowerShellエラー文字列の構造体パース ---
 Public Function ParseRpaError(ByVal errString As String) As RpaExceptionInfo
     Dim result As RpaExceptionInfo
     result.HasError = False
     result.RawText = errString
     
-    ' 正規表現オブジェクトの生成（遅延バインディング）
+    ' 正規表現オブジェクトの生成： 遅延バインディングでVBScript.RegExpを生成する。
     Dim regEx As Object
     Set regEx = CreateObject("VBScript.RegExp")
     
-    ' パターン: [関数名] [種別]: メッセージ (詳細)
-    ' ※詳細は省略される場合があるため、末尾の括弧はオプショナル(?: \((.*)\))?とする
+    ' パターン定義（[関数名] [種別]: メッセージ (詳細)）： 詳細は省略される場合があるため、末尾の括弧はオプショナルとする。
     regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.*?)(?:\s*\((.*)\))?$"
```

### 2. Ps_Engine.cls（RPAエンジンのコアクラス）
* **修正点1（64bit VBAでのSleep API定義バグ防止）：** Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) となっていますが、Win32 API の Sleep の引数 dwMilliseconds は64bit環境でも32bit整数（DWORD = Long）です。Mod_TestDevTools.bas や UserForm_Picker.frm では正しく As Long になっているため、そちらと統一します。 
* **修正点2（GetTickCount の24.9日オーバーフロー対策）：** VBAの Long 型は符号付き32bit整数であるため、PCを約24.8日間再起動せずに稼働させ続けると GetTickCount() が負の値に反転し、引き算でオーバーフローエラー（実行時エラー6）が発生します。64bit対応の GetTickCount64（LongLong / または符号なし差分計算）を用いるか、未使用変数 settings の整理を行います。

```Diff
@@ 10行目付近：API宣言の型修正（Sleepの引数はLong） @@
 ' --- API宣言 ---
-Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
-Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long ' 日跨ぎ防止用API
+Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
+Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long ' 日跨ぎ防止用API
 ' --- グローバル定数および変数定義 ---
-Private PS_TIMEOUT_SEC    As Integer            ' メイン側から制御
-Private Const PS_READ_RETRY_MS  As Integer = 50 ' CPU負荷と通信レスポンスの最適化を目的とした待機ミリ秒数 (50ms)
+Private PS_TIMEOUT_SEC    As Integer            ' メイン側から制御するタイムアウト秒数
+Private Const PS_READ_RETRY_MS  As Integer = 50 ' CPU負荷と通信レスポンスの最適化を目的とした待機ミリ秒数(50ms)
 Private psProcess   As Object
-Private m_IsBusy    As Boolean ' 実行中フラグ
+Private m_IsBusy    As Boolean ' 実行中ロックフラグ
```

```Diff
@@ 80行目付近：BuildPayload の未使用変数整理とコメント見直し @@
 ' --- コマンドおよびパラメータのJSONペイロード構築 ---
 Private Function BuildPayload(cmdName As String, dictParams As Scripting.Dictionary) As String
     Dim root As New Scripting.Dictionary
-    Dim settings As New Scripting.Dictionary
+    ' 仕様書第4.1章の拡張用Settings枠（将来拡張用として保持）
+    ' Dim settings As New Scripting.Dictionary
     
     root.Add "Command", cmdName
     If dictParams Is Nothing Then
         root.Add "Parameters", New Scripting.Dictionary
     Else
         root.Add "Parameters", dictParams
     End If
     
-    ' VBA-JSONによる安全なJSON文字列への自動変換
+    ' VBA-JSONによる安全なJSON文字列への自動変換を実行する。
     BuildPayload = JsonConverter.ConvertToJson(root)
 End Function
```

---
#### 3. Mod_TestDevTools.bas（ピッカー制御・自動テスト）
* **修正点1（重大バグ修正：タイマー計測の開始時刻取得漏れ）：** Debug_Export_Loop の255行目（Export-WebDomSnapshot の計測）において、QueryPerformanceFrequency freq は呼ばれていますが、肝心の QueryPerformanceCounter startTime が抜けています。そのため、前回のピッカー実行時の startTime との差分が計算されてしまい、異常に大きなミリ秒が出力されるバグを修正します。 
* **修正点2（シート参照の不整合修正）：** Test_DevTools_AutoTests の430行目 ws.Range(Cells(glb_TestLoop_STA, 7), Cells(99, 8)).ClearContents において、内側の Cells に ws. が付いていないため、TestCases 以外のシートがアクティブな状態で実行すると実行時エラー（1004）になります。ws.Cells に修正します。 
* **修正点3（JSコーディング規約準拠）：** Debug_Picker_Loop（リアルタイムハイライトJS）および TestId_Type_iframe1 内のJSコードに var が使われているため、規約（var 禁止、const/let 統一）に合わせて修正します。

```Diff
@@ 250行目付近：Debug_Export_Loop 内の startTime 計測漏れ修正 @@
     ' --- Export-WebDomSnapshot ---
     If glb_CheckBox4 = True And useCdpPort > 0 Then
-        ' CPUのタイマー周波数を取得 (高精度計測の準備)
+        ' CPUのタイマー周波数および開始時刻を取得する (高精度計測の準備)。
         QueryPerformanceFrequency freq
+        QueryPerformanceCounter startTime
     
         fileName = "T" & Format(loopCount, "00") & "-" & "DomSnapshot"
         rpaEngine.RunAction "Export-WebDomSnapshot", CreateParams("Prefix", fileName)
         QueryPerformanceCounter endTime
         elapsedTimeMs = (endTime - startTime) / freq * 1000
         Debug.Print " ◆ 計測/Export-WebDomSnapshot (TimeMs) : " & Format(elapsedTimeMs, "0.00")
     End If
```

```Diff
@@ 425行目付近：Test_DevTools_AutoTests の Range(ws.Cells...) 修飾漏れ修正 @@
 Private Sub Test_DevTools_AutoTests()
     Dim ws As Worksheet
     Set ws = ThisWorkbook.Sheets("TestCases")
-    ' 結果列をクリア
-'    ws.Range("G14:H100").ClearContents
-    ws.Range(Cells(glb_TestLoop_STA, 7), Cells(99, 8)).ClearContents
+    ' 結果列をクリアする。
+'    ws.Range("G14:H100").ClearContents
+    ws.Range(ws.Cells(glb_TestLoop_STA, 7), ws.Cells(99, 8)).ClearContents
```

```Diff
@@ 540行目付近：TestId_Type_iframe1 の JS var 排除（const統一） @@
 ' --- TestId: iframe 固有のポップアップ消去 ---
 Private Sub TestId_Type_iframe1()
-    ' 警告バナー（Learn more）の閉じるボタン(X)をJSでクリックして非表示にする
+    ' 警告バナー（Learn more）の閉じるボタン(X)をJSでクリックして非表示にする。
     Dim jsClearBanner As String
-    jsClearBanner = "var btn = document.querySelector('.tox-notification__dismiss'); if(btn) btn.click();"
+    jsClearBanner = "const btn = document.querySelector('.tox-notification__dismiss'); if(btn) btn.click();"
     rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsClearBanner)
     Sleep 1000  ' バナーが消えるアニメーション待ち
 End Sub
```

#### 4. UserForm_Picker.frm（ピッカー操作フォーム）
* **修正点（コメントと実値の整合性）：** セクション見出しのタイポ（[選択Ur] → [選択Url]）を整えます。

```Diff
@@ 185行目付近：コメント見出しのタイポ修正 @@
-' --- [選択Ur]コンボボックスでの項目選択 ---
+' --- [選択Url] コンボボックスでの項目選択 ---
 Private Sub cmbUrlList_Change()
```

---
### ◆◆ Step2 ◆◆
VBA: 修正コードの検証コード(主な改修点を検証する)、モジュールの項目番号を確認○で表現した。

**Mod_TestRun2_Base少し高度**

```vba
    prompt = "パートを選んでください：" & vbCrLf & _

            " ---" & vbCrLf & _
            " 21. PDFをブラウザ内蔵ビューアで（注:Adobe Acrobat等）" & vbCrLf & _
            " 99. 【改修コード確認】v301(Argon)"
             
    ans = InputBox(prompt, "テストパート選択")
```

```vba
Private Function TEST_Part99() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 99: 【v302 改修】検証テスト"
    Debug.Print "=========================================================="
'1. Ps_Engine_Core_v302.ps1（司令塔・ルーター）　の　@@ 415行目付近：Switch-Tab 内のCDP再接続制御 @@
'3. Lib-WebView2_Init_v302.ps1（ブラウザ画面起動）　の　@@ 185行目付近：仮想タブ初期化時のUrlキー初期定義 @@
'4. Lib-WebView2_Native_v302.ps1（ネイティブ通信）　の　@@ 130行目付近：JSエラーの二重ラップ防止およびスペース修正 @@
'5. Lib-WebCDP_v302.ps1（WebSocket・CDP高速通信）　の　@@ 8行目付近：グローバル変数の明示的初期化 @@
'5. Lib-WebCDP_v302.ps1（WebSocket・CDP高速通信）　の　@@ 315行目付近：Invoke-CdpScript の複数行文字列アンエスケープ対応 @@
'7. Lib-WebXPath_v302.ps1（XPath特殊操作）　の　@@ 225行目付近：Set-WebXPathTextInput のピッカー生成パラメータ受け皿追加 @@
'8. Lib-WebAction_v302.ps1（Web標準操作）　の　@@ 585行目付近：Get-WebAttribute の多段iframe/ShadowDOM対応および型評価純粋化 @@

    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String

    ' --------------------------------------------------------------------------
    ' 【確認 3 & 5-1】
    ' ・3: Lib-WebView2_Init_v302.ps1 (仮想タブ初期化時のUrlキー初期定義)
    ' ・5-1: Lib-WebCDP_v302.ps1 (グローバル変数の明示的初期化)
    ' --------------------------------------------------------------------------
    Debug.Print "----------------------------------------------------------"
    Debug.Print " >> [確認 3 & 5-1] 初期状態の List-Tabs (Urlキー定義) と CDP初回通信"
    
    ' 画面遷移前の初期状態でも、仮想タブオブジェクトに Url キーが存在しエラーにならないか確認
    result = rpaEngine.RunAction("List-Tabs")
    Debug.Print " ◆ [確認3] 初期List-Tabs応答: " & result
    
    ' グローバル変数が明示的に初期化され、初回のCDPコマンドがStrictModeエラーなく通るか確認
    result = rpaEngine.RunAction("Invoke-CdpCommand", CreateParams("Method", "Browser.getVersion"))
    Debug.Print " ◆ [確認5-1] 初回CDPコマンド応答(先頭80文字): " & Left(result, 80) & "..."

    ' --------------------------------------------------------------------------
    ' 【確認 4】
    ' ・4: Lib-WebView2_Native_v302.ps1 (JSエラーの二重ラップ防止およびスペース修正)
    ' --------------------------------------------------------------------------
    Debug.Print "----------------------------------------------------------"
    Debug.Print " >> [確認 4] Invoke-WebView2NativeScript のJSエラー二重ラップ防止確認"
    
    targetUrl = "https://the-internet.herokuapp.com/login"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"

    ' 意図的にJavaScript内で例外をスローし、エラーメッセージの整形状態を確認する
    On Error Resume Next
    Err.Clear
    result = rpaEngine.RunAction("Invoke-WebView2NativeScript", CreateParams("Js", "throw new Error('CustomJsError_Test4');"))
    If Err.Number <> 0 Then
        Dim rpaErr4 As RpaExceptionInfo
        rpaErr4 = ParseRpaError(Err.Description)
        Debug.Print " ◆ [確認4] 捕捉したJSエラーメッセージ: " & rpaErr4.Message
        
        ' 二重ラップ("JS Error: JS Error:" 等)になっていないかをチェック
        If InStr(1, rpaErr4.Message, "JS Error: JS Error:", vbTextCompare) > 0 Then
            Debug.Print " !! [NG] JSエラーが二重ラップされています: " & rpaErr4.Message
        Else
            Debug.Print " >> [OK] 二重ラップなし・正常なフォーマットで捕捉できました。"
        End If
        Err.Clear
    Else
        Debug.Print " !! [NG] JSエラーが発生しませんでした。"
    End If
    On Error GoTo ErrorHandler

    ' --------------------------------------------------------------------------
    ' 【確認 5-2】
    ' ・5-2: Lib-WebCDP_v302.ps1 (Invoke-CdpScript の複数行文字列アンエスケープ対応)
    ' --------------------------------------------------------------------------
    Debug.Print "----------------------------------------------------------"
    Debug.Print " >> [確認 5-2] Invoke-CdpScript 複数行文字列のアンエスケープ確認"
    
    ' 改行(\n)とタブ(\t)を含む文字列をCDP経由で返却させる
    result = rpaEngine.RunAction("Invoke-CdpScript", CreateParams("Js", "return 'Line1\nLine2\tTabbed';"))
    Debug.Print " ◆ [確認5-2] 取得結果:" & vbCrLf & result
    
    If InStr(result, vbLf) > 0 And InStr(result, "\n") = 0 Then
        Debug.Print " >> [OK] 複数行文字列(\\n, \\t)が実際の改行・タブに正しくアンエスケープされています。"
    Else
        Debug.Print " !! [要確認] エスケープ文字(\\n)がそのまま残っている可能性があります: " & result
    End If

    ' --------------------------------------------------------------------------
    ' 【確認 7】
    ' ・7: Lib-WebXPath_v302.ps1 (Set-WebXPathTextInput のピッカー生成パラメータ受け皿追加)
    ' --------------------------------------------------------------------------
    Debug.Print "----------------------------------------------------------"
    Debug.Print " >> [確認 7] Set-WebXPathTextInput のピッカー生成パラメータ受け皿確認"
    
    ' ピッカー自動生成コードが渡してくる補助パラメータ(FrameSelector や OuterHtml 等)を付与して実行
    rpaEngine.RunAction "Set-WebXPathTextInput", CreateParams( _
        "XPath", "//input[@id='username']", _
        "Value", "tomsmith_picker_test", _
        "OuterHtml", "", _
        "TimeoutSec", 5)
        
    result = rpaEngine.RunAction("Get-WebXPathText", CreateParams("XPath", "//input[@id='username']"))
    Debug.Print " ◆ [確認7] 入力後の値取得: " & result
    If result = "tomsmith_picker_test" Then
        Debug.Print " >> [OK] 追加パラメータ付きでもエラーなくXPath入力が成功しました。"
    End If

    ' --------------------------------------------------------------------------
    ' 【確認 8】
    ' ・8: Lib-WebAction_v302.ps1 (Get-WebAttribute の多段iframe/ShadowDOM対応および型評価純粋化)
    ' --------------------------------------------------------------------------
    Debug.Print "----------------------------------------------------------"
    Debug.Print " >> [確認 8] Get-WebAttribute の型評価純粋化 ＆ iframe/ShadowDOM対応確認"
    
    ' (8-A) チェックボックスページで checked 属性 / type 属性 / 存在しない属性 の型評価を確認
    targetUrl = "https://the-internet.herokuapp.com/checkboxes"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    result = rpaEngine.RunAction("Get-WebAttribute", CreateParams("Selector", "form#checkboxes input:nth-of-type(1)", "Attribute", "type"))
    Debug.Print " ◆ [確認8-A1] type属性 (期待値: checkbox): [" & result & "]"
    
    result = rpaEngine.RunAction("Get-WebAttribute", CreateParams("Selector", "form#checkboxes input:nth-of-type(2)", "Attribute", "checked"))
    Debug.Print " ◆ [確認8-A2] checked属性 (ON側): [" & result & "]"
    
    result = rpaEngine.RunAction("Get-WebAttribute", CreateParams("Selector", "form#checkboxes input:nth-of-type(1)", "Attribute", "data-non-existent"))
    Debug.Print " ◆ [確認8-A3] 存在しない属性 (期待値: 空文字): [" & result & "]"

    ' (8-B) Shadow DOM ページでの Get-WebAttribute 確認
    targetUrl = "https://the-internet.herokuapp.com/shadowdom"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    result = rpaEngine.RunAction("Get-WebAttribute", CreateParams("Selector", "span[slot='my-text']", "Attribute", "slot"))
    Debug.Print " ◆ [確認8-B] ShadowDOM関連要素の slot属性 (期待値: my-text): [" & result & "]"

    ' (8-C) iframe ページでの FrameSelector 指定による Get-WebAttribute 確認
    targetUrl = "https://the-internet.herokuapp.com/iframe"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    rpaEngine.RunAction "Wait-WebElementInFrame", CreateParams("FrameSelector", "#mce_0_ifr", "ElementSelector", "body#tinymce")
    
    result = rpaEngine.RunAction("Get-WebAttribute", CreateParams("FrameSelector", "#mce_0_ifr", "Selector", "body#tinymce", "Attribute", "id"))
    Debug.Print " ◆ [確認8-C] iframe内部要素の id属性 (期待値: tinymce): [" & result & "]"

    ' --------------------------------------------------------------------------
    ' 【確認 1】
    ' ・1: Ps_Engine_Core_v302.ps1 (Switch-Tab 内のCDP再接続制御)
    ' --------------------------------------------------------------------------
    Debug.Print "----------------------------------------------------------"
    Debug.Print " >> [確認 1] Switch-Tab 実行時の CDP再接続制御の確認"
    
    targetUrl = "https://the-internet.herokuapp.com/windows"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"
    
    ' 「Click Here」をクリックして新規タブ(New Window)を開く
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", ".example a")
    Sleep 1500 ' 新規仮想タブの生成待機
    
    ' 新規タブ生成後の List-Tabs を確認 (確認3の仮想タブUrlキーの追加検証も兼ねる)
    result = rpaEngine.RunAction("List-Tabs")
    Debug.Print " ◆ [確認1 & 3] 新規タブ生成後の List-Tabs: " & result
    
    ' 新規タブ(タイトルに "New Window" を含むタブ、または P02)へ切り替え
    rpaEngine.RunAction "Switch-TabByTitle", CreateParams("TitleSubstring", "New Window")
    ' ※ もし TabId 指定でテストする場合は以下でも可:
    ' rpaEngine.RunAction "Switch-Tab", CreateParams("TabId", "P02")
    
    ' Switch-Tab 内で CDP が新しいタブへ再接続されているか、Invoke-CdpScript で現在のURLとTitleを取得して検証
    result = rpaEngine.RunAction("Invoke-CdpScript", CreateParams("Js", "return document.title + ' | ' + location.href;"))
    Debug.Print " ◆ [確認1] 切替後のタブでのCDP評価結果: " & result
    If InStr(result, "New Window") > 0 Then
        Debug.Print " >> [OK] Switch-Tab 後にCDPが新タブへ正常に再接続されています！"
    Else
        Debug.Print " !! [NG] CDPの向き先が旧タブのままになっている可能性があります: " & result
    End If
    
    ' 元のタブ(P01)へ戻して再度CDP再接続を確認
    rpaEngine.RunAction "Switch-Tab", CreateParams("TabId", "P01")
    result = rpaEngine.RunAction("Invoke-CdpScript", CreateParams("Js", "return document.title + ' | ' + location.href;"))
    Debug.Print " ◆ [確認1] P01復帰後のCDP評価結果: " & result

    Debug.Print "=== テスト完了 ==="

    TEST_Part99 = True
    Exit Function
ErrorHandler:
    Call HandleError_Type1("TEST_Part99")
    TEST_Part99 = False
End Function
```

---
### ◆◆ Step3 ◆◆
### 1.[（vbaモジュール）確認テスト: 5-2] 複数行文字列で Line1 しか取得できない。

1.	PowerShell側が [Console]::WriteLine("[RESULT]$($res.Trim())") を実行する。 
2.	$res に本物の改行が含まれているため、標準出力（パイプ）には

```text
1行目：[RESULT]Line1
2行目：Line2 Tabbed
3行目：[SUCCESS]
```

3.	一方、VBA側の Ps_Engine.cls（SendAndReceive 120行目付近）では、psProcess.StdOut.ReadLine() で1行ずつ読み取り、Left(line, 8) = "[RESULT]" の行だけを capturedResult に格納し、次の行（Line2 Tabbed）はプレフィックスが無いため Else 節の Debug.Print "PS_STDOUT: " & line として捨ててしまっていた。

PowerShell側が複数行のテキスト（複数行のセル文字やHTML、今回の Line1\nLine2 など）を返した際、VBA側で [RESULT] から [SUCCESS] の直前までの行をすべて連結して受け取るように Ps_Engine.cls の SendAndReceive を改良します。 

```Diff
@@ 110行目付近：SendAndReceive 内の複数行 [RESULT] 受信対応 @@
 Private Function SendAndReceive(jsonStr As String) As String
     Dim line As String
     Dim capturedResult As String
+    Dim isCapturingResult As Boolean
     Dim startTick As Long
     
     Debug.Print "VBA -> PS: " & jsonStr
     
     ' PowerShellへのコマンド送信を実行する。
     psProcess.StdIn.WriteLine jsonStr
     
     ' 監視開始時刻を記録する (GetTickCountを使用)。
     startTick = GetTickCount()
+    isCapturingResult = False
     
     ' 応答待ちループを実行する。
     Do While True
@@ (中略) @@
         ' 3. 標準出力の読み取りと内容解析を実行する。
         If Not psProcess.StdOut.AtEndOfStream Then
             line = psProcess.StdOut.ReadLine()
 
             If Left(line, 8) = "[RESULT]" Then
                 capturedResult = Mid(line, 9)
+                isCapturingResult = True
             ElseIf Left(line, 9) = "[SUCCESS]" Then
+                isCapturingResult = False
                 SendAndReceive = capturedResult
                 Exit Function
             ElseIf Left(line, 7) = "[ERROR]" Then
+                isCapturingResult = False
                 m_IsBusy = False
                 Debug.Print "PS_STDOUT: " & line
                 Err.Raise 999, "PS_Engine", Trim(Mid(line, 8))
                 Exit Function
+            ElseIf Left(line, 5) = "[LOG]" Then
+                ' ログ行は複数行RESULTの途中であってもコンソール出力のみとする。
+                Debug.Print "PS_STDOUT: " & line
             Else
-                Debug.Print "PS_STDOUT: " & line
+                ' [RESULT]受信開始後に改行された2行目以降のデータ行を改行付きで連結する。
+                If isCapturingResult Then
+                    capturedResult = capturedResult & vbLf & line
+                Else
+                    Debug.Print "PS_STDOUT: " & line
+                End If
             End If
         End If
```

### 2. [確認テスト: 1] の Switch-Tab における2つの問題
#### 🔍 現象A：新規タブ（仮想タブ）の Url が空文字のままになっている

```text
[Switch-TabByTitle] 情報: タブ切り替え (dc0066a5-... - New Window)
[Connect-CdpSession] アクティブURL = ''
```

> なぜ List-Tabs では "Url":"https://.../windows/new" と出ていたのに、Connect-CdpSession では アクティブURL = '' になったのか？
> 理由は、Lib-WebView2_Init_v302.ps1 の仮想タブ生成イベント（215行目付近）の add_NavigationCompleted 内で、クロージャ（.GetNewClosure()）が付いていなかったため、イベント発火時に $tabId が参照できず $global:Tabs[$tabId].Url = $navSender.Source が空振りしていたためです！（List-Tabs は Url が空のとき $webview.Source.ToString() を直接見に行くため取得できていました）。 

#### 🔍 現象B：P01 に戻った際、CDPが New Window に誤爆接続している

```text
VBA -> PS: {"Command":"Switch-Tab","Parameters":{"TabId":"P01"}}
[Connect-CdpSession] アクティブURL = 'https://the-internet.herokuapp.com/windows'
[Connect-CdpSession] 情報: CDP接続完了 (Target: New Window)
 ◆ [確認1] P01復帰後のCDP評価結果: New Window | https://the-internet.herokuapp.com/windows/new
```

> P01（URL: .../windows）に戻したのに、CDPの評価結果が New Window（.../windows/new）のままになっています！
> 原因は、Lib-WebCDP_v302.ps1 の Connect-CdpSession（60行目付近）のURL柔軟一致条件にあります： <br>
> •	$_.url -like "$normActive*" これがあると、P01のURL（.../windows）で検索した際に、CDPのターゲット一覧の先頭にある新規タブ（.../windows/new）が ".../windows*" の前方一致にヒットしてしまうのです！ 

**（Lib-WebView2_Init と Lib-WebCDP）**<br>

1 **Lib-WebView2_Init_v302.ps1（215行目付近）：** add_NavigationCompleted に .GetNewClosure() を付与し、さらに Connect-CdpSession 側でも $activeTab.WebView.Source をフォールバック参照するようにします。<br>

```Diff
@@ Lib-WebView2_Init_v302.ps1 215行目付近：URL更新イベントへの GetNewClosure 付与 @@
                         # ポップアップ画面のURL更新イベントを登録する。
                         $senderInit.CoreWebView2.add_NavigationCompleted({
                             param($navSender, $navArgs)
                             try {
                                 $global:Tabs[$tabId].Url = $navSender.Source
                                 Write-DebugLog -Message "[$tabId] URL更新: $($navSender.Source)" -Level Info
                             } catch {}
-                        })
+                        }.GetNewClosure())
```

2 **Lib-WebCDP_v302.ps1（25行目付近 & 55行目付近）：**<br>
  * $activeUrl 取得時に $activeTab.Url が空なら $activeTab.WebView.Source.ToString() から確実に取得する。 <br>
  * URLマッチングで「完全一致」を最優先（ステップ2-A）とし、それでも見つからない場合のみ「前方一致（部分一致）」（ステップ2-B）を行うように分離する。

```Diff
@@ Lib-WebCDP_v302.ps1 25行目付近：アクティブURLの確実な取得 @@
     # アクティブタブ情報の取得： アクティブなタブのURLやTargetIdを変数に保持する。
     $activeTab    = $global:Tabs[$global:ActiveTabId]
     $activeUrl    = if ($activeTab.Url) { $activeTab.Url } else { "" }
+    if ([string]::IsNullOrWhiteSpace($activeUrl) -and $null -ne $activeTab.WebView -and $null -ne $activeTab.WebView.Source) {
+        try { $activeUrl = $activeTab.WebView.Source.ToString() } catch {}
+    }
     $activeTarget = $activeTab.TargetId
```

```Diff
@@ Lib-WebCDP_v302.ps1 55行目付近：URL完全一致の優先化（前方一致による子画面誤爆防止） @@
-            # 2. URLによる柔軟一致： TargetIdがない、またはマッチしなかった場合
+            # 2. URLによる一致探索： TargetIdがない、またはマッチしなかった場合、完全一致を最優先して子画面誤爆を防ぐ。
             if (-not $pageTarget -and -not [string]::IsNullOrWhiteSpace($activeUrl)) {
                 $normActive = $activeUrl.TrimEnd('/')
+                # 2-A. URLの完全一致を最優先で探索する。
+                $pageTarget = $pages | Where-Object {
+                    $_.type -eq 'page' -and
+                    $_.url.TrimEnd('/') -eq $normActive -and
+                    -not [string]::IsNullOrEmpty($_.webSocketDebuggerUrl)
+                } | Select-Object -First 1
+
+                # 2-B. 完全一致がない場合のみ、前方一致(リダイレクト等考慮)で探索する。
+                if (-not $pageTarget) {
+                    $pageTarget = $pages | Where-Object {
+                        $_.type -eq 'page' -and
+                        ($_.url -like "$normActive*" -or $normActive -like "$($_.url.TrimEnd('/'))*") -and
+                        -not [string]::IsNullOrEmpty($_.webSocketDebuggerUrl)
+                    } | Select-Object -First 1
+                }
-                $pageTarget = $pages | Where-Object {
-                    $_.type -eq 'page' -and
-                    ($_.url.TrimEnd('/') -eq $normActive -or $_.url -like "$normActive*" -or $normActive -like "$($_.url.TrimEnd('/'))*") -and
-                    -not [string]::IsNullOrEmpty($_.webSocketDebuggerUrl)
-                } | Select-Object -First 1
             }
```

3 **[確認テスト: 8-C] の Get-WebAttribute における FrameSelector 引数について**<br>

```vba
result = rpaEngine.RunAction("Get-WebAttribute", CreateParams("FrameSelector", "#mce_0_ifr", "Selector", "body#tinymce", "Attribute", "id"))
```

> 今回の改修で Get-WebAttribute は utilFindInFrames を使うようになったため、FrameSelector を指定しなくても自動的に全iframeを透過探索して [tinymce] を取得できています（ログでも [tinymce] と成功しています）。 ただし、このようにVBA側から明示的に "FrameSelector" パラメータを渡された場合でも、将来 StrictMode 等で「未定義パラメータ」とならないよう、Get-WebAttribute の param にオプショナルな受け皿として [string]$FrameSelector = "" を1行追加しておくと、後方互換性が完璧になります。

```Diff
@@ Lib-WebAction_v302.ps1 565行目付近：Get-WebAttribute のパラメータ受け皿補強 @@
 function Get-WebAttribute {
     param (
         [string]$Selector = "",
         [string]$XPath = "",
         [string]$Attribute = "",
+        [string]$FrameSelector = "",
         [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
     )
```
