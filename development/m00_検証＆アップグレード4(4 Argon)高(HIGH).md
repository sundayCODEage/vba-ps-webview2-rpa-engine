
<h4 align="center">汎用RPA操作エンジン (VBA × PowerShell × WebView2 Hybrid-Engine)</h4>
<h4 align="center">V302 ⇒ V303 （Gemini 4 Argon ／【Google AI Pro】）：　思考レベル: 高 ( HIGH )</h4>

<p align="right">
  作成日: 2026年10月26日 - 10月29日<br>
  作成者: System beginner & AI Collaborator (Gemini)
</p>

---
* _Step6: 「モジュール間の整合性、バグ検証、リファクタリング、安定化の包括的な分析」 を指示。_<br>
* _Step7: （一部、コメント修正:Diffなし）、最終(微)調整 **v303完成**_<br>
  * 主観的な強調語（「完全」「完璧」「絶対」等）や感嘆符（！）が点在していた。有る程度、検索で修正した。
  * テストログでの疑問点（非同期パイプ通信の問題としていた件、Get-WebAttributの修正など）、比較演算ルールなど
  * 仮想タブ（ポップアップ）でのサイレントダウンロード自動継承

😊😊😊

---
### ◆◆ Step6 ◆◆
#### Ps_Engine_Core_v303.ps1
1. Set-ActiveTab でタブを切り替えた際、Switch-Tab とは異なり $global:ConnectedCdpTabId = $null が行われていないため、Set-ActiveTab 単体でタブを戻した直後にCDPコマンドを実行すると旧タブへ送信されるリスクがあります。 
2. Wait-Condition でタイムアウトした際、$func = $MyInvocation.MyCommand.Name によりエラー元が常に [Wait-Condition] となり、呼び出し元（Wait-WebElement 等）が $TimeoutMessage 側に埋め込まれる二重表記（例：[ERROR][Wait-Condition] [Timeout]: [Wait-WebUrlContains] タイムアウト...）になっています。仕様書第4章の表「例: Wait-WebElement」と整合させるため、呼び出し元関数名を渡せるオプション引数（または自動抽出）にすると ParseRpaError の FunctionName が正確になります。 

```diff
@@ -355,6 +355,8 @@
     if ($global:Tabs.ContainsKey($TabId)) {
         $global:ActiveTabId = $TabId
         $webview = $global:Tabs[$TabId].WebView
         $webview.BringToFront()
         $global:WebViewCtrl = $webview
+        # タブ切替に伴うCDPターゲット同期： 次回CDP実行時に再接続させるため接続済みタブIDをリセットする。
+        $global:ConnectedCdpTabId = $null
         return "ActiveTab = $TabId"
     }
     throw (New-EngineException -Func $func -Type "引数エラー" -Message "指定されたタブIDが見つかりません" -Details $TabId)
```

```Diff
@@ -458,7 +460,8 @@
     param (
         [Parameter(Mandatory=$true)][scriptblock]$ConditionBlock,
         [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
         [int]$PollIntervalMs = $global:CONFIG.PollIntervalMs,
-        [string]$TimeoutMessage = "待機処理がタイムアウトしました"
+        [string]$TimeoutMessage = "待機処理がタイムアウトしました",
+        [string]$CallerFunc = ""
     )
 
-    $func = $MyInvocation.MyCommand.Name
+    $func = if (-not [string]::IsNullOrWhiteSpace($CallerFunc)) { $CallerFunc } else { $MyInvocation.MyCommand.Name }
 
     # UIフリーズを防止して条件成立を待機する。
     $sw = [System.Diagnostics.Stopwatch]::StartNew()
```

---
#### Lib-WebView2_Init_v303.ps1
*	仮想タブ閉鎖時（WindowCloseRequested および ProcessFailed）に親タブ（P01 等）へ自動復帰する際、$global:ConnectedCdpTabId = $null をリセットしておくと、直後の処理でより確実にCDP再接続が行われます（現状でも $global:ActiveTabId との不一致比較で再接続されますが、明示リセットしておくと安全です）。

```Diff
@@ -234,6 +234,8 @@
                                         if ($null -ne $nextTabId) {
                                             # アクティブタブIDを親のIDへ更新する。
                                             $global:ActiveTabId = $nextTabId
                                             # 紐づくWebView2コントロールも親のものへ切り替える。
                                             $global:WebViewCtrl = $parentCtrl
+                                            # CDP接続先タブIDをリセットし、次回操作時に親タブへ再接続させる。
+                                            $global:ConnectedCdpTabId = $null
                                             Write-DebugLog -Message "[WebView2] 情報: ゴーストを削除し、親タブ ($nextTabId) へ自動復帰しました。" -Level Info
                                         } else {
```

---
#### Lib-WebCDP_v303.ps1
1. Invoke-CdpCommand 内の try ... finally の後に catch が続く箇所について、PowerShell 5.1 では外側と内側の try が適切に分かれていますが、$cts（CancellationTokenSource）が finally で Dispose() されていないため、数千回のCDP通信時にハンドルリークを防ぐべく $cts.Dispose() を追加します。 
2. Invoke-CdpNativeClick および Set-CdpNativeTextInput がトップドキュメントのみ（deepQuerySelector(selector, document)）を対象としており、utilFindInFrames を使っていないため、同一オリジンiframe内の要素にもCDPネイティブ操作が届くよう補強すると一貫性が高まります（現状維持でも動作しますが、微修正として提示します）。 

```Diff
@@ -286,6 +286,8 @@
             } else {
                 throw (New-EngineException -Func $func -Type "CDPエラー" -Message "CDPコマンドの送受信中に予期せぬ例外が発生しました" -Details $_.Exception.Message)
             }
+        } finally {
+            if ($null -ne $cts) { try { $cts.Dispose() } catch {} }
         }
     }
 }
```

---
#### Lib-DevTools_v303.ps1
* コーディング規約違反（ローカル変数のPascalCase）: Get-WebSelectorFromUiaFuzzy 内のヘルパー関数 Get-LevenshteinDistanceFast は param([string]$SourceStr, [string]$TargetStr) をローカル変数 $s, $t に受けており規約通りですが、Build-FallbackStepCode および Build-RpaStepCode が生成するデスクトップUIA用コマンドが旧関数名 Execute-UIRpaStep になっています。

```Diff
@@ -161,7 +161,7 @@
         $domText = "$($node.placeholder) $($node.name) $($node.id) $($node.tagName) $($node.innerText)".Trim().ToLower()
         if (-not $domText) { continue }
 
-        # --- 新スコアリングロジック ---
+        # 新スコアリングロジック： 属性の一致度と文字列の類似度から総合スコアを算出する。
         $score = 0
 
         # 1. IDまたはNameの直接比較： UIAのID・NameとDOM属性を比較しスコアを加算する。
```

```Diff
@@ -307,7 +307,7 @@
 
                 // 隠しフィールド(type="hidden")は操作不能なためスキップする。
                 if (el.tagName === 'INPUT' && el.type === 'hidden') continue;
-                // --- RPAの主要操作対象の特定 (優先順位付き) ---
+                // RPAの主要操作対象の特定（優先順位付き）： 操作頻度の高い要素から順に判定する。
              
                 // 優先度1: contenteditableな要素 (TinyMCE等のリッチテキストエディタのBodyを最優先)
                 if (el.isContentEditable) return el;
```

ご指摘の通り不要になっていた変数・引数<br>
Build-FallbackStepCode および Build-RpaStepCode において、旧仕様（Execute-UIRpaStep）の名残で残っていた以下の引数・変数は、現在の Invoke-UiaAction / Switch-AppWindow 体系では使用されておらず、不要変数となっていました。 
* **不要だった引数:** $Index, $AnchorName, $AnchorDirection, $AnchorSteps 
* **未活用だった変数:** $safeWindowName（ウィンドウ切り替え Switch-AppWindow として出力するか、不要なら整理）、$ControlType（Edit なら Input アクション、それ以外は Click アクションへの分岐に活用すべき変数）、$TargetId（AutomationId として引き継ぐとUIAの精度が向上） 
これらを整理し、ControlType が "Edit" の場合は Action "Input"、それ以外は Action "Click" を生成し、AutomationId も出力する実用的なコードへ修正します。

```Diff
@@ -612,23 +612,21 @@
 function Build-FallbackStepCode {
     param(
         [string]$WindowName,
         [string]$TargetName,
+        [string]$TargetId = "",
         [string]$ControlType,
         [string]$CurrentValue,
         [bool]$IsGetNameMode
     )
 
     # GetName(値の取得)モードの場合、TargetNameをワイルドカード(*)にして柔軟性を持たせる。
     $stepTargetName = if ($IsGetNameMode) { "*" } else { $TargetName }
+    # コントロール種別に応じて発行するUIAアクション(GetText / Input / Click)を決定する。
+    $uiaAction = if ($IsGetNameMode) { "GetText" } elseif ($ControlType -eq "Edit") { "Input" } else { "Click" }
 
-    # ネイティブUIA用のコマンド(Execute-UIRpaStep)を組み立てる関数を呼び出す。
+    # ネイティブUIA用のコマンド(Invoke-UiaAction)を組み立てる関数を呼び出す。
     $stepCode = Build-RpaStepCode `
-        -Action "Click" `
+        -Action $uiaAction `
         -WindowName $WindowName `
         -TargetName $stepTargetName `
+        -TargetId $TargetId `
-        -ControlType $ControlType `
-        -Index 1 `-        -ActionValue $CurrentValue `
-        -AnchorName "" `
-        -AnchorDirection "None" `
-        -AnchorSteps 0
+        -ActionValue $CurrentValue
```

```Diff
@@ -725,7 +723,7 @@
         # ルートC： デスクトップ要素へのフォールバック。
-        $resObj = Build-FallbackStepCode -WindowName $stepWindowName -TargetName $targetName -ControlType $targetControlType -CurrentValue $currentValue -IsGetNameMode $IsGetNameMode
+        $resObj = Build-FallbackStepCode -WindowName $stepWindowName -TargetName $targetName -TargetId $targetId -ControlType $targetControlType -CurrentValue $currentValue -IsGetNameMode $IsGetNameMode
         $resObj.command = "[Desktop-UIA] " + $resObj.command
         return "[RESULT]" + ($resObj | ConvertTo-Json -Compress)
```

```Diff
@@ -835,31 +833,31 @@
 # --- 実行コマンド文字列の組み立て ---
-# UIA/Web用の実行コマンド文字列(PowerShellコード)を組み立てる。
+# デスクトップUIA用の実行コマンド文字列(Invoke-UiaAction)を組み立てる。
 function Build-RpaStepCode {
     param(
         [string]$Action, 
         [string]$WindowName, 
         [string]$TargetName, 
-        [string]$ControlType, 
-        [int]$Index, 
-        [string]$ActionValue, 
-        [string]$AnchorName, 
-        [string]$AnchorDirection, 
-        [int]$AnchorSteps
+        [string]$TargetId = "", 
+        [string]$ActionValue = ""
     )
 
     $func = $MyInvocation.MyCommand.Name
 
     try {
         # 文字列内のダブルクォートをエスケープ： VBAやPSで安全に実行するため処理する。
         $safeActionValue = $ActionValue -replace '"', '`"'
         $safeTargetName  = $TargetName -replace '"', '`"'
         $safeWindowName  = $WindowName -replace '"', '`"'
-        $safeAnchorName  = $AnchorName -replace '"', '`"'
+        $safeTargetId    = $TargetId -replace '"', '`"'
+
+        $idParam  = if (-not [string]::IsNullOrWhiteSpace($safeTargetId)) { " -AutomationId `"$safeTargetId`"" } else { "" }
+        $winNote  = if (-not [string]::IsNullOrWhiteSpace($safeWindowName)) { " (Window: $safeWindowName)" } else { "" }
 
-        if ($Action -eq "GetName") {
-            # 取得アクションの場合、結果を変数($extractedValue)へ代入し、コンソールへ出力するコードを組み立てる。
-            return "`$extractedValue = Execute-UIRpaStep -TargetWindowName `"$safeWindowName`" -TargetName `"$safeTargetName`" -TargetType `"$ControlType`" -Index $Index -Action `"$Action`" -ActionValue `"$safeActionValue`" -AnchorName `"$safeAnchorName`" -AnchorDirection `"$AnchorDirection`" -AnchorSteps $AnchorSteps`r`nWrite-Host `"  => 取得した値: `$extractedValue`" -ForegroundColor Cyan"
+        if ($Action -in @("GetName", "GetText", "GetValue")) {
+            # 取得アクションの場合、Invoke-UiaAction (GetText) 用の実行コマンド文字列を組み立てる。
+            return "Invoke-UiaAction -Action `"GetText`" -Name `"$safeTargetName`"$idParam -TimeoutSec 10$winNote"
+        } elseif ($Action -eq "Input") {
+            # 入力アクションの場合、Value付きの Invoke-UiaAction コマンド文字列を組み立てる。
+            return "Invoke-UiaAction -Action `"Input`" -Name `"$safeTargetName`"$idParam -Value `"$safeActionValue`" -TimeoutSec 10$winNote"
         } else {
             # クリックなどの標準アクションの場合、そのまま実行用コマンドを組み立てて返却する。
-            return "Execute-UIRpaStep -TargetWindowName `"$safeWindowName`" -TargetName `"$safeTargetName`" -TargetType `"$ControlType`" -Index $Index -Action `"$Action`" -ActionValue `"$safeActionValue`" -AnchorName `"$safeAnchorName`" -AnchorDirection `"$AnchorDirection`" -AnchorSteps $AnchorSteps"
+            return "Invoke-UiaAction -Action `"$Action`" -Name `"$safeTargetName`"$idParam -TimeoutSec 10$winNote"
         }
     } catch {
         throw (New-EngineException -Func $func -Type "内部エラー" -Message "ステップコードの生成に失敗しました" -Details $_.Exception.Message)
     }
 }
```

---
#### Lib-WebDebug_v303.ps1

```Diff
@@ -429,15 +429,15 @@
         # RPAブラウザフォーム(WinForms)の論理的な境界矩形を取得する。
         $bounds = $global:BrowserForm.Bounds
 
-        # --- DPIスケーリング(125%など)の補正計算 ---
+        # DPIスケーリング(125%など)の補正計算： OSレベルのGraphicsオブジェクトを生成し、X軸のDPIを取得する。
-        # OSレベルのGraphicsオブジェクトを生成し、X軸のDPIを取得する。
         $tmpGraphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
         $dpiX = $tmpGraphics.DpiX
         $tmpGraphics.Dispose()
         # 標準DPI(96)で除算し、スケーリング倍率を算出する。
         $scale = $dpiX / 96.0  # 100%時は1.0、125%時は1.25になる。
 
-        # --- 物理ピクセル座標・サイズへの変換 ---
+        # 物理ピクセル座標・サイズへの変換： 論理座標とサイズにスケーリング倍率を乗算し、物理ピクセル値を確定する。
-        # 論理座標とサイズにスケーリング倍率を乗算し、スクリーンショット用の物理ピクセル値を確定する。
         $phyX = [int][Math]::Round($bounds.X * $scale)
         $phyY = [int][Math]::Round($bounds.Y * $scale)
         $phyW = [int][Math]::Round($bounds.Width * $scale)
```

```Diff
@@ -608,7 +608,7 @@
                 innerText      = ""
             }
 
-            # --- タグ名 (nodeName) の復元 ---
+            # タグ名 (nodeName) の復元： インデックス番号をもとに辞書からタグ名文字列を引き当てる。
             # nodeName配列には直接の文字列ではなく、$strings 配列への「インデックス番号」が入っている。
             if ($nodes.nodeName -and $null -ne $nodes.nodeName[$i]) {
                 # 該当インデックス番号をもとに、辞書($strings)から実際のタグ名文字列を引き当てる。
```

```Diff
@@ -618,7 +618,7 @@
                 }
             }
 
-            # --- 属性 (attributes) の復元 ---
+            # 属性 (attributes) の復元： 2つずつ(Step 2)取り出して属性名と値のペアを復元する。
             <# attributes配列は [名前のインデックス, 値のインデックス, 名前のインデックス, 値のインデックス...] 
                という連続した1次元配列として格納されているため、2つずつ(Step 2)取り出して復元する。 #>
             if ($nodes.attributes -and $nodes.attributes[$i]) {
```

---
#### Lib-WebAction_v303.ps1
1. JS内の let 再代入不要変数（コーディング規約8.3「再代入が不要な変数は const、必要な変数は let に統一する」）: Wait-WebScreenUnlock 内の for (let d of divs) および for (let f of frames) はループ内で再代入されないため const に統一します。 
2. Get-WebAttribute における動的プロパティ（checked 等）の取得補強: Mod_TestRun1_Base Part 5 のログで Get-WebAttribute -Attribute "checked" が空文字になる現象への対策として、el.getAttribute(attr) が null だった場合、Booleanプロパティ（checked, disabled, readonly, selected）が true であれば "true" を返し、value 属性が null であれば el.value を返すフォールバックを1行入れると、HTML初期属性がない動的変更後の要素でも確実に状態を取得できるようになります。

```Diff
@@ -370,7 +370,7 @@
         function checkMask(doc, win) {
             // 画面上のすべてのdiv要素を取得する。
             const divs = doc.querySelectorAll('div');
-            for (let d of divs) {
+            for (const d of divs) {
                 const s = win.getComputedStyle(d);
                 // s.zIndexが "auto" や空文字だった場合に "0" としてパースさせる。
                 const z = parseInt(s.zIndex || "0", 10);
```

```Diff
@@ -383,7 +383,7 @@
             
             // iframe / frame の中も再帰的に探査する。
             const frames = doc.querySelectorAll('iframe, frame');
-            for (let f of frames) {
+            for (const f of frames) {
                 try {
                     if (f.contentDocument && f.contentWindow) {
                         if (checkMask(f.contentDocument, f.contentWindow)) return true;
```

```Diff
@@ -718,13 +718,17 @@
                 const sel = decodeURIComponent(escape(atob('$selB64')));
                 const attr = decodeURIComponent(escape(atob('$attrB64')));
                 const frmSel = decodeURIComponent(escape(atob('$frameB64')));
+                const extractAttr = function(el, a) {
+                    if (!el) return null;
+                    const val = el.getAttribute(a);
+                    if (val !== null) return val;
+                    if (a === 'checked' || a === 'disabled' || a === 'readonly' || a === 'selected') return el[a] ? 'true' : null;
+                    return null;
+                };
                 if (frmSel) {
                     const frm = utilFindInFrames(window, function(win) {
                         try { return deepQuerySelector(frmSel, win.document); } catch(e) { return null; }
                     });
                     if (!frm || !frm.contentDocument) return null;
                     const el = deepQuerySelector(sel, frm.contentDocument);
-                    return el ? el.getAttribute(attr) : null;
+                    return extractAttr(el, attr);
                 }
                 return utilFindInFrames(window, function(win) {
                     try {
                         const el = deepQuerySelector(sel, win.document);
-                        return el ? el.getAttribute(attr) : null;
+                        return extractAttr(el, attr);
                     } catch(e) { return null; }
                 });
 "@
```

---
#### Lib-DesktopUIA_v303.ps1
* Mod_TestRun2_Base Part 99「検証3-B（XPath物理クリックフォールバック）」が [NG] になった原因および、Mod_TestDevTools.bas の TestId_Type_PhysicalFallback のコメントにある「Invoke-DesktopCoordinateClick が True + 改行(vbLf) + [FALLBACK_PHYSICAL]... の2行データを返す」現象の根本対策です。 
* Invoke-DesktopCoordinateClick の冒頭で $global:BrowserForm が存在する場合は SetForegroundWindow で確実に最前面化してから物理クリックを発火させることで、ウィンドウ非アクティブによる空振りを防ぎます。

```Diff
@@ -370,6 +370,11 @@
     $func = $MyInvocation.MyCommand.Name
 
     try {
+        # 物理クリック前にRPAブラウザウィンドウを最前面へ引き上げ、非アクティブによるクリック空振りを防ぐ。
+        if ($null -ne $global:BrowserForm -and $global:BrowserForm.Handle -ne [IntPtr]::Zero) {
+            try { [Win32Api.Win32Utils]::SetForegroundWindow($global:BrowserForm.Handle) | Out-Null } catch {}
+        }
+
         # マウスカーソルの移動： 物理ピクセル座標へマウスカーソルを瞬間移動させる。
         [Win32Api.Win32Utils]::SetCursorPos($X, $Y) | Out-Null
         Start-Sleep -Milliseconds $WaitMs
```

---
#### Mod_TestRun2_Base少し高度.bas

```Diff
@@ -668,6 +668,11 @@
     ' (3-B) 存在しないXPathに BoundingBox を添えて実行した際、物理クリックへフォールバックすることを確認
+    ' 3-Aの要素追加後の最新座標を再取得し、確実にボタン中心を狙う
+    bboxJson = rpaEngine.RunAction("Invoke-WebScript", CreateParams("Js", _
+        "const el = document.querySelector('.example button'); const r = el.getBoundingClientRect(); " & _
+        "return JSON.stringify({x: Math.round(r.left + window.scrollX), y: Math.round(r.top + window.scrollY), width: Math.round(r.width), height: Math.round(r.height)});"))
     result = rpaEngine.RunAction("Invoke-WebXPathClick", CreateParams( _
         "XPath", "//button[@id='not_exist_xpath_btn']", _
         "BoundingBox", bboxJson, _
         "TimeoutSec", 2))
```

---
### ◆◆ Step7 ◆◆
### Ps_Engine_Core_v303.ps1
* 【バグ修正】 メインループの catch ブロックにて、Write-DebugLog を先に出力してから [Console]::WriteLine("[ERROR]$errMsg") を出力するように順序を入れ替え、VBA側のイミディエイトウィンドウでエラーログが次コマンドに遅延混入する現象を解消します。

```Diff
@@ -380,7 +380,7 @@
 # --- アクティブタブ設定 ---
 # タブIDを直接指定してアクティブタブを切り替える（前面化）。
 function Set-ActiveTab {
-    param ([string]$TabId)
+    param ([Parameter(Mandatory=$true)][string]$TabId)
 
     $func = $MyInvocation.MyCommand.Name
 ```

```Diff
@@ -705,8 +705,6 @@
 
             } catch {
                 # エラー結果を出力する。
                 $errMsg = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
-                [Console]::WriteLine("[ERROR]$errMsg")
 <# **保留**  
                 # ログファイルに詳細なスタックトレース(PositionMessage)を記録する。
                 $posMsg = if ($_.InvocationInfo) {$_.InvocationInfo.PositionMessage -replace "`r`n", " " -replace "`n", " " } else { "" }
```

```Diff
@@ -721,6 +719,8 @@
                     "不明"
                 }
                 Write-DebugLog -Message "[Engine] コマンド実行エラー: $errMsg | トレース: $stackTrace" -Level Error
+                # VBA側のSendAndReceiveループが[LOG]行を読み切ってから終了するよう、最後に[ERROR]シグナルを出力する。
+                [Console]::WriteLine("[ERROR]$errMsg")
             }
         }
         Start-Sleep -Milliseconds 20
```

### Lib-WebJS_v303.ps1
* 【安定化】 isVisible 内で window.getComputedStyle(el) を呼ぶ際、対象要素が同一オリジン iframe 内の要素であっても該当フレームの defaultView から正確にスタイルを取得できるよう補正します。 

```Diff
@@ -65,11 +65,12 @@
 // RPAで誤動作の原因となる「透明化された要素」「サイズゼロの要素」「裏に隠れた要素」を、人間の視覚に近い基準で正確に除外するための安全装置として機能する。
 function isVisible(el, { checkOpacity = true, minOpacity = 0.01, silent = false } = {}) {
     // 要素が存在しない、またはDOMノード(NodeType=1)でない場合は非表示とみなす。
     if (!el || el.nodeType !== 1) {
         domTracker.log('isVisible: 失敗 - 要素が null または無効なノードです。');
         return false;
     }
     // 要素に適用されている最終的なCSSスタイルをOS(ブラウザ)から取得する。
-    const style = window.getComputedStyle(el);
+    const win = (el.ownerDocument && el.ownerDocument.defaultView) ? el.ownerDocument.defaultView : window;
+    const style = win.getComputedStyle(el);
     if (!style) return false;
```

### Lib-WebView2_Init_v303.ps1
* 【バグ修正】 ポップアップタブの add_ProcessFailed（クラッシュ復帰時）および add_WindowCloseRequested のフォールバック分岐に $global:ConnectedCdpTabId = $null を追加し、タブ破棄後のCDP再接続漏れを防ぎます。

```Diff
@@ -169,7 +169,7 @@
                         }
 
                         try {
-                            # CDPを利用して、このポップアップ画面を直接操作・デバッグする際に必要となる内部ID(TargetId)を抽出して保持する。
+                            # CDPを利用して、このポップアップ画面を直接操作・デバッグする際に必要となる内部ID(TargetId)を抽出して保持する（未取得時はURLマッチへフォールバックする）。
                             $global:Tabs[$tabId].TargetId = $evtArgs.WindowFeatures.TargetId
                         } catch {
                             Write-DebugLog -Message "[WebView2] 警告: targetId の取得に失敗" -Level Warning
```

```Diff
@@ -248,11 +248,13 @@
                                                 # 残存タブの先頭[0]を新しいアクティブタブとして強制的に設定する。
                                                 $global:ActiveTabId = $remainingTabs[0]
                                                 $global:WebViewCtrl = $global:Tabs[$global:ActiveTabId].WebView
+                                                $global:ConnectedCdpTabId = $null
                                                 Write-DebugLog -Message "[WebView2] 警告: 親タブ不明のため、残存タブ ($($global:ActiveTabId)) へ強制移行しました。" -Level Warning
                                             } else {
                                                 # 残存タブが存在しない場合、状態をリセットする。
                                                 $global:ActiveTabId = $null
                                                 $global:WebViewCtrl = $null
+                                                $global:ConnectedCdpTabId = $null
                                                 Write-DebugLog -Message "[WebView2] 情報: 全タブが閉じられたため、認識をリセットしました。" -Level Info
                                             }
                                         }
```

```Diff
@@ -296,6 +298,7 @@
                                         if ($null -ne $nextTabId) {
                                             $global:ActiveTabId = $nextTabId
                                             $global:WebViewCtrl = $parentCtrl
+                                            $global:ConnectedCdpTabId = $null
                                         }
                                     }
                                 } catch {}
```

### Lib-WebView2_Native_v303.ps1
* 【安定化】 ExecuteScriptAsync 失敗時のエラーメッセージ内の丸括弧 (ExecuteScriptAsync) を角括弧 [ExecuteScriptAsync] に変更し、VBA側の ParseRpaError で Details の括弧と混同されないようにします。 
* 【安定化】 JSスタックトレースの改行置換を -replace "rn|n|r", " | " に拡張し、V8エンジンが出力する \n 単体の改行も確実に1行化します。 

```Diff
@@ -66,7 +66,7 @@
             }
 
             if ($task.IsFaulted) {
-                throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "WebView2ネイティブAPI(ExecuteScriptAsync)の実行に失敗しました" -Details $task.Exception.InnerException.Message)
+                throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "WebView2ネイティブAPI [ExecuteScriptAsync] の実行に失敗しました" -Details $task.Exception.InnerException.Message)
             }
 
             # 正常終了した非同期タスクから、実行結果(文字列)を取り出す。
```

```Diff
@@ -122,7 +122,7 @@
 
             # JS側で発生した例外のエラー処理： スタックトレース等を抽出し、PowerShell側の例外として再送出する。
             if ($response.status -ne "success") {
-                $stack = $response.stack -replace "`r`n", " | "
+                $stack = $response.stack -replace "`r`n|`n|`r", " | "
                 throw (New-EngineException -Func $func -Type "JSエラー" -Message "$($response.name): $($response.message)" -Details $stack)
             }
```

### Lib-WebCDP_v303.ps1
* 【安定化】 Invoke-CdpScript 内のエラースロー時にも $response.stack の改行を " | " に置換し、Lib-WebView2_Native_v303.ps1 とフォーマットを統一します。 

```Diff
@@ -335,7 +335,8 @@
     }
 
     if ($response.status -ne "success") {
-        throw (New-EngineException -Func $func -Type "JSエラー" -Message "$($response.name): $($response.message)" -Details $response.stack)
+        $stack = $response.stack -replace "`r`n|`n|`r", " | "
+        throw (New-EngineException -Func $func -Type "JSエラー" -Message "$($response.name): $($response.message)" -Details $stack)
     }
     return $response.data
 }
```

### Lib-WebAction_v303.ps1
* 【バグ修正】 Get-WebAttribute において、checked, disabled, readonly, selected の4属性は初期HTML属性（getAttribute）よりも先に現在のDOMプロパティ（el[a]）を評価するように変更し、チェックボックス操作後の状態取得が常に正しく "true" または "" を返すように改善します。また、XPath指定時にも同じ extractAttr ロジックを適用して挙動を統一します。 
* 【安定化】 Convert-BoundingBoxToScreenPoint の $dpr パースにおいて、カルチャ非依存の InvariantCulture を使用します。 

```Diff
@@ -31,7 +31,8 @@
 
     # 3. クライアント物理座標への変換： DPRを掛けて物理ピクセル座標へ変換する。
     $dprStr = Invoke-WebScript -Js "return window.devicePixelRatio || 1;"
-    $dpr = if ([double]::TryParse($dprStr, [ref]$null)) { [double]$dprStr } else { 1.0 }
+    $parsedDpr = 1.0
+    $dpr = if ([double]::TryParse([string]$dprStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedDpr)) { $parsedDpr } else { 1.0 }
     $clientX = [int][Math]::Round($clientLogicalX * $dpr)
     $clientY = [int][Math]::Round($clientLogicalY * $dpr)
```

```Diff
@@ -645,9 +646,9 @@
                 const frmSel = decodeURIComponent(escape(atob('$frameB64')));
                 const extractAttr = function(el, a) {
                     if (!el) return null;
+                    if (a === 'checked' || a === 'disabled' || a === 'readonly' || a === 'selected') return el[a] ? 'true' : null;
                     const val = el.getAttribute(a);
                     if (val !== null) return val;
-                    if (a === 'checked' || a === 'disabled' || a === 'readonly' || a === 'selected') return el[a] ? 'true' : null;
                     return null;
                 };
                 if (frmSel) {
```

```Diff
@@ -676,10 +677,16 @@
             $js = @"
                 const xpath = decodeURIComponent(escape(atob('$xpathB64')));
                 const attr = decodeURIComponent(escape(atob('$attrB64')));
+                const extractAttr = function(el, a) {
+                    if (!el) return null;
+                    if (a === 'checked' || a === 'disabled' || a === 'readonly' || a === 'selected') return el[a] ? 'true' : null;
+                    const val = el.getAttribute(a);
+                    return val !== null ? val : null;
+                };
                 return utilFindInFrames(window, function(win) {
                     try {
                         const node = win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
-                        return node ? node.getAttribute(attr) : null;
+                        return extractAttr(node, attr);
                     } catch(e) { return null; }
                 });
 "@
```

### Lib-WebXPath_v303.ps1
* 【安定化】 Normalize-XPath において、[contains(@class, 'btn') and text()='送信'] のように同一角括弧内に複数条件があるXPathが渡された場合に contains(@class, 'btn') and まで消えてしまうのを防ぐため、text()='xxx' または .='xxx' の条件式部分のみをピンポイントで contains(normalize-space(.), 'xxx') に置換するように改善します。 

```Diff
@@ -16,16 +16,16 @@
     }
 
     # 完全一致から部分一致への変換（テキスト検索）： [text()='xxx'] または [text()="xxx"] から、空白を無視した contains(normalize-space(.), 'xxx') へ変換する。
-    $pattern = "\[[^\]]*text\(\)\s*=\s*['`"]([^'`"]+)['`"]\]"
+    $pattern = "text\(\)\s*=\s*['`"]([^'`"]+)['`"]"
     if ($XPath -match $pattern) {
         $value = $Matches[1]
-        return ($XPath -replace $pattern, "[contains(normalize-space(.), '$value')]")
+        return ($XPath -replace $pattern, "contains(normalize-space(.), '$value')")
     }
 
     # 完全一致から部分一致への変換（ドット自身）： [.='xxx'] または [.="xxx"] から、contains(normalize-space(.), 'xxx') へ変換する。
-    $pattern2 = "\[[^\]]*\.\s*=\s*['`"]([^'`"]+)['`"]\]"
+    $pattern2 = "(?<=\[\s*|\band\s+|\bor\s+)\.\s*=\s*['`"]([^'`"]+)['`"]"
     if ($XPath -match $pattern2) {
         $value = $Matches[1]
-        return ($XPath -replace $pattern2, "[contains(normalize-space(.), '$value')]")
+        return ($XPath -replace $pattern2, "contains(normalize-space(.), '$value')")
     }
     return $XPath
 }
```

### Lib-DesktopUIA_v303.ps1
* 【規約・安定化】 Invoke-UiaAction の $Name パラメータに [AllowEmptyString()] を追加し、AutomationId 指定時に Name に空文字が渡されてもPowerShellのパラメータ検証エラーにならないように保護します。 
* 【規約統一】 $hwnd -eq $null / $hwnd -ne $null の比較を、他モジュールと同様に $null 左辺配置（$null -eq $hwnd）に統一します。

```Diff
@@ -13,7 +13,7 @@
      # ウィンドウの検索： 完全一致および部分一致によるウィンドウ検索を行う。
     $hWnd = [Win32Api.Win32Utils]::FindWindow($null, $Name)
-    if ($hWnd -eq [IntPtr]::Zero -or $hWnd -eq $null) {
+    if ($hWnd -eq [IntPtr]::Zero -or $null -eq $hWnd) {
         # FindWindowで完全一致取得できなかった場合、プロセス一覧からタイトルを部分一致検索してハンドルを取得する。
         $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$Name*" } | Select-Object -First 1
         if ($proc) { $hWnd = $proc.MainWindowHandle }
```

```Diff
@@ -82,7 +82,7 @@
             if ($editChild -and $editChild.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
                 $valuePattern.SetValue($Value)
             } else {
-                # それでもダメなら、エラーで落とさず自動でSafeモードへ移行する。
+                # 子要素でも非対応の場合は、エラーとせず自動でSafeモードへ移行する。
                 Write-DebugLog -Message "[$Func] 警告: ValuePattern非対応のため、Safeモード(物理入力)へ自動フォールバックします。" -Level Warning
                 try { 
                     if ($editChild) { $editChild.SetFocus() } else { $Element.SetFocus() } 
```

```Diff
@@ -129,7 +129,7 @@
 function Invoke-UiaAction {
     param (
         [Parameter(Mandatory=$true)][string]$Action,
-        [Parameter(Mandatory=$true)][string]$Name = "",
+        [Parameter(Mandatory=$true)][AllowEmptyString()][string]$Name = "",
         [string]$AutomationId = "",
         [string]$Value = "",
         [string]$Mode = "Pattern",
```

```Diff
@@ -211,18 +211,18 @@
             # パターンA： 標準のダイアログクラス名(#32770)で検索する。
             $hwnd = [Win32Api.Win32Utils]::FindWindow("#32770", $name)
             
-            if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $null) {
+            if ($hwnd -eq [IntPtr]::Zero -or $null -eq $hwnd) {
                 # パターンB： クラス名問わず、完全な null ポインタとして検索する。
                 $hwnd = [Win32Api.Win32Utils]::FindWindow([NullString]::Value, $name)
             }
 
-            if ($hwnd -ne [IntPtr]::Zero -and $hwnd -ne $null) { break }
+            if ($hwnd -ne [IntPtr]::Zero -and $null -ne $hwnd) { break }
         }
-        if ($hwnd -ne [IntPtr]::Zero -and $hwnd -ne $null) { break }
+        if ($hwnd -ne [IntPtr]::Zero -and $null -ne $hwnd) { break }
         Start-Sleep -Milliseconds 200
     }
 
-    if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $null) {
+    if ($hwnd -eq [IntPtr]::Zero -or $null -eq $hwnd) {
         throw (New-EngineException -Func $func -Type "未発見" -Message "保存ダイアログのウィンドウが見つかりませんでした")
     }
```

### Lib-DevTools_v303.ps1
* 【安定化】 Invoke-BrowserPointReverse（行552〜558）にて、$domData の $null チェックを $domData.log 参照より前に移動します。 
* 【規約統一】 Get-UiaWindowFromElement と Get-UiaTargetInfoFromCursor の $current -ne $null を $null -ne $current に統一します。

```Diff
@@ -549,13 +549,13 @@
     if (-not $resRaw) { return $null }
     # JSON文字列をPowerShellオブジェクトへ変換する。
     $domData = if ($resRaw -is [string]) { try { $resRaw | ConvertFrom-Json } catch { $null } } else { $resRaw }
+    if ($null -eq $domData) { return $null }
     
     if ($domData.log) {
         Write-DebugLog -Message "[PointReverse] Dpr=$($domData.log.dpr), Target(X=$($domData.log.targetX), Y=$($domData.log.targetY))" -Level Info
         Write-DebugLog -Message "[PointReverse] elementFromPoint => Tag=$($domData.log.elementTag), Id=$($domData.log.elementId), Name=$($domData.log.elementName)" -Level Info
     }
 
-    if ($null -eq $domData -or $domData.error) { return $null }
+    if ($domData.error) { return $null }
 
     # 3. 座標補正ロジック： JSで取得した座標はズレが生じやすいため、CDP(DevTools)で事前取得した正確なDOMスナップショットの座標で上書き補正する。
     if ($domData.boundingBox -and $global:DomNodes) {
```

```Diff
@@ -683,7 +683,7 @@
         $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
         $current = $TargetElement
         # 要素がnullになるか、タイプがWindow(ウィンドウそのもの)に到達するまでループする。
-        while ($current -ne $null -and $current.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
+        while ($null -ne $current -and $current.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
             # ひとつ上の親要素へ移動する。
             $current = $walker.GetParent($current)
         }
```

```Diff
@@ -718,7 +718,7 @@
         $depth = 0
 
         # 無限ループ防止のため最大20階層まで親を遡るループを実行する。
-        while ($current -ne $null -and $depth -lt 20) {
+        while ($null -ne $current -and $depth -lt 20) {
             # 現在評価中の要素のプロパティを取得する。
             $name = $current.Current.Name
             $id   = $current.Current.AutomationId
```

### Lib-WebDebug_v303.ps1
* 【安定化】 Export-WebTableToCsv（行229）において、$result がオブジェクト（正常時）の場合に不要な文字列変換と正規表現評価が走らないよう、$result -is [string] の型判定を先に行います。 

```Diff
@@ -225,9 +225,11 @@
          $result = Invoke-WebScript -Js $js
 
-        if ($result -eq 'NOT_FOUND')    { throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたテーブルが見つかりません" -Details $Selector) }
-        if ($result -eq 'EMPTY_TABLE')  { throw (New-EngineException -Func $func -Type "未発見" -Message "テーブル内に行データが存在しません" -Details $Selector) }
-        if ($result -match '^JS_EXCEPTION:') { throw (New-EngineException -Func $func -Type "JSエラー" -Message "テーブル解析中にJavaScript例外が発生しました" -Details $result) }
+        if ($result -is [string]) {
+            if ($result -eq 'NOT_FOUND')         { throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたテーブルが見つかりません" -Details $Selector) }
+            if ($result -eq 'EMPTY_TABLE')       { throw (New-EngineException -Func $func -Type "未発見" -Message "テーブル内に行データが存在しません" -Details $Selector) }
+            if ($result -match '^JS_EXCEPTION:') { throw (New-EngineException -Func $func -Type "JSエラー" -Message "テーブル解析中にJavaScript例外が発生しました" -Details $result) }
+        }
 
         if ($FileName -ne "") {
             $savePath = Join-Path $global:LogDir $FileName
```

```Diff
@@ -663,7 +665,7 @@
                         # リクエストした StyleNames の順序に合わせて文字列辞書から抽出する。
                         for ($styleIdx = 0; $styleIdx -lt $styleValues.Count -and $styleIdx -lt $StyleNames.Count; $styleIdx++) {
                             $strIdx = $styleValues[$styleIdx]
-                            # PowerShell特有の $null -ge 0 が True になる問題を防ぐ厳密チェック
+                            # インデックス値の $null および範囲外参照を防ぐ厳密チェックを行う。
                             if ($null -ne $strIdx -and $strIdx -ge 0 -and $strIdx -lt $strings.Count) {
                                 # 指定されたスタイル名をキーとして、辞書から引き当てた値を設定する。
                                 $targetNode.computedStyles[$StyleNames[$styleIdx]] = $strings[$strIdx]
```

---
### VBAモジュール
※ Ps_Engine.cls、Mod_TestDevTools.bas、UserForm_Picker.frm は現状のままで問題なく安定稼働しています。 
### Ps_Bridge.bas
* ParseRpaError の正規表現パターンを改善し、メッセージ本文中に括弧が含まれていても末尾の (Details) のみを正確に分離できるようにします。 

```Diff
@@ -42,7 +42,7 @@
     Set regEx = CreateObject("VBScript.RegExp")
     
     ' パターン定義（[関数名] [種別]: メッセージ (詳細)）： 詳細は省略される場合があるため、末尾の括弧はオプショナルとする。
-    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*([^\(]*?)(?:\s*\((.*)\))?$"
+    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.+?)(?:\s*\(([^\(\)]*)\))?$"
     regEx.IgnoreCase = True
     regEx.Global = False
```

### 仮想タブ（ポップアップ）でのサイレントダウンロード自動継承
ご提示いただいた Lib-WebAction_v303.ps1 の $global:SilentDownloadHandler 定義コードおよびコメントは、構文・設計ともに問題ありません。また、Lib-WebView2_Init_v303.ps1 側の自動継承処理もそのままで正常に機能します。 
ただし、公開サイトのダウンロードリンクに target="_blank" を付与してクリックした場合（今回のPart 22のテスト1）、Chromium（WebView2）の仕様上、以下の挙動となる点にご注意ください。
* Chromiumの仕様（別ウィンドウでの直接ファイルDL）：
`<a href="xxx.pdf" target="_blank">` のように、HTML画面を持たないファイル直リンクを別ウィンドウで開いた場合、ブラウザは新規ウィンドウ（仮想タブ）の生成をキャンセルするか、一瞬だけ空の仮想タブ（about:blank）を生成してダウンロードを親（または新規タブ）で開始し、その空タブは自動で閉じられずに残る（または親タブ側でイベントが発火する）場合があります。
* 確実な制御のポイント：
Lib-WebAction_v303.ps1 で $global:SilentDownloadRegistered = $true にする際、もし「すでに開いている別の仮想タブ」に切り替えた状態で Enable-SilentDownload が呼ばれたケースにも対応できるよう、念のため現在の $webview に対する登録確認を入れておくとより安全です（現状の「親画面P01で事前予約してからポップアップを開く」運用であれば、ご提示のコードのままで100%動作します）。 

#### Lib-WebAction_v303.ps1（ご提示いただいたコードの適用箇所）

```Diff
@@ -767,8 +767,9 @@
             return
         }
 
-        # new() を使わず、PowerShellの「型キャスト」を使って安全にイベントを登録する。
-        $handler = [System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs]] {
+        # イベントハンドラをグローバル変数へ保持し、新規仮想タブ生成時にも再利用可能にする。
+        # (ボタンを押すと別ウィンドウ（仮想タブ）が開いて即座に自動DLが始まる業務システムに対応させる)
+        $global:SilentDownloadHandler = [System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs]] {
             param($sender, $e)
 
             # バックグラウンドスレッドを守る。
```

```Diff
@@ -789,7 +790,7 @@
         }
 
         # 作成した安全なハンドラをセットする。
-        $webview.CoreWebView2.add_DownloadStarting($handler)
+        $webview.CoreWebView2.add_DownloadStarting($global:SilentDownloadHandler)
         $global:SilentDownloadRegistered = $true
```

#### Lib-WebView2_Init_v303.ps1（add_CoreWebView2InitializationCompleted 内）

```Diff
@@ -168,6 +168,13 @@
                             Write-DebugLog -Message "[WebView2] 警告: ポップアップへのJS自動注入に失敗" -Level Warning
                         }
 
+                        # 親画面でサイレントダウンロードが有効化されている場合、ポップアップ（仮想タブ）にも同じハンドラを自動継承する。
+                        if ($global:SilentDownloadRegistered -and $null -ne $global:SilentDownloadHandler) {
+                            try {
+                                $senderInit.CoreWebView2.add_DownloadStarting($global:SilentDownloadHandler)
+                                Write-DebugLog -Message "[WebView2] 情報: ポップアップ($tabId)へサイレントダウンロード設定を継承しました。" -Level Info
+                            } catch {}
+                        }
+
                         try {
                             # CDPを利用して、このポップアップ画面を直接操作・デバッグする際に必要となる内部ID(TargetId)を抽出して保持する。
                             $global:Tabs[$tabId].TargetId = $evtArgs.WindowFeatures.TargetId
```

※ 継承確認がログ（Engine_SystemLog.txt およびイミディエイトウィンドウ）で一目で分かるよう、Write-DebugLog を1行入れています。 

---
（少し微調整、、）完成としよう！
😊😊😊
