
#### 🔧 【v303 修正レポート ②】待機ループ内の `DoEvents()` 記述漏れ続き。と エラーログの微修正について

<p align="right">
  作成日: 2026年11月12日 / 作成者: System beginner & AI Collaborator (Gemini)
</p>

### 1. 待機・リトライループ内の [System.Windows.Forms.Application]::DoEvents() 記述漏れを、再度、全モジュールのループ等の総点検を指示した。
PowerShellを -STA（単一UIスレッド）で稼働させているため、ポーリング待機や通信リトライの Start-Sleep 前後に DoEvents() がないと、スリープ中にWinFormsのメッセージポンプおよびWebView2の非同期イベント（仮想タブ生成・window.close 破棄の BeginInvoke・ナビゲーション完了通知）が一時停止します。 

#### ① Lib-WebCDP_v303.ps1（1箇所：Connect-CdpSession / 102行目付近）
* CDPエンドポイント（/json）からアクティブタブのターゲットを探すリトライループ（500ms × 最大10回 ＝ 最大5秒）に DoEvents() がありません。新規ポップアップ（仮想タブ）が開いた直後に Connect-CdpSession が呼ばれた際、ここで DoEvents() を回さないとWebView2側の新タブ初期化やURL更新（NavigationCompleted）イベントが進行せず、URL一致に失敗して別タブへフォールバック接続するリスクがあります。 
```diff
             if ($pageTarget) { break }
         } catch {}

+        [System.Windows.Forms.Application]::DoEvents()
         Start-Sleep -Milliseconds 500
     }
```

#### ② Lib-WebView2_Native_v303.ps1（3箇所：Invoke-WebView2NativeScript / 88・119・142行目付近）
*	ページ遷移直後などでJSの戻り値が空・デコード失敗となった際、指数バックオフ（200ms ➔ 400ms ➔ 800ms... 最大2000ms）で Start-Sleep -Milliseconds $delayMs を実行していますが、その直前に DoEvents() がありません。画面遷移の過渡期で無応答になっている時こそ、DoEvents() を回してWebView2の描画・ロードを進行させる必要があります。 
```diff
                 # 1. 戻り値が空またはnull時のリトライ (88行目付近)
                 if ($attempt -lt $Retries) {
+                    [System.Windows.Forms.Application]::DoEvents()
                     Start-Sleep -Milliseconds $delayMs
                     # 次回の待機時間を2倍に設定する。(上限2000ミリ秒の指数バックオフ)

                 # 2. JSONデコード失敗時のリトライ (119行目付近)
                 if ($attempt -lt $Retries) {
+                    [System.Windows.Forms.Application]::DoEvents()
                     Start-Sleep -Milliseconds $delayMs
                     $delayMs = [Math]::Min(2000, $delayMs * 2)

             # 3. 予期せぬ例外発生時のリトライ (142行目付近)
             if ($attempt -lt $Retries) {
+                [System.Windows.Forms.Application]::DoEvents()
                 Start-Sleep -Milliseconds $delayMs
                 $delayMs = [Math]::Min(2000, $delayMs * 2)
```

#### ③ 補助箇所：Ps_Engine_Core_v303.ps1（378行目）および Lib-DevTools_v303.ps1（1009行目）
*	現象： Ps_Engine_Core_v303.ps1 起動時のCDPポート確認ループ（500ms × 6回）、および Lib-DevTools_v303.ps1 の Get-UiaElementFromPoint の要素取得リトライループ（200ms × 15回 ＝ 最大3秒） にも DoEvents() を入れておくと、待機中の画面固まりを防げます。 
```diff
         # Ps_Engine_Core_v303.ps1 (378行目付近)
         for ($i = 0; $i -lt 6; $i++) {
+            [System.Windows.Forms.Application]::DoEvents()
             Start-Sleep -Milliseconds 500

         # Lib-DevTools_v303.ps1 Get-UiaElementFromPoint内 (1009行目付近)
+            [System.Windows.Forms.Application]::DoEvents()
             Start-Sleep -Milliseconds $waitMs
         }
```

### 2. VBA側 Ps_Bridge.bas：ParseRpaError の正規表現バグ（Details が空になる問題）
エラーログで、詳細(Details) が出力されない。
```text
PS_STDOUT: [ERROR][Invoke-WebXPathClick] [未発見]: XPath要素のクリック実行に失敗しました (//a[contains(normalize-space(.),'079186.csv')])
【実行中エラー】 関数: Invoke-WebXPathClick / 種別: 未発見
 メッセージ: XPath要素のクリック実行に失敗しました (//a[contains(normalize-space(.),'079186.csv')])
 詳細(Details):  ★★ここ
```

* 原因： Ps_Bridge.bas（49行目）の正規表現パターンが以下の記述になっています。 regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.+?)(?:\s*\(([^\(\)]*)\))?$" 末尾の Details キャプチャ部が [^\(\)]*（内部に丸括弧 ( ) を一切含まない）となっているため、$Details に渡されたXPath（contains(...)）やCSSセレクタ（:nth-of-type(1)）、あるいは Wait-WebXPathElementDisappear の 30秒経過 (//*[@id=...]) のように丸括弧が1つでも含まれるとマッチに失敗し、rpaErr.Details が常に空文字になり、rpaErr.Message 側に丸ごと吸収されてしまいます。 
•	修正コード（Ps_Bridge.bas 49行目付近）： ([^\(\)]*) を (.*) に変更することで、XPath等の丸括弧を含む詳細情報も正確に Details フィールドへ分離されます。
```diff
     ' パターン定義（[関数名] [種別]: メッセージ (詳細)）： 詳細は省略される場合があるため、末尾の括弧はオプショナルとする。
-    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.+?)(?:\s*\(([^\(\)]*)\))?$"
+    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.+?)(?:\s*\((.*)\))?$"
     regEx.IgnoreCase = True
```
     
### 3. Wait-Condition 呼び出し時の -CallerFunc $func 渡し漏れ
Lib-WebCustomUI_v303.ps1 の関数では、Wait-Condition で、-CallerFunc $func を適用したが、従来モジュールの再点検をお願い。
```text
<xx> <Info> [Engine] 実行: Wait-WebUrlContains | Params: {"Substring":"dummy_not_found","TimeoutSec":3}
<xx> <Error> [Engine] コマンド実行エラー: [Wait-Condition] [Timeout]: [Wait-WebUrlContains] タイムアウト: URL不一致 (dummy_not_found) | トレース: Wait-Condition、以下（略）
```

#### ① Lib-WebAction_v303.ps1（6箇所：111, 176, 204, 253, 303, 385行目付近）
```diff
     # 1. Wait-WebPageLoad (111行目付近)
-    $errMsg = "[$func] タイムアウト: ページまたは iframe のロード未完了"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: ページまたは iframe のロード未完了"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg -CallerFunc $func | Out-Null

     # 2. Wait-WebUrlContains (176行目付近)
-    $errMsg = "[$func] タイムアウト: URL不一致 ($Substring)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: URL不一致 ($Substring)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg -CallerFunc $func | Out-Null

     # 3. Wait-WebTitleContains (204行目付近)
-    $errMsg = "[$func] タイムアウト: タイトル不一致 ($Substring)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: タイトル不一致 ($Substring)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg -CallerFunc $func | Out-Null

     # 4. Wait-WebElement (253行目付近)
-    $errMsg = "[$func] タイムアウト: 要素の未出現または非表示 ($Selector)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: 要素の未出現または非表示 ($Selector)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg -CallerFunc $func | Out-Null

     # 5. Wait-WebElementInFrame (303行目付近)
-    $errMsg = "[$func] タイムアウト: iframe内要素の未出現 ($ElementSelector)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: iframe内要素の未出現 ($ElementSelector)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg -CallerFunc $func | Out-Null

     # 6. Wait-WebElementInvisible (385行目付近)
-    $errMsg = "[$func] タイムアウト: 要素が非表示になりません ($Selector)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: 要素が非表示になりません ($Selector)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg -CallerFunc $func | Out-Null
```

#### ② Lib-WebSafeAction_v303.ps1（1箇所：Invoke-WebSafeClick / 178行目付近）
```diff
-    $errMsg = "[$func] タイムアウト: 要素の出現・可視化、またはクリックに失敗しました ($Selector)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: 要素の出現・可視化、またはクリックに失敗しました ($Selector)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg -CallerFunc $func | Out-Null
```

#### ③ Lib-WebXPath_v303.ps1（1箇所：Wait-WebXPathElement / 102行目付近）
```diff
-    $errMsg = "[$func] タイムアウト: XPath要素が見つからない ($XPath)"
-    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
+    $errMsg = "タイムアウト: XPath要素が見つからない ($XPath)"
+    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg -CallerFunc $func | Out-Null
```
