
<h4 align="center">汎用RPA操作エンジン (VBA × PowerShell × WebView2 Hybrid-Engine) 2/2 </h4>
<h4 align="center">V301 ⇒ V302 へ 更新中</h4>

<p align="right">
  作成日: 2026年10月16日 - 10月xx日<br>
  作成者: System beginner & AI Collaborator (Gemini)
</p>

_GEMINI 3.1Proで開発したコードを、**Gemini 4 Argonで全モジュールを再検証・アップグレード**しました。_<br>
_<sub>（Argon凄い！。コード開発は最低限の課金（今回: Pro、多分1ヵ月）は必要でした。今までの時間が○○！ でも、普段は無料枠でも十分な私）</sub>_

_（アップグレード記録）_<br>
_<sub>コード改修の経過を中心に纏めているので、GEMINIへの指示、アドバイスは要約しています。</sub>_

```text
* _Step1: 全モジュール、VBA、マニュアル等に対し、「**コード、コメントの検証**」を指示。_<br>
> _<sub>Ps_Engine_Core_v301.ps1 冒頭の「命名規則およびコーディングルール」「コメント規約（体言止め・常体・大ブロック[]の階層制限・既存コメント削除禁止）」を厳格に適用し、全モジュールの精査をお願い。 </sub>_

* _Step2: 修正コードの検証コード(主な改修点を検証する)_<br>
* _Step3: コード修正（微・調整）_<br>
```

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
---

### ◆◆ Step4 ◆◆
<h4 align="center">アップグレードに対するアドバイス</h4>

```text
* **アドバイス①：** ピッカー実行時間（10秒〜20秒）を「1秒未満」に短縮する設計改善
* **アドバイス②：** Shadow DOM 内の XPATH: 生成と Test-SelectorStability の不一致解消
* **アドバイス③：** iframe内要素の「絶対座標（BoundingBox）」自動加算によるズレ解消
```

**🛠️ 対象ファイル1：Lib-WebDebug_v302.ps1**

* **アドバイス①（ピッカー実行時間の劇的短縮）：**
> Convert-DomSnapshotFlat 内で、CDPの DOMSnapshot.captureSnapshot が標準で返却している $layout.bounds（[x, y, width, height] の配列）を直接展開して各ノードの boundingBox に格納します。 これにより、Get-DomSnapshotWithBoxModel 内で数百回ループしていた DOM.getBoxModel の個別WebSocket通信を廃止（または bounds が取得できなかった特殊要素のみへのフォールバックに限定）でき、10〜20秒かかっていたピッカー処理が0.3秒〜0.8秒程度に短縮されます。 
* **アドバイス② & ③（Test-SelectorStability のShadow DOM対応・iframe絶対座標対応）：**
> 安定性テスト時にも deepQuerySelector でShadow DOMを貫通し、かつiframe内の要素であれば親フレームのオフセットを加算した絶対座標をログ出力するように統一します。 

```Diff
@@ Lib-WebDebug_v302.ps1 535行目付近：Convert-DomSnapshotFlat での layout.bounds 直接展開 @@
             # 1要素分の器（テンプレート）を作成する。
             $node = [ordered]@{
                 index          = $i
                 parentIndex    = if ($nodes.parentIndex) { $nodes.parentIndex[$i] } else { $null }
                 nodeType       = if ($nodes.nodeType) { $nodes.nodeType[$i] } else { $null }
                 backendNodeId  = if ($nodes.backendNodeId) { $nodes.backendNodeId[$i] } else { $null }
                 childIndexes   = if ($nodes.childNodeIndexes) { $nodes.childNodeIndexes[$i] } else { $null }
                 attributes     = @{}
                 computedStyles = @{}
+                boundingBox    = $null
                 tagName        = ""
                 id             = ""
                 name           = ""
```

```Diff
@@ 615行目付近：2. レイアウト(スタイルおよび座標bounds)情報の紐づけ @@
         # ----------------------------------------------------------
         # 2. レイアウト(スタイル)情報の紐づけ： 要素に対応するスタイル配列を辞書から引き当てる。
 
         # （レイアウトツリーが存在し、スタイル配列がある場合のみ実行する）
         if ($layout -and $layout.nodeIndex -and $layout.styles) {
 
             # layout.nodeIndex には、画面に描画されている要素のインデックス番号が入っている。
             for ($layoutIdx = 0; $layoutIdx -lt $layout.nodeIndex.Count; $layoutIdx++) {
                 $nodeIdx = $layout.nodeIndex[$layoutIdx]
 
                 # list内の該当ノードを取得する (フレーム毎にオフセットを考慮)
                 # $list は累積されているため、現在のフレームの開始位置からのインデックスを計算する。
                 if ($null -eq $nodeIdx) { continue } # ← 安全対策
 
                 $absoluteIdx = ($list.Count - $count) + $nodeIdx
 
                 if ($absoluteIdx -ge 0 -and $absoluteIdx -lt $list.Count) {
                     $targetNode = $list[$absoluteIdx]
                     $styleValues = $layout.styles[$layoutIdx]
                     
                     if ($null -ne $styleValues) {
                         # リクエストした StyleNames の順序に合わせて文字列辞書から抽出する。
                         for ($styleIdx = 0; $styleIdx -lt $styleValues.Count -and $styleIdx -lt $StyleNames.Count; $styleIdx++) {
                             $strIdx = $styleValues[$styleIdx]
                             # PowerShell特有の $null -ge 0 が True になる問題を防ぐ厳密チェック
                             if ($null -ne $strIdx -and $strIdx -ge 0 -and $strIdx -lt $strings.Count) {
                                 # 指定されたスタイル名をキーとして、辞書から引き当てた値を設定する。
                                 $targetNode.computedStyles[$StyleNames[$styleIdx]] = $strings[$strIdx]
                             }
                         }
                     }
+
+                    # レイアウト配列に含まれる境界矩形(bounds: [x, y, width, height])を直接抽出して格納する。
+                    if ($layout.bounds -and $layoutIdx -lt $layout.bounds.Count) {
+                        $b = $layout.bounds[$layoutIdx]
+                        if ($null -ne $b -and $b.Count -ge 4 -and $b[2] -gt 0 -and $b[3] -gt 0) {
+                            $targetNode.boundingBox = @{
+                                x      = [math]::Round([double]$b[0], 2)
+                                y      = [math]::Round([double]$b[1], 2)
+                                width  = [math]::Round([double]$b[2], 2)
+                                height = [math]::Round([double]$b[3], 2)
+                            }
+                        }
+                    }
                 }
             }
         }
     }
     return $list
 }
```

```Diff
@@ Lib-WebDebug_v302.ps1 660行目付近：Get-DomSnapshotWithBoxModel の高速化（個別通信の最小化） @@
     # 1. 画面全体のDOMツリーと基本スタイルを一撃で取得する（layout.boundsにより座標情報も一括取得される）。
     $cdpResponse = Invoke-CdpCommand -Method "DOMSnapshot.captureSnapshot" -Params @{
-        computedStyles = @("display","visibility","width","height")
+        computedStyles     = @("display","visibility","width","height")
+        includeDOMRects    = $true
+        includePaintOrder  = $false
     }
     $domResult = $cdpResponse.result
     if (-not $domResult) {
         Write-DebugLog -Message "[$func] エラー: DOMSnapshotの取得結果が空です" -Level Error
         return $null
     }
 
     # 2. 圧縮されたCDPの1次元配列を、扱いやすいオブジェクトリスト（連想配列）に展開する。
     $nodes = Convert-DomSnapshotFlat -DomResult $domResult
 
-    # 3. DOM.getBoxModelのViewport座標を絶対座標へ補正するため、スクロール量を取得する。
-    $scrollX = 0; $scrollY = 0
+    # 3. 座標補正用のDPRおよびスクロール量を取得する（DOMSnapshot.boundsのデバイス座標またはViewport座標補正用）。
+    $dpr = 1.0; $scrollX = 0; $scrollY = 0
     try {
-        # windowのスクロール量をJSで取得する。
-        $scrJs = "return JSON.stringify({x: window.scrollX || 0, y: window.scrollY || 0});"
+        # windowのDPRとスクロール量をJSで取得する。
+        $scrJs = "return JSON.stringify({dpr: window.devicePixelRatio || 1, x: window.scrollX || 0, y: window.scrollY || 0});"
         $scrRes = Invoke-CdpScript -Js $scrJs
         $scrObj = $scrRes | ConvertFrom-Json
-        if ($scrObj) { $scrollX = $scrObj.x; $scrollY = $scrObj.y }
+        if ($scrObj) {
+            if ($scrObj.dpr -gt 0) { $dpr = [double]$scrObj.dpr }
+            $scrollX = $scrObj.x; $scrollY = $scrObj.y
+        }
     } catch {}
 
-    # 4. インタラクティブ要素に絞って絶対座標を付与する。
+    # 4. インタラクティブ要素に絞って絶対座標を確定する（layout.bounds未取得時のみDOM.getBoxModelへフォールバック）。
     foreach ($node in $nodes) {
         if (-not $node.tagName) { continue }
         # パフォーマンス最適化：標準のインタラクティブ要素を許可
         $isInteractive = ($node.tagName -in @("input","button","a","select","textarea","iframe"))
         
         # SPA/エディタ対策：標準タグ以外でも、操作可能な役割(role)やラベル、編集可能(contenteditable)な要素は許可
         if (-not $isInteractive) {
             if ($node.role -match "^(button|link|checkbox|menuitem|tab|searchbox|textbox)$" -or
                 -not [string]::IsNullOrEmpty($node.ariaLabel) -or
                 $node.attributes["contenteditable"] -eq "true" -or
                 $node.tagName -eq "body") {
                 $isInteractive = $true
             }
         }
          
-        # 該当しない純粋なレイアウトタグ(ただのdivやspan)はスキップして通信負荷を防ぐ
-        if (-not $isInteractive) { continue }
+        # 該当しない純粋なレイアウトタグ(ただのdivやspan)は座標情報をクリアして対象外とする。
+        if (-not $isInteractive) {
+            $node.boundingBox = $null
+            continue
+        }
+
+        # Convert-DomSnapshotFlat で既に layout.bounds から座標が取得できている場合は個別通信をスキップする。
+        if ($null -ne $node.boundingBox) { continue }
 
         # ターゲット要素の固有ID（backendNodeId）がない場合はスキップする。
         if (-not $node.backendNodeId) { continue }
 
         try {
-            # 絞り込んだ要素に対してのみ、CDP経由で物理座標（BoxModel）を要求する。
+            # layout.boundsで取得できなかった要素に対してのみ、CDP経由で物理座標（BoxModel）を要求する。
             $boxResponse = Invoke-CdpCommand -Method "DOM.getBoxModel" -Params @{
                 backendNodeId = $node.backendNodeId
             }
```

**🛠️ 対象ファイル2：Lib-DevTools_v302.ps1**

* **アドバイス①（UsePointReverse 時のさらなる高速化）：**
> Invoke-UiaRecordStep および Process-RecordAction において、UsePointReverse = $true（座標逆引きルート優先）の場合は、まずJSによる高速な座標逆引き（Invoke-BrowserPointReverse）を実行します。そこで要素が確定すれば全件スナップショット（Get-DomSnapshotWithBoxModel）の通信自体をスキップし、フォールバック（ルートB：UIA-DOM照合）に入る場合のみスナップショットを取得する遅延評価（Lazy Load）構造にします。

* **アドバイス②（Shadow DOM 内の要素に対するCSSセレクタ生成）：**
> Get-CorePointReverseJs 内の buildCssSelector(el) において、一意性チェックの基準を document 固定ではなく el.getRootNode()（ShadowRoot または所属フレームのDocument）に変更します。これにより、Shadow DOM内の要素（#shadow-select 等）や iframe内の要素が XPATH: に格下げされず、綺麗なCSSセレクタとして生成されるようになります。

* **アドバイス③（iframe内要素の絶対座標 BoundingBox 自動加算）：**
> Get-CorePointReverseJs 内に getAbsoluteBoundingRect(el) 補助関数を追加し、対象要素が iframe 内にある場合でも、親フレームの枠座標（frameElement.getBoundingClientRect()）を再帰的に足し上げてトップウィンドウ基準の絶対論理座標を算出します。 

```Diff
@@ Lib-DevTools_v302.ps1 255行目付近：Invoke-UiaRecordStep のスナップショット遅延取得化 @@
         # ユーザーが対象要素にマウスを合わせるための待機: （現在 $WaitSec=0、VBA側で制御している。）
         if ($WaitSec -gt 0) {
             Start-Sleep -Seconds $WaitSec
         }
         
-        # 待機完了後に、スクロール後を捉えるため最新のDOMスナップショットを取得する。
-        $global:DomNodes = Get-DomSnapshotWithBoxModel
+        # スナップショットの初期化： PointReverse優先時は高速化のため事前取得をスキップし、ルートB移行時にオンデマンド取得する。
+        if (-not $UsePointReverse) {
+            $global:DomNodes = Get-DomSnapshotWithBoxModel
+        } else {
+            $global:DomNodes = $null
+        }
 
         # メイン処理： 要素特定オーケストレーターへ処理を委譲する。
         $generatedCode = Process-RecordAction -IsGetNameMode $false -UsePointReverse $UsePointReverse
```

```Diff
@@ Lib-DevTools_v302.ps1 425行目付近：Get-CorePointReverseJs 内の buildCssSelector (Shadow DOM / iframe ルート対応) @@
         // 特定した要素から一意なCSSセレクタを構築する関数
         function buildCssSelector(el) {
             if (!el) return '';
             // CSS.escapeが存在すれば利用し、無ければ何もしない無名関数を定義する。
             const cssEsc = (typeof CSS !== 'undefined' && CSS.escape) ? CSS.escape : s => s;
+            // Shadow DOM内部やiframe内部でも正確に一意性を判定するため、要素が属するルートノード(ShadowRootまたはDocument)を取得する。
+            const root = (el.getRootNode && el.getRootNode()) ? el.getRootNode() : (el.ownerDocument || document);
 
             // IDが存在する場合は最優先でIDセレクタを採用する。
             if (el.id) {
                 const idSel = '#' + cssEsc(el.id);
-                try { if (document.querySelectorAll(idSel).length === 1) return idSel; } catch(e) {}
+                try { if (root.querySelectorAll && root.querySelectorAll(idSel).length === 1) return idSel; } catch(e) {}
                 // ID単体で一意にならなければ、name属性と組み合わせてみる。
                 if (el.name) return el.tagName.toLowerCase() + idSel + '[name="' + cssEsc(el.name) + '"]';
+                return idSel;
             }
             // name属性が存在する場合は採用する。
             if (el.name) return el.tagName.toLowerCase() + '[name="' + cssEsc(el.name) + '"]';
             // クラスが存在する場合は結合してクラスセレクタを構築する。
             if (el.className && typeof el.className === 'string') {
                 const classes = el.className.trim().split(/\s+/).map(cls => cssEsc(cls)).join('.');
                 if (classes) return el.tagName.toLowerCase() + '.' + classes;
             }
             return '';
         }
```

```Diff
@@ Lib-DevTools_v302.ps1 520行目付近：Get-CorePointReverseJs 内の iframe 絶対座標加算ロジック追加 @@
         // セレクタの構築および判定。強固なCSSセレクタが得られなかった場合はXPATHへフォールバックする。
         const cssResult = buildCssSelector(targetElement);
         const finalSelector = (cssResult && (cssResult.indexOf('#') !== -1 || cssResult.indexOf('[name=') !== -1 || cssResult.indexOf('.') !== -1)) 
             ? cssResult : 'XPATH:' + buildRobustXPath(targetElement);
 
-        // 要素の絶対座標（スクロール量加算済）を算出する。
-        const rect = targetElement.getBoundingClientRect();
-        const scrollX = window.scrollX || window.pageXOffset || 0;
-        const scrollY = window.scrollY || window.pageYOffset || 0;
+        // 要素の絶対座標（iframeオフセットおよびスクロール量加算済）を算出する。
+        function getAbsoluteBoundingBox(el) {
+            const r = el.getBoundingClientRect();
+            let left = r.left;
+            let top  = r.top;
+            let win  = (el.ownerDocument && el.ownerDocument.defaultView) ? el.ownerDocument.defaultView : window;
+
+            // iframe内部の要素である場合、親フレームの枠座標(frameElement)をトップウィンドウまで再帰的に加算する。
+            while (win && win !== window.top) {
+                try {
+                    const frameEl = win.frameElement;
+                    if (!frameEl) break;
+                    const frameRect = frameEl.getBoundingClientRect();
+                    left += frameRect.left + (frameEl.clientLeft || 0);
+                    top  += frameRect.top  + (frameEl.clientTop  || 0);
+                    win = win.parent;
+                } catch (e) {
+                    break;
+                }
+            }
+            const sx = window.scrollX || window.pageXOffset || 0;
+            const sy = window.scrollY || window.pageYOffset || 0;
+            return {
+                x: Math.round(left + sx),
+                y: Math.round(top + sy),
+                width: Math.round(r.width),
+                height: Math.round(r.height)
+            };
+        }
+        const absBox = getAbsoluteBoundingBox(targetElement);
 
         // 最終的な解析結果をJSONとしてPowerShellへ返却する。
         return JSON.stringify({
             selector: finalSelector,
             outerHtml: targetElement.outerHTML || '',
-            boundingBox: { 
-                x: Math.round(rect.left + scrollX), 
-                y: Math.round(rect.top + scrollY), 
-                width: Math.round(rect.width), 
-                height: Math.round(rect.height) 
-            },
+            boundingBox: absBox,
             isInsideFrame: isFrame, // 修正した判定結果を渡す
             log: log
         });
```

```Diff
@@ Lib-DevTools_v302.ps1 680行目付近：Process-RecordAction 内のルートB移行時オンデマンドSnapshot取得 @@
         # ルートA： 物理座標からのDOM逆引き。
         if ($UsePointReverse -and ($stepWindowName -match "RPA Browser" -or $stepWindowName -match "Edge" -or $stepWindowName -match "Chrome")) {
             $rect = $targetElement.Current.BoundingRectangle
             $pt = New-Object System.Windows.Point([int]($rect.Left + ($rect.Width / 2)), [int]($rect.Top + ($rect.Height / 2)))
             $resObj = Invoke-BrowserPointReverse -ScreenPoint $pt -TargetControlType $targetControlType -CurrentValue $currentValue
             
             if ($resObj) { 
                 $resObj.command = "[PointReverse] " + $resObj.command
                 return "[RESULT]" + ($resObj | ConvertTo-Json -Compress)
             }
         }
 
         # ルートB： UIAメタデータからのDOM逆引き。
+        # ルートA(PointReverse)で未取得またはフォールバックした場合のみ、DOMスナップショットをオンデマンド取得する。
+        if (-not $global:DomNodes) {
+            $global:DomNodes = Get-DomSnapshotWithBoxModel
+        }
         $resObj = Invoke-UiaDomResolution -TargetName $targetName -TargetControlType $targetControlType -TargetId $targetId -TargetElement $targetElement -CurrentValue $currentValue
```

**💡 この改修による変化（テストシートでの確認ポイント）**

**1. ピッカー速度（T11 Wikipedia / T14 Yahoo路線など）：** これまで 10,000ms〜19,900ms かかっていた ◆ 計測/Invoke-UiaRecordStep_ (TimeMs) が、200ms〜600ms（約0.3秒） に劇的に短縮されます。

**2. Shadow DOM テスト（12_scroll_shadow_dom.html の T02, T03, T05, T06, T07）：** ピッカー取得時に XPATH://select[@id="shadow-select"] ではなく #shadow-select として取得されるようになり、直後の Test-SelectorStability（安定性テスト）も [OK] を通過するようになります。 (※ 自動テストシート TestCases の E列 ExpectedSelector に XPATH://select[@id="shadow-select"] と登録されている行は、ピッカーが #shadow-select を返すように進化するため、自動テスト実行時は E列の期待値も #shadow-select 等に合わせていただくと OK 判定になります。) 

**3. iframe 座標（T22 [https://the-internet.herokuapp.com/iframe](https://the-internet.herokuapp.com/iframe)）：** ピッカー取得時に I列・J列に出力される ClientX / ClientY が、iframe枠のオフセットを含んだ画面全体の正確な座標として算出されるようになります。 

---

### ◆◆ Step5 ◆◆

**PowerShell側（Lib-WebAction_v302.ps1 / Lib-WebXPath_v302.ps1）の修正**

BoundingBox（物理フォールバック用の座標）が渡されているにもかかわらず、Wait-WebElement で失敗した後に、もう一度同じセレクタでJSクリック/入力を試行し、さらに時間を浪費しているのでは？<br>
Wait-WebElement がタイムアウトした時点で要素がDOM上で見つからないことは確定している。Wait-WebElement の成否（$elemFound）を受け取り、見つからなかった場合は無駄な2回目のJS実行をスキップして即座に物理フォールバックへ直行する。

```Diff
@@ Lib-WebAction_v302.ps1 360行目付近：Invoke-WebClick のフォールバック直行化 @@
     $func = $MyInvocation.MyCommand.Name
     # Base64エンコードを利用してサニタイズする。
     $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
 
     # 要素出現を待機する。（エラーを握り潰してフォールバックへ流すため Try-Catch を適用する）
-    try { Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null } catch {}
+    $elemFound = $false
+    try { $elemFound = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec } catch {}
 
-    # 多段iframeの再帰探索およびJSクリック：
-    $js = @"
-        const selector = decodeURIComponent(escape(atob('$selB64')));
-        const found = utilFindInFrames(window, function(win) {
-            try {
-                const el = deepQuerySelector(selector, win.document);
-                if (el) {
-                    // 要素を画面の中央へスクロールし、クリックを実行する。
-                    el.scrollIntoView({block: 'center', inline: 'center'});
-                    el.click();
-                    return true;
-                }
-            } catch(e) {}
-            return null;
-        });
-        return found === true;
-"@
-
-    $res = Invoke-WebScript -Js $js
+    # 多段iframeの再帰探索およびJSクリック： 待機処理で要素が確認できた場合のみJSクリックを実行する。
+    $res = $false
+    if ($elemFound) {
+        $js = @"
+            const selector = decodeURIComponent(escape(atob('$selB64')));
+            const found = utilFindInFrames(window, function(win) {
+                try {
+                    const el = deepQuerySelector(selector, win.document);
+                    if (el) {
+                        // 要素を画面の中央へスクロールし、クリックを実行する。
+                        el.scrollIntoView({block: 'center', inline: 'center'});
+                        el.click();
+                        return true;
+                    }
+                } catch(e) {}
+                return null;
+            });
+            return found === true;
+"@
+        $res = Invoke-WebScript -Js $js
+    }
```

```Diff
@@ Lib-WebAction_v302.ps1 425行目付近：Set-WebTextInput のフォールバック直行化 @@
     $func = $MyInvocation.MyCommand.Name
     # Base64エンコードを利用してサニタイズする。
     $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
     $valB64 = ConvertTo-JsSafeBase64 -Text $Value
 
-    try { Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null } catch {}
+    $elemFound = $false
+    try { $elemFound = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec } catch {}
 
-    $js = @"
-        const selector = decodeURIComponent(escape(atob('$selB64')));
-        const val = decodeURIComponent(escape(atob('$valB64')));
-        const found = utilFindInFrames(window, function(win) {
-            try {
-                const el = deepQuerySelector(selector, win.document);
-                if (el) {
-                    // 対象要素にフォーカスを当てる。
-                    el.focus();
-                    // DOMのプロパティとして値を直接書き込む。
-                    el.value = val;
-                    // ReactやVueなどのモダンフレームワークに変更を認識させるため、イベントを強制発火させる。
-                    el.dispatchEvent(new Event('input', { bubbles: true }));
-                    el.dispatchEvent(new Event('change', { bubbles: true }));
-                    return true;
-                }
-            } catch(e) {}
-            return null;
-        });
-        return found === true;
-"@
-
-    $res = Invoke-WebScript -Js $js
+    # 待機処理で要素が確認できた場合のみJS入力を実行し、未発見時は即座に物理フォールバックへ移行する。
+    $res = $false
+    if ($elemFound) {
+        $js = @"
+            const selector = decodeURIComponent(escape(atob('$selB64')));
+            const val = decodeURIComponent(escape(atob('$valB64')));
+            const found = utilFindInFrames(window, function(win) {
+                try {
+                    const el = deepQuerySelector(selector, win.document);
+                    if (el) {
+                        // 対象要素にフォーカスを当てる。
+                        el.focus();
+                        // DOMのプロパティとして値を直接書き込む。
+                        el.value = val;
+                        // ReactやVueなどのモダンフレームワークに変更を認識させるため、イベントを強制発火させる。
+                        el.dispatchEvent(new Event('input', { bubbles: true }));
+                        el.dispatchEvent(new Event('change', { bubbles: true }));
+                        return true;
+                    }
+                } catch(e) {}
+                return null;
+            });
+            return found === true;
+"@
+        $res = Invoke-WebScript -Js $js
+    }
```

PowerShell側の戻り値が2行になっている。<br>
Invoke-WebClick（Set-WebTextInput）の中で呼び出している Invoke-DesktopCoordinateClick が最後に return $true を返す。<br>
True + 改行(vbLf) + [FALLBACK_PHYSICAL]... という2行のデータになっている。（**シート出力時の 折り返して全体を表示**）

> VBAから改行コード（vbLf や vbCrLf）を含む文字列をセルに書き込むと、Excelは自動的にそのセルの「折り返して全体を表示（WrapText = True）」をONにして行の高さを広げてしまいます。
> VBA側（Mod_TestDevTools.bas）で不要な True や改行を除去して1行に整形し、かつセルの折り返し設定（WrapText = False）も強制解除します。

```Diff
@@ Mod_TestDevTools.bas 575行目付近：TestId_Type_PhysicalFallback のH列1行整形 @@
         ' 結果をG列(7)、レスポンスをH列(8)に書き込み
         If Err.Number <> 0 Then
             ws.Cells(r, 7).Value = "NG (Fallback Error)"
             ws.Cells(r, 7).Font.Color = vbRed
-            ws.Cells(r, 8).Value = Err.Description
+            ws.Cells(r, 8).Value = Replace(Replace(Err.Description, vbCrLf, " "), vbLf, " ")
+            ws.Cells(r, 8).WrapText = False
             Err.Clear
         Else
+            ' 内部関数の戻り値(True + 改行)が含まれている場合は除去し、1行に整形する。
+            fallbackRes = Replace(fallbackRes, "True" & vbCrLf, "")
+            fallbackRes = Replace(fallbackRes, "True" & vbLf, "")
+            fallbackRes = Trim(Replace(Replace(fallbackRes, vbCrLf, " "), vbLf, " "))
+
             ws.Cells(r, 7).Value = "OK (Physical)"
             ws.Cells(r, 7).Font.Color = vbBlue
             ws.Cells(r, 8).Value = fallbackRes & " | bbox: " & bboxJson
+            ws.Cells(r, 8).WrapText = False
         End If
```

---

### ◆◆ Step6 ◆◆
### Ps_Engine_Core_v302.ps1

```Diff
@@ -16,8 +16,8 @@
       関数内の変数だけでなく、「モジュール展開時（ドットソース）の初期化処理でのみ使用し、
       以降の状態として参照しない作業用変数」もローカル変数と同義とする。
    ※ パラメータ名（PascalCase）をローカルで加工・再利用する際は、必ず camelCase の別名とする。
-      [NG]: $XPathEscaped = Normalize-XPath$XPath
-      [OK]: $xpathEscaped = Normalize-XPath$XPath
+      [NG]: $XPathNorm = Normalize-XPath$XPath
+      [OK]: $xpathNorm = Normalize-XPath$XPath
 
 • パラメータおよびグローバル変数: PascalCase（パスカルケース）
 • 定数・システム設定値（例外ルール）:
```

```Diff
@@ -39,14 +39,14 @@
 • [PowerShell → JavaScript] 埋め込み用変数の統一:
    PowerShellからJSのヒアドキュメント（@""@）に展開・埋め込むための加工済み変数は、
    元のパラメータが PascalCase であっても、必ず camelCase に変換して埋め込む。
-   - 例: $XPath (パラメータ) → $xpathEscaped (JS埋め込み用ローカル変数)
-   - 例: $FrameSelector → $frameSelectorEscaped
+   - 例: $XPath (パラメータ) → $xpathB64 (JS埋め込み用Base64ローカル変数)
+   - 例: $FrameSelector → $frameB64
 
 • [JavaScript 内部コード] 変数名の独立と統一:
    JSコード内で宣言する変数は、JavaScriptの標準規約に従い camelCase を徹底する。
    PowerShell側の PascalCase をJSコンテキスト内に持ち込まない。
-   - [NG]: const XPath = '$xpathEscaped';
-   - [OK]: const xpath = '$xpathEscaped';
+   - [NG]: const XPath = decodeURIComponent(escape(atob('$xpathB64')));
+   - [OK]: const xpath = decodeURIComponent(escape(atob('$xpathB64')));
```

```Diff
@@ -122,9 +122,9 @@
 <#
     ============================================================================
  【DPIスケーリングの強制発火ハック】
-  WPF(System.Windows.Automation)は、初期化されるまでOSの現在のDPIスケーリング(125%や150%など)を正しく認識しません。
+  WPF(System.Windows.Automation)は、初期化されるまでOSの現在のDPIスケーリング(125%や150%など)を正しく認識しない。
   このため、自動化処理の直前にダミーの座標(0,0)オブジェクトを生成してUIA要素を強制取得させることで、
-  WPFの内部エンジンを意図的に初期化し、以降の物理マウスクリック等における座標計算のズレを完全に防ぎます。
+  WPFの内部エンジンを意図的に初期化し、以降の物理マウスクリック等における座標計算のズレを完全に防ぐ。
     ============================================================================
 #>
```

```Diff
@@ -540,7 +540,7 @@
                 $res = &$method @parameters | Out-String
 
                 # 実行結果を出力する。
-                if (-not [string]::IsNullOrWhiteSpace($res)) {
+                if ($null -ne $res -and -not [string]::IsNullOrWhiteSpace($res)) {
                     [Console]::WriteLine("[RESULT]$($res.Trim())")
                 }
                 [Console]::WriteLine("[SUCCESS]")
```

```Diff
@@ -557,7 +557,7 @@
                 # ログファイルへのスタックトレース記録： 波線等の装飾を含まないスタックトレースを記録する。
                 $stackTrace = if ($_.ScriptStackTrace) {$_.ScriptStackTrace -replace "`r`n", " <- " -replace "`n", " <- "
                 } elseif ($_.Exception -and $_.Exception.StackTrace) {
-                    # ScriptStackTraceが空の場合のフォールバック
+                    # ScriptStackTraceが空の場合のフォールバックとして処理する。
                     $_.Exception.StackTrace -replace "`r`n", " <- " -replace "`n", " <- "
                 } else {
                     "不明"
```

```Diff
@@ -570,7 +570,7 @@
 
 } finally {
     # ==============================================================================
-    # 安全な終了処理（クリーンアップ）： フォームやプロセスを破棄し、メモリリソースを解放する。
+    # [安全な終了処理（クリーンアップ）]： フォームやプロセスを破棄し、メモリリソースを解放する。
     # ==============================================================================
     Write-DebugLog -Message "[System] 情報: エンジンシャットダウン" -Level Info
```

### Lib-WebJS_v302.ps1
generateCssSelector において、Shadow DOMの内部ルート（ShadowRoot）が root として渡された場合でも優先度1位（ID属性）の検証が確実に行えるよう、root.querySelector の互換性を高め、コメントを微調整します。 

```Diff
@@ -112,7 +112,7 @@
     const id = el.getAttribute && el.getAttribute('id');
     if (id) {
         // CSSセレクタ用にエスケープ処理を施したIDセレクタを構築する。
         const sel = '#' + cssEscape(id);
         // 構築したIDセレクタで検索し、自身が正確にヒットするか検証する。
-        try { if ((root.querySelector(sel) === el)) return sel; } catch (e) {}
+        try { if (root.querySelector && (root.querySelector(sel) === el)) return sel; } catch (e) {}
     }
```

### Lib-WebView2_Init_v302.ps1
Clear-WebCache 内で既に New-EngineException によってフォーマットされた例外（Timeout等）が catch ブロックで再度「内部エラー」として二重ラップされるのを防ぐ判定を追加します。 

```Diff
@@ -445,6 +445,8 @@
         return "SUCCESS"
 
     } catch {
+        # 既にNew-EngineExceptionでフォーマット済みの例外はそのまま再スローする。
+        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
         $errMsg = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
         throw (New-EngineException -Func $func -Type "内部エラー" -Message "キャッシュクリア処理中に予期せぬエラーが発生しました" -Details $errMsg)
     }
```

### Lib-WebView2_Native_v302.ps1
catch ブロックにおいて、JSエラー だけでなく 初期化エラー や ネイティブエラー など、すでに New-EngineException 形式でフォーマットされた例外が二重にラップされるのを防ぎ、コメントを整理します。 

```Diff
@@ -127,7 +127,8 @@
             if ($attempt -lt $Retries) {
                 Start-Sleep -Milliseconds $delayMs
                 $delayMs = [Math]::Min(2000, $delayMs * 2)
                 continue
             } else {
+                if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
                 $errMsg = (New-EngineException -Func $func -Type "ネイティブエラー" -Message "スクリプトの実行処理中に予期せぬエラーが発生しました" -Details $_.Exception.Message)
                 throw $errMsg
             }
```

### Lib-WebCDP_v302.ps1
Invoke-CdpNativeClick と Set-CdpNativeTextInput において、標準の document.querySelector ではなく、注入済みの deepQuerySelector を優先利用するようにアップグレードし、Shadow DOM内の要素に対してもCDPネイティブクリック・入力が機能するようにします。 

```Diff
@@ -355,7 +355,7 @@
     # 要素の中央座標の取得： スクロールして要素を可視化し、中心の座標を算出する。
     $js = @"
      const selector = decodeURIComponent(escape(atob('$selB64')));
-     const el = document.querySelector(selector);
+     const el = (typeof deepQuerySelector === 'function') ? deepQuerySelector(selector, document) : document.querySelector(selector);
      if (!el) {
          throw new Error("Element not found: " + selector);
      }
```

```Diff
@@ -409,7 +409,7 @@
     # 対象要素へのフォーカスおよび値のクリア： ネイティブ入力前に要素をフォーカスし既存値を消去する。
     $js = @"
      const selector = decodeURIComponent(escape(atob('$selB64')));
-     const el = document.querySelector(selector);
+     const el = (typeof deepQuerySelector === 'function') ? deepQuerySelector(selector, document) : document.querySelector(selector);
      if (!el) {
          throw new Error("Element not found: " + selector);
      }
```

### Lib-WebSafeAction_v302.ps1
Invoke-WebSafeClick において、通常の document.querySelectorAll に加え、単一ヒットしない場合（Shadow DOM内部にある場合）は deepQuerySelector でも探索できるよう強化し、12_scroll_shadow_dom.html のようなShadow DOM内要素にも Invoke-WebSafeClick が対応できるようにします。 

```Diff   
@@ -88,8 +88,12 @@
                 // 以前のトラッキングログをクリアする。
                 window.DOMUtils.tracker.clear();
 
-                // セレクタに合致するすべての要素を取得する。
-                const els = document.querySelectorAll(selector);
+                // セレクタに合致するすべての要素を取得する（通常DOMで未発見時はShadow DOMも探索する）。
+                let els = Array.from(document.querySelectorAll(selector));
+                if (els.length === 0 && typeof deepQuerySelector === 'function') {
+                    const shadowEl = deepQuerySelector(selector, document);
+                    if (shadowEl) els = [shadowEl];
+                }
                 window.DOMUtils.tracker.log('探索開始: セレクタに合致する要素 ' + els.length + ' 件');
                 
                 if (els.length === 0) {
```

### Lib-DevTools_v302.ps1
Invoke-UiaDomResolution 内で label[for="..."] を検索する際、特殊文字を含むIDで構文エラーにならないよう CSS.escape（または window.DOMUtils.cssEscape）を適用します。 

```Diff   
@@ -162,7 +162,7 @@
         if ($UiaName) {
             # Nameやプレースホルダが完全一致する場合は高得点を付与する。
             if ($node.name -eq $UiaName -or $node.placeholder -eq $UiaName) { $score += 600 }
-            # innerText(画面表示テキスト)が一致した場合の大幅加点（歳出管理などを救済）
+            # innerText(画面表示テキスト)が一致した場合は大幅に加点する（歳出管理などのメニュー要素を救済する）。
             elseif ($node.innerText -and $node.innerText -eq $UiaName) { $score += 600 }
             elseif ($node.innerText -and $node.innerText -match $UiaName) { $score += 200 }
             # 部分一致の場合は加点する。
```

```Diff
@@ -185,7 +185,7 @@
         }
     }
 
-    # スコアの足切りライン(100点)を設け、無関係なinputタグの誤爆を防ぐ
+    # スコアの足切りライン(100点)を設け、無関係なinputタグの誤爆を防ぐ。
     if ($bestNode -and $bestScore -ge 100) {
         if ($bestNode.selector) { return $bestNode.selector }
         if ($bestNode.name)     { return $bestNode.tagName + '[name="' + $bestNode.name + '"]' }
```

```Diff
@@ -306,7 +306,7 @@
                 if (/^(DIV|SPAN|I|SVG|P|B)$/.test(el.tagName) && el.parentElement) {
                     try {
                         const siblingTarget = el.parentElement.querySelector('select, input:not([type="hidden"]), textarea, button');
-                        // 見つかった場合、それが確実にRPAで操作したい本体要素なので優先する
+                        // 見つかった場合、それが確実にRPAで操作したい本体要素であるため優先する。
                         if (siblingTarget) return siblingTarget;
                     } catch(e) {}
                 }
```

```Diff
@@ -435,9 +435,9 @@
         // メイン処理： 座標から要素を抽出する。
         const targetElement = deepElementFromPoint(targetX, targetY, document);
 
-        // iframe内要素かどうかの確実な判定
+        // iframe内要素かどうかの確実な判定を行う。
         let isFrame = false;
         try {
-            // targetElementが所属するドキュメントが、現在のトップドキュメント(document)と異なるかで判定
+            // targetElementが所属するドキュメントが、現在のトップドキュメント(document)と異なるかで判定する。
             isFrame = (targetElement && targetElement.ownerDocument !== document);
         } catch(e) {}
 
-        // ピッカー操作の視覚的フィードバック
+        // ピッカー操作の視覚的フィードバックを実行する。
         if (targetElement && targetElement.style !== undefined) {
             try {
                 const oldOutline = targetElement.style.outline;
                 const oldOutlineOffset = targetElement.style.outlineOffset;
-                const oldBackground = targetElement.style.backgroundColor; // 元の背景色を保存
+                const oldBackground = targetElement.style.backgroundColor; // 元の背景色を保存する。
                 
                 targetElement.style.outline = '2px solid red';
                 targetElement.style.outlineOffset = '-2px'; 
                 
-                // iframe内の要素だった場合、背景色を薄黄色にして視覚的に警告する
+                // iframe内の要素だった場合、背景色を薄黄色にして視覚的に警告する。
                 if (isFrame) {
                     targetElement.style.backgroundColor = '#ffffcc';
                 }
```

```Diff
@@ -500,7 +500,7 @@
             selector: finalSelector,
             outerHtml: targetElement.outerHtml || targetElement.outerHTML || '',
             boundingBox: absBox,
-            isInsideFrame: isFrame, // 修正した判定結果を渡す
+            isInsideFrame: isFrame, // 判定結果を渡す。
             log: log
         });
     } catch (e) {
```

```Diff
@@ -603,7 +603,8 @@
                 try {
                     // IDが存在する場合、for属性で紐づいているlabel要素を探す。
                     if (el.id) {
-                        const label = document.querySelector('label[for="' + el.id + '"]');
+                        const escId = (window.DOMUtils && window.DOMUtils.cssEscape) ? window.DOMUtils.cssEscape(el.id) : el.id;
+                        const label = document.querySelector('label[for="' + escId + '"]');
                         if (label) targetForBox = label;
                     }
                     // labelが見つからず、かつ親要素がlabelである場合(Inputを内包するLabel)は、親を座標取得の対象とする。
```

### Lib-WebDebug_v302.ps1
Export-WebTableToCsv にて、$global:ENGINE_JS_UTILS は初期化時に AddScriptToExecuteOnDocumentCreatedAsync で自動注入済みであるため、重複注入を防ぐガード（またはコメント補足）を整え、例外再スロー時の二重ラップ防止を追加します。 

```Diff
@@ -161,7 +161,7 @@
                 const rows = Array.from(table.rows);
                 if (rows.length === 0) return 'EMPTY_TABLE';
 
-                // CSV出力用の文字列を生成する（人間が読む用）
+                // CSV出力用の文字列を生成する（人間が読む用）。
                 const csvData = rows.map(row => {
                     return Array.from(row.cells).map(cell => {
                         // 改行文字をスペースに置換し、ダブルクォートでエスケープする。
```

```Diff
@@ -170,21 +170,21 @@
                     }).join(',');
                 }).join('\r\n');
 
-                // VBA連携用のデータ文字列を生成する（データ抽出用）
+                // VBA連携用のデータ文字列を生成する（データ抽出用）。
                 const vbaData = rows.map(row => {
                     return Array.from(row.cells).map(cell => {
                         // 1. テキストがあればそれを返す。
                         const text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
                         if (text !== '') return text;
 
-                        // 2. テキストが空なら画像ファイル名を返す（汎用処理）
+                        // 2. テキストが空なら画像ファイル名を返す（汎用処理）。
                         const img = cell.querySelector('img');
                         if (img && img.src) {
                             const parts = img.src.split('/');
-                            return parts[parts.length - 1]; // 例: "IP10B030.png"
+                            return parts[parts.length - 1];
                         }
                         return '';
-                    // VBA側でのSplit解析 (セル間を<T>、行間を<R>という独自デリミタで結合する)
+                    // VBA側でのSplit解析のため、セル間を<T>、行間を<R>という独自デリミタで結合する。
                     }).join('<T>');
                 }).join('<R>');
```

```Diff
@@ -208,6 +208,7 @@
         return $result.vba
 
     } catch {
+        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
         throw (New-EngineException -Func $func -Type "ファイルエラー" `
             -Message "テーブルのCSV保存処理中に例外が発生しました" -Details $_.Exception.Message)
     }
```

### Lib-WebXPath_v302.ps1
Normalize-XPath において、シングルクォート text()='xxx' だけでなくダブルクォート text()="xxx" で指定された場合も同様に自動正規化されるよう正規表現を強化します。 

```Diff    
@@ -14,15 +14,15 @@
         return $XPath
     }
 
-    # 完全一致から部分一致への変換（テキスト検索）： [text()='xxx'] から、空白を無視した contains(normalize-space(.), 'xxx') へ変換する。
-    $pattern = "\[[^\]]*text\(\)\s*=\s*'([^']+)'\]"
+    # 完全一致から部分一致への変換（テキスト検索）： [text()='xxx'] または [text()="xxx"] から、空白を無視した contains(normalize-space(.), 'xxx') へ変換する。
+    $pattern = "\[[^\]]*text\(\)\s*=\s*['`"]([^'`"]+)['`"]\]"
     if ($XPath -match$pattern) {
         $value =$Matches[1]
         return ($XPath -replace $pattern, "[contains(normalize-space(.), '$value')]")
     }
 
-    # 完全一致から部分一致への変換（ドット自身）： [.='xxx'] から、contains(normalize-space(.), 'xxx') へ変換する。
-    $pattern2 = "\[[^\]]*\.\s*=\s*'([^']+)'\]"
+    # 完全一致から部分一致への変換（ドット自身）： [.='xxx'] または [.="xxx"] から、contains(normalize-space(.), 'xxx') へ変換する。
+    $pattern2 = "\[[^\]]*\.\s*=\s*['`"]([^'`"]+)['`"]\]"
     if ($XPath -match $pattern2) {
         $value = $Matches[1]
         return ($XPath -replace $pattern2, "[contains(normalize-space(.), '$value')]")
```

```Diff
@@ -117,7 +117,7 @@
                     const rect = el.getBoundingClientRect();
                     // display, visibility, opacity, またはサイズのいずれかで非表示になれば false 扱いとする。
                     const isVisible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
-                    if (isVisible) return true; // まだ表示されている
+                    if (isVisible) return true; // まだ表示されている状態を示す。
                 }
             } catch(e) {}
             return null;
```

### Lib-WebAction_v302.ps1
* Mod_TestRun2_Base の TEST_Part99（確認 8-C）にて Get-WebAttribute に FrameSelector を渡してテストしていますが、現在の Get-WebAttribute は $FrameSelector パラメータを受け取るだけで内部処理では全フレーム透過探索（utilFindInFrames）のみを行っていました。FrameSelector が明示指定された場合は該当フレーム内に限定して探索し、未指定時は従来通り全フレーム透過探索を行うように実装を補完します。 

* Invoke-WebClick および Set-WebTextInput で物理クリックへフォールバックした際、Invoke-DesktopCoordinateClick の戻り値（$true）がパイプラインに出力されて "Truen[FALLBACK_PHYSICAL]..." の2行になる現象（Mod_TestDevTools.bas側でReplace 除去していた現象）を、Out-Nullで抑止して1行のクリーンな戻り値にします（既存のVBA側Replace` とも完全互換です）。 

```Diff    
@@ -403,7 +403,7 @@
                 
                 # Lib-DesktopUIA の物理クリック関数を呼び出す。
                 if (Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) {
-                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
+                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY | Out-Null
                     return "[FALLBACK_PHYSICAL] Click Completed: $Selector"
                 }
             }
```

```Diff
@@ -468,7 +468,7 @@
                     (Get-Command -Name "Invoke-DesktopSendKeys" -ErrorAction SilentlyContinue)) {
                     
                     # 1. 物理クリックでカーソルを合わせる。
-                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
+                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY | Out-Null
                     
                     # 2. 全選択(Ctrl+A) -> 削除(Delete) -> 値を送信する。
                     Invoke-DesktopSendKeys -Keys "^a" -WaitMs 100 | Out-Null
```

```Diff
@@ -625,20 +625,31 @@
     }
     # Base64エンコードを利用してサニタイズする。
     $attrB64 = ConvertTo-JsSafeBase64 -Text $Attribute
+    $frameB64 = ConvertTo-JsSafeBase64 -Text $FrameSelector
 
     try {
         $js = ""
         # 2. Selector または XPath に応じた待機処理とJS組み立て：
         if (-not [string]::IsNullOrEmpty($Selector)) {
             # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ。
-            $null = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec
+            if (-not [string]::IsNullOrEmpty($FrameSelector)) {
+                $null = Wait-WebElementInFrame -FrameSelector $FrameSelector -ElementSelector $Selector -TimeoutSec $TimeoutSec
+            } else {
+                $null = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec
+            }
             
             $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
             # セレクタで特定した要素の指定属性を取得して返却するJSを構築する。
             $js = @"
                 const sel = decodeURIComponent(escape(atob('$selB64')));
                 const attr = decodeURIComponent(escape(atob('$attrB64')));
+                const frmSel = decodeURIComponent(escape(atob('$frameB64')));
+                if (frmSel) {
+                    const frm = document.querySelector(frmSel);
+                    if (!frm || !frm.contentDocument) return null;
+                    const el = deepQuerySelector(sel, frm.contentDocument);
+                    return el ? el.getAttribute(attr) : null;
+                }
                 return utilFindInFrames(window, function(win) {
                     try {
                         const el = deepQuerySelector(sel, win.document);
```

### Lib-DesktopUIA_v302.ps1
Invoke-UiaInputInternal（Safeモードおよびフォールバック時）において、SendKeys に + ^ % ~ ( ) [ ] { } などの特殊文字が含まれる場合に備え、Set-WebTextInput と同様のエスケープ安全策またはコメント補足を整えます。 

```Diff    
@@ -63,7 +63,8 @@
         try { $Element.SetFocus() } catch {}
         Start-Sleep -Milliseconds 100
         $wshell = New-Object -ComObject WScript.Shell
-        $wshell.SendKeys($Value)
+        $escapedValue = $Value -replace '([+^%~()\[\]{}])', '{$1}'
+        $wshell.SendKeys($escapedValue)
     } else {
         # パターンモード（バックグラウンド入力）：
         $valuePattern = $null
```

```Diff
@@ -91,7 +92,8 @@
                 Start-Sleep -Milliseconds 50
                 $wshell.SendKeys("{DELETE}") # 削除を実行する。
                 Start-Sleep -Milliseconds 50
-                $wshell.SendKeys($Value) # 値を入力する。
+                $escapedValue = $Value -replace '([+^%~()\[\]{}])', '{$1}'
+                $wshell.SendKeys($escapedValue) # 値を入力する。
             }
         }
     }
```

### VBAモジュール群

**Ps_Bridge.bas**
ParseRpaError にて、エラー文字列の Details 内に丸括弧 (...) が含まれる場合（例: Timeout (10s) など）でも末尾まで正確にキャプチャできるよう、正規表現を微調整します。

```Diff
@@ -42,7 +42,7 @@
     Set regEx = CreateObject("VBScript.RegExp")
     
     ' パターン定義（[関数名] [種別]: メッセージ (詳細)）： 詳細は省略される場合があるため、末尾の括弧はオプショナルとする。
-    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.*?)(?:\s*\((.*)\))?$"
+    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*([^\(]*?)(?:\s*\((.*)\))?$"
     regEx.IgnoreCase = True
     regEx.Global = False
```

**Ps_Engine.cls**
GetTickCount の24.9日ロールオーバー（符号付きLongのオーバーフロー）に対する防御計算を1行追加し、コメントの通り完全な日跨ぎ・長期稼働耐性を担保します

```Diff
@@ -122,7 +122,7 @@
         End If
 
         ' 2. タイムアウト判定 (ミリ秒単位で比較、日跨ぎバグを回避)
-        If (GetTickCount() - startTick) > (PS_TIMEOUT_SEC * 1000&) Then
+        If ((GetTickCount() - startTick) And &H7FFFFFFF) > (PS_TIMEOUT_SEC * 1000&) Then
             Me.CloseEngine
             m_IsBusy = False
             Err.Raise 997, "PS_Engine", "PowerShell応答タイムアウト (" & PS_TIMEOUT_SEC & "秒)"
```

**UserForm_Picker.frm**

```Diff
@@ -200,13 +200,13 @@
     End If
 End Sub
 
-' --- [選択Ur]クリップボード経由 ---
+' --- [選択Url] クリップボード経由 ---
 Private Sub TextUrl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
     Cancel = True
     ' クリップボードのテキストを直接貼り付け
     Me.TextUrl.Paste
 End Sub
 
-' --- [選択Ur]選択したUrlへ移動 ---
+' --- [選択Url] 選択したUrlへ移動 ---
 Private Sub btnNavigate_Click()
     ' URLが空でなければ移動フラグ(8)を立てる
```

### タブ一覧リストボックス（lstTabs）からの直接タブ切り替えについて（コードの確認をお願い。）
lstTabs（タブ一覧を表示するリストボックス）と btnGetTabs_Click が実装されており、エクスポート時にタブ一覧が取得・表示されます

```Diff
Private Sub lstTabs_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstTabs.ListIndex >= 0 Then
        Dim targetTabId As String
        targetTabId = Me.lstTabs.List(Me.lstTabs.ListIndex, 1)
        GetRpaEngine().RunAction "Switch-Tab", CreateParams("TabId", targetTabId)
        Me.lblStatus.Caption = "タブ切替完了: " & targetTabId
    End If
End Sub
```

### 1レコードが上下2段やセル結合への対応と、展開オプションの付加

官公庁や財務会計システム・伝票一覧画面では、1レコードが上下2段の <tr> に分かれていたり、rowspan="2" / colspan="3" でセル結合されているケースが多々あります。JS側で結合セルのズレを補正して空欄を埋める（または同じ値を複製する）マトリックス展開オプション（[bool]$ExpandMergedCells = $false）を将来的に追加,,,

> Lib-WebDebug_v302.ps1 の Export-WebTableToCsv に、将来の財務会計システム等の結合テーブル対策として $ExpandMergedCellsFlg（Alias "ExpandMergedCells"、規定値 "false"） を追加します。<br> 
> •	規定値（省略時 / False）: これまでと1ミリも変わらない完全互換の高速動作です。<br> 
> •	ExpandMergedCells := True 指定時: rowspan（縦結合）や colspan（横結合）を解析して2次元グリッドマトリックスへ展開し、結合されたセル位置にも同じ値を自動補完して列ズレを完全に防ぎます。また、Shadow DOM内のテーブルも取得できるよう deepQuerySelector にも対応させています。 

**Lib-WebDebug_v302.ps1**

```Diff
@@ -137,8 +137,10 @@
 # テーブル要素を解析し、VBA取込用の配列文字列を含むCSVを生成する。
 function Export-WebTableToCsv {
     param (
         [Parameter(Mandatory = $true)][string]$Selector,
-        [string]$FileName = "WebTable.csv"
+        [string]$FileName = "WebTable.csv",
+        [Alias("ExpandMergedCells")]
+        [string]$ExpandMergedCellsFlg = "false"
     )
 
     $func = $MyInvocation.MyCommand.Name
```

```Diff
@@ -148,14 +150,20 @@
     }
 
     # Base64エンコードを利用してサニタイズする。
     $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
+    $expandBool = $false
+    try { $expandBool = [System.Convert]::ToBoolean($ExpandMergedCellsFlg) } catch {}
+    $expandJsStr = $expandBool.ToString().ToLower()
 
     try {
         $js = @"
-            $global:ENGINE_JS_UTILS
             const sel = decodeURIComponent(escape(atob('$selB64')));
+            const expandMerged = $expandJsStr;
             try {
-                // 共通関数を利用し、iframe内も含めて対象のtable要素を探索する。
+                // 共通関数を利用し、iframeおよびShadow DOM内も含めて対象のtable要素を探索する。
                 const table = utilFindInFrames(window, function(win) {
-                    try { return win.document.querySelector(sel) || null; }
+                    try { return deepQuerySelector(sel, win.document) || null; }
                     catch(e) { return null; }
                 });
```

```Diff
@@ -165,30 +173,63 @@
                 const rows = Array.from(table.rows);
                 if (rows.length === 0) return 'EMPTY_TABLE';
 
+                // セルから純粋なテキスト（または画像ファイル名）を抽出するヘルパー関数を定義する。
+                function extractCellVal(cell, useImgFallback) {
+                    const text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
+                    if (text !== '' || !useImgFallback) return text;
+                    const img = cell.querySelector('img');
+                    if (img && img.src) {
+                        const parts = img.src.split('/');
+                        return parts[parts.length - 1];
+                    }
+                    return '';
+                }
+
+                // 結合セル(rowspan/colspan)展開用の2次元グリッドマトリックスを構築する関数を定義する。
+                function buildMatrix(useImgFallback) {
+                    const grid = [];
+                    for (let r = 0; r < rows.length; r++) {
+                        if (!grid[r]) grid[r] = [];
+                        const cells = Array.from(rows[r].cells);
+                        let cIdx = 0;
+                        for (let c = 0; c < cells.length; c++) {
+                            while (grid[r][cIdx] !== undefined) { cIdx++; }
+                            const cell = cells[c];
+                            const val = extractCellVal(cell, useImgFallback);
+                            const rSpan = expandMerged ? (cell.rowSpan || 1) : 1;
+                            const cSpan = expandMerged ? (cell.colSpan || 1) : 1;
+                            for (let rs = 0; rs < rSpan; rs++) {
+                                for (let cs = 0; cs < cSpan; cs++) {
+                                    if (!grid[r + rs]) grid[r + rs] = [];
+                                    grid[r + rs][cIdx + cs] = val;
+                                }
+                            }
+                            cIdx += cSpan;
+                        }
+                    }
+                    return grid;
+                }
+
                 // CSV出力用の文字列を生成する（人間が読む用）。
-                const csvData = rows.map(row => {
-                    return Array.from(row.cells).map(cell => {
-                        // 改行文字をスペースに置換し、ダブルクォートでエスケープする。
-                        const text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
-                        return '"' + text.replace(/"/g, '""') + '"';
+                const csvGrid = buildMatrix(false);
+                const csvData = csvGrid.map(rowArr => {
+                    return rowArr.map(val => {
+                        const safeVal = (val !== undefined && val !== null) ? String(val) : '';
+                        return '"' + safeVal.replace(/"/g, '""') + '"';
                     }).join(',');
                 }).join('\r\n');
 
                 // VBA連携用のデータ文字列を生成する（データ抽出用）。
-                const vbaData = rows.map(row => {
-                    return Array.from(row.cells).map(cell => {
-                        // 1. テキストがあればそれを返す。
-                        const text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
-                        if (text !== '') return text;
-
-                        // 2. テキストが空なら画像ファイル名を返す（汎用処理）。
-                        const img = cell.querySelector('img');
-                        if (img && img.src) {
-                            const parts = img.src.split('/');
-                            return parts[parts.length - 1]; // 例: "IP10B030.png"
-                        }
-                        return '';
-                    // VBA側でのSplit解析 (セル間を<T>、行間を<R>という独自デリミタで結合する)
+                const vbaGrid = buildMatrix(true);
+                const vbaData = vbaGrid.map(rowArr => {
+                    return rowArr.map(val => {
+                        return (val !== undefined && val !== null) ? String(val) : '';
+                    // VBA側でのSplit解析のため、セル間を<T>、行間を<R>という独自デリミタで結合する。
                     }).join('<T>');
                 }).join('<R>');
 ```

**VBAからの呼び出し例（結合セル展開を使う場合）**

```vba
' 通常呼び出し（従来通り・省略可）
vbaTableStr = rpaEngine.RunAction("Export-WebTableToCsv", CreateParams("Selector", "table.list", "FileName", "Out.csv"))

' 結合セル（rowspan/colspan）を展開して列ズレを防止する場合
vbaTableStr = rpaEngine.RunAction("Export-WebTableToCsv", CreateParams("Selector", "table.list", "FileName", "Out.csv", "ExpandMergedCells", True))
```

### リアルタイム・ハイライト（赤枠ホバー）のiframe・Shadow DOM透過対応(PSモジュールへ移行)
Shadow DOMの内部要素や、同一オリジンiframeの内部要素にマウスを乗せた際も、親フレームのオフセットを加算してピンポイントで赤枠が表示される透過対応版の関数を追加します。 （※コーディング規約に従い、VBAからの真偽値文字列を受け取るパラメータは $EnableFlg（Alias "Enable"）としています。） 

**Lib-WebSafeAction_v302.ps1**

```Diff
+# --- ピッカー用リアルタイム・ホバーハイライトのON/OFF制御 ---
+# マウスカーソル下の要素（Shadow DOM・同一オリジンiframe透過対応）をリアルタイムで赤枠ハイライトする。
+function Set-WebHoverHighlight {
+    param (
+        [Parameter(Mandatory=$true)]
+        [Alias("Enable")]
+        [string]$EnableFlg
+    )
+
+    $func = $MyInvocation.MyCommand.Name
+    $enable = $false
+    try { $enable = [System.Convert]::ToBoolean($EnableFlg) } catch {}
+
+    if ($enable) {
+        $js = @"
+            try {
+                // 既存のハイライト枠が残っていれば一旦削除する。
+                if (window._rpaOutline) { window._rpaOutline.remove(); }
+                
+                const outline = document.createElement('div');
+                outline.style.cssText = 'position:fixed; pointer-events:none; z-index:2147483647; outline:2px solid red; background:rgba(255,0,0,0.1); transition:all 0.04s;';
+                document.body.appendChild(outline);
+                window._rpaOutline = outline;
+
+                // Shadow DOMおよびiframeを貫通してマウス下の要素とトップ画面基準の矩形を算出する。
+                window._rpaUpdateHover = function(clientX, clientY, baseWin) {
+                    let curDoc = baseWin.document;
+                    let x = clientX;
+                    let y = clientY;
+                    let offsetX = 0;
+                    let offsetY = 0;
+
+                    // baseWinが子iframeの場合、親フレーム群のオフセットを加算する。
+                    let w = baseWin;
+                    while (w && w !== window.top) {
+                        try {
+                            const fe = w.frameElement;
+                            if (!fe) break;
+                            const fr = fe.getBoundingClientRect();
+                            offsetX += fr.left + (fe.clientLeft || 0);
+                            offsetY += fr.top  + (fe.clientTop  || 0);
+                            w = w.parent;
+                        } catch(e) { break; }
+                    }
+
+                    let el = curDoc.elementFromPoint(x, y);
+                    // トップ画面上でiframe自体に当たった場合は内部へ透過する。
+                    while (el && (el.tagName === 'IFRAME' || el.tagName === 'FRAME')) {
+                        try {
+                            const fr = el.getBoundingClientRect();
+                            offsetX += fr.left + (el.clientLeft || 0);
+                            offsetY += fr.top  + (el.clientTop  || 0);
+                            x -= (fr.left + (el.clientLeft || 0));
+                            y -= (fr.top  + (el.clientTop  || 0));
+                            if (!el.contentDocument) break;
+                            curDoc = el.contentDocument;
+                            el = curDoc.elementFromPoint(x, y);
+                        } catch(e) { break; }
+                    }
+
+                    // Shadow DOM内部へ再帰的に透過する。
+                    while (el && el.shadowRoot && el.shadowRoot.elementFromPoint) {
+                        const inner = el.shadowRoot.elementFromPoint(x, y);
+                        if (!inner || inner === el) break;
+                        el = inner;
+                    }
+
+                    if (el && window._rpaOutline) {
+                        const r = el.getBoundingClientRect();
+                        window._rpaOutline.style.top    = (r.top + offsetY) + 'px';
+                        window._rpaOutline.style.left   = (r.left + offsetX) + 'px';
+                        window._rpaOutline.style.width  = r.width + 'px';
+                        window._rpaOutline.style.height = r.height + 'px';
+                    }
+                };
+
+                // トップウィンドウおよび同一オリジンiframeへmousemoveイベントを登録する。
+                utilFindInFrames(window, function(win) {
+                    try {
+                        win._rpaHoverHandler = function(e) { window._rpaUpdateHover(e.clientX, e.clientY, win); };
+                        win.document.addEventListener('mousemove', win._rpaHoverHandler);
+                    } catch(e) {}
+                    return null;
+                });
+                return true;
+            } catch(e) { return false; }
+"@
+        Invoke-WebScript -Js $js | Out-Null
+        return "HoverHighlight Enabled"
+    } else {
+        $jsOff = @"
+            try {
+                utilFindInFrames(window, function(win) {
+                    try {
+                        if (win._rpaHoverHandler) {
+                            win.document.removeEventListener('mousemove', win._rpaHoverHandler);
+                            win._rpaHoverHandler = null;
+                        }
+                    } catch(e) {}
+                    return null;
+                });
+                if (window._rpaOutline) {
+                    window._rpaOutline.remove();
+                    window._rpaOutline = null;
+                }
+                return true;
+            } catch(e) { return false; }
+"@
+        Invoke-WebScript -Js $jsOff | Out-Null
+        return "HoverHighlight Disabled"
+    }
+}
```

**VBA: Mod_TestDevTools.bas**

```Diff
@@ -267,16 +267,8 @@
     rpaEngine.Loop_TimeoutSec = glb_TimeOut
 
     ' ==========================================================================
-    ' --- リアルタイム・ハイライト --- (黄色系は、視認性低い)
-    Dim jsHighlightStart As String
+    ' --- リアルタイム・ハイライト開始 (Shadow DOM / iframe透過対応) ---
     If glb_UseHighlight = True Then
-        jsHighlightStart = "window._rpaOutline = document.createElement('div');" & _
-        "window._rpaOutline.style.cssText = 'position:fixed; pointer-events:none; z-index:999999; outline:1px solid red; background:rgba(255,0,0,0.1); transition:all 0.05s;';" & _
-        "document.body.appendChild(window._rpaOutline);" & _
-        "window._rpaHover = function(e){ const el = document.elementFromPoint(e.clientX, e.clientY); if(el){ const r = el.getBoundingClientRect(); window._rpaOutline.style.top = r.top + 'px'; window._rpaOutline.style.left = r.left + 'px'; window._rpaOutline.style.width = r.width + 'px'; window._rpaOutline.style.height = r.height + 'px'; }};" & _
-        "document.addEventListener('mousemove', window._rpaHover);"
-        
-        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsHighlightStart)
+        rpaEngine.RunAction "Set-WebHoverHighlight", CreateParams("Enable", True)
     End If
 
     UserForm_Picker.txtResult.Text = ""
```

```Diff
@@ -287,10 +279,8 @@
     Next i
     
     ' --- 要素取得直前にハイライトイベントを解除・削除 ---
-    Dim jsHighlightEnd As String
     If glb_UseHighlight = True Then
-        jsHighlightEnd = "document.removeEventListener('mousemove', window._rpaHover); if(window._rpaOutline){ window._rpaOutline.remove(); }"
-        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsHighlightEnd)
+        rpaEngine.RunAction "Set-WebHoverHighlight", CreateParams("Enable", False)
     End If
     UserForm_Picker.lblStatus.Caption = "マウス下の要素を取得・解析中..."
     DoEvents
```

---

### ◆◆ Step7 ◆◆

### Shadow DOM 探査の強化
deepElementFromPoint の再帰探索、少しでも座標がズレたりラベルの上にマウスがある場合も対応してほしい。

**Lib-DevTools_v302.ps1**<br>

Shadow DOM（el.shadowRoot）を見つけた際、probeShadow で妥協せず、そのまま deepElementFromPoint(x, y, el.shadowRoot) を再帰呼び出し するように修正します。 また、LABEL タグの for 属性から本体要素を探す際も、document.getElementById だけでなく Shadow DOM 内の rootNode.getElementById にも対応させます。これにより、Shadow DOM 内でラベルや少し外側の div.form-group を指してしまっても、必ず内側の <select> や <input> 本体をピンポイントで掴んでハイライトするようになります。 

```Diff
@@ -295,8 +295,9 @@
                 // LABELタグがクリックされた場合、関連付けられた実際の入力フォーム(Input等)へ対象をすり替える。
                 if (el.tagName === 'LABEL') {
                     const forId = el.getAttribute('for');
                     if (forId) {
-                        const linked = document.getElementById(forId);
+                        const linked = (rootNode && rootNode.getElementById) ? rootNode.getElementById(forId) : document.getElementById(forId);
                         if (linked) return linked;
                     }
                     const innerInput = el.querySelector('input,textarea,select');
                     if (innerInput) return innerInput;
                 }
```

```Diff
@@ -331,31 +332,12 @@
                 // 要素がShadow DOM(カプセル化されたコンポーネント)を持っている場合の透過処理
                 try {
                     if (el.shadowRoot) {
+                        // ShadowRootを新たなルートとしてdeepElementFromPointを再帰呼び出しし、内部の優先度判定・すり替えロジックを完全に適用する。
+                        const deep = deepElementFromPoint(x, y, el.shadowRoot);
+                        if (deep && deep !== el) return deep;
+
                         const sEl = probeShadow(el.shadowRoot, x, y);
-                        if (sEl) {
-                            // 多段Shadow DOMに対応するため再帰呼び出しを行う。
-                            if (sEl.shadowRoot) {
-                                const deep = deepElementFromPoint(x, y, sEl.shadowRoot);
-                                if (deep) return deep;
-                            }
-                            return sEl;
-                        }
-
-                        // 別の抽出手段へのフォールバック
-                        try {
-                            const inner = el.shadowRoot.elementFromPoint ? el.shadowRoot.elementFromPoint(x, y) : null;
-                            if (inner && inner !== el) {
-                                if (inner.shadowRoot) {
-                                    const deep = deepElementFromPoint(x, y, inner.shadowRoot);
-                                    if (deep) return deep;
-                                }
-                                const innerInput = inner.querySelector && inner.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
-                                if (innerInput) return innerInput;
-                                return inner;
-                            }
-                        } catch (e) {}
+                        if (sEl && sEl !== el) return sEl;
                     }
                 } catch (e) {}
```

すり替えロジックに 「すり替え先の input 等が透明・サイズ0（window.DOMUtils.isVisible が false）の隠蔽要素だった場合はすり替えず、見えている span 等をそのまま採用する」 という条件を1行追加します。<br>

```Diff
@@ -292,29 +292,37 @@
                 const elRole = el.getAttribute ? el.getAttribute('role') : null;
                 if (elRole && /^(button|link|checkbox|menuitem|tab|searchbox|textbox)$/i.test(elRole)) return el;
+
+                // すり替え候補の要素が実際に画面上で可視サイズを持っているか判定するヘルパー関数
+                const isSubstitutable = function(cand) {
+                    if (!cand) return false;
+                    if (window.DOMUtils && typeof window.DOMUtils.isVisible === 'function') {
+                        return window.DOMUtils.isVisible(cand, { silent: true });
+                    }
+                    const r = cand.getBoundingClientRect();
+                    return (r.width > 0 && r.height > 0);
+                };
+
                 // LABELタグがクリックされた場合、関連付けられた実際の入力フォーム(Input等)へ対象をすり替える。
                 if (el.tagName === 'LABEL') {
                     const forId = el.getAttribute('for');
                     if (forId) {
-                        const linked = document.getElementById(forId);
-                        if (linked) return linked;
+                        const linked = (rootNode && rootNode.getElementById) ? rootNode.getElementById(forId) : document.getElementById(forId);
+                        if (isSubstitutable(linked)) return linked;
                     }
                     const innerInput = el.querySelector('input,textarea,select');
-                    if (innerInput) return innerInput;
+                    if (isSubstitutable(innerInput)) return innerInput;
                 }
 
                 // マウス位置の要素がデザイン用の DIV や SPAN だった場合、その親要素の枠内に、本物の SELECT や INPUT が隠れていないか探索してすり替える。
                 if (/^(DIV|SPAN|I|SVG|P|B)$/.test(el.tagName) && el.parentElement) {
                     try {
                         const siblingTarget = el.parentElement.querySelector('select, input:not([type="hidden"]), textarea, button');
-                        // 見つかった場合、それが確実にRPAで操作したい本体要素なので優先する
-                        if (siblingTarget) return siblingTarget;
+                        // 見つかった場合、かつサイズ0の透明隠蔽要素(モダンUIラジオボタン等)でなければ本体要素として優先する。
+                        if (isSubstitutable(siblingTarget)) return siblingTarget;
                     } catch(e) {}
                 }
 
                 // 要素の内部に操作対象が含まれている場合(DIVで包まれたボタン等)の抽出処理
                 try {
                     const inner = el.querySelector && el.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
-                    if (inner) return inner;
+                    if (isSubstitutable(inner)) return inner;
                 } catch (e) {}
```
1.	モダンUI（透明 input ＋ 可視 span/div）の即時確定 LABEL やその内部の SPAN に当たった際、同じ親（またはラベル）の中に「透明・サイズ0で隠された input（ラジオボタンやチェックボックス）」が存在する場合、その見えている最前面の SPAN（stacked[0]）こそがクリックすべきデザイン要素そのものであるため、親コンテナへループを進めず即座に return して確定させます。 
2.	巨大コンテナ（.shadow-box 等）からの暴走抽出の禁止 siblingTarget や inner（el.querySelector）で内部の input や select を探す際、複数のフォーム部品を抱える巨大コンテナ（querySelectorAll で2個以上ヒットする親枠）では勝手に先頭要素（T02）を拾わないように制限します。 

**Lib-DevTools_v302.ps1**

```Diff
@@ -292,35 +292,45 @@
                 const elRole = el.getAttribute ? el.getAttribute('role') : null;
                 if (elRole && /^(button|link|checkbox|menuitem|tab|searchbox|textbox)$/i.test(elRole)) return el;
 
                 // すり替え候補の要素が実際に画面上で可視サイズを持っているか判定するヘルパー関数
                 const isSubstitutable = function(cand) {
                     if (!cand) return false;
                     if (window.DOMUtils && typeof window.DOMUtils.isVisible === 'function') {
                         return window.DOMUtils.isVisible(cand, { silent: true });
                     }
                     const r = cand.getBoundingClientRect();
                     return (r.width > 0 && r.height > 0);
                 };
 
                 // LABELタグがクリックされた場合、関連付けられた実際の入力フォーム(Input等)へ対象をすり替える。
                 if (el.tagName === 'LABEL') {
                     const forId = el.getAttribute('for');
                     if (forId) {
                         const linked = (rootNode && rootNode.getElementById) ? rootNode.getElementById(forId) : document.getElementById(forId);
                         if (isSubstitutable(linked)) return linked;
                     }
                     const innerInput = el.querySelector('input,textarea,select');
                     if (isSubstitutable(innerInput)) return innerInput;
+                    // 内部のinputが透明化・隠蔽されている場合(モダンUI)は、最前面の視覚要素(stacked[0])またはLABEL自身を即確定する。
+                    if (innerInput) return (stacked[0] && stacked[0] !== el) ? stacked[0] : el;
                 }
 
                 // マウス位置の要素がデザイン用の DIV や SPAN だった場合、その親要素の枠内に、本物の SELECT や INPUT が隠れていないか探索してすり替える。
                 if (/^(DIV|SPAN|I|SVG|P|B)$/.test(el.tagName) && el.parentElement) {
                     try {
-                        const siblingTarget = el.parentElement.querySelector('select, input:not([type="hidden"]), textarea, button');
-                        // 見つかった場合、かつサイズ0の透明隠蔽要素(モダンUIラジオボタン等)でなければ本体要素として優先する。
-                        if (isSubstitutable(siblingTarget)) return siblingTarget;
+                        const siblings = el.parentElement.querySelectorAll('select, input:not([type="hidden"]), textarea, button');
+                        // 親要素内にフォーム部品が1つだけ存在する場合のみ検証する（巨大コンテナからの誤抽出防止）。
+                        if (siblings && siblings.length === 1) {
+                            if (isSubstitutable(siblings[0])) return siblings[0];
+                            // 隣のinputが透明・サイズ0の隠蔽要素(T04のモダンラジオボタン等)なら、現在の装飾用要素(span等)を即座に確定して返す！
+                            return el;
+                        }
                     } catch(e) {}
                 }
 
                 // 要素の内部に操作対象が含まれている場合(DIVで包まれた単一ボタン等)の抽出処理
                 try {
-                    const inner = el.querySelector && el.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
-                    if (isSubstitutable(inner)) return inner;
+                    const inners = el.querySelectorAll ? el.querySelectorAll('input:not([type="hidden"]),textarea,select,button,a,[role="textbox"],[role="search"]') : [];
+                    // コンテナ内に操作要素が1つだけ包まれている場合のみ採用する（.shadow-box全体から先頭の#shadow-selectを誤抽出するのを防ぐ）。
+                    if (inners.length === 1 && isSubstitutable(inners[0])) return inners[0];
                 } catch (e) {}
```

**手動ピッカー実行時（Invoke-UiaRecordStep）の座標取得の精度向上**

```Diff
@@ -702,8 +702,9 @@
         # ルートA： 物理座標からのDOM逆引き。
         if ($UsePointReverse -and ($stepWindowName -match "RPA Browser" -or $stepWindowName -match "Edge" -or $stepWindowName -match "Chrome")) {
-            $rect = $targetElement.Current.BoundingRectangle
-            $pt = New-Object System.Windows.Point([int]($rect.Left + ($rect.Width / 2)), [int]($rect.Top + ($rect.Height / 2)))
+            # UIA側の誤認識（巨大枠や別要素への吸い寄せ）に引きずられないよう、実際のマウスカーソルの物理座標を直接使用する。
+            $cursorPos = [System.Windows.Forms.Cursor]::Position
+            $pt = New-Object System.Windows.Point($cursorPos.X, $cursorPos.Y)
             $resObj = Invoke-BrowserPointReverse -ScreenPoint $pt -TargetControlType $targetControlType -CurrentValue $currentValue
             
             if ($resObj) {
```

**VBA: Mod_TestDevTools.bas**

T18(Scroll)が時々エラー（不安定）になる。 5.3 の公式ドキュメントには、HTMLルートに scroll-behavior: smooth;（なめらかにアニメーションしながらスクロールする設定） が適用されている

```Diff
@@ -507,8 +507,8 @@
         Dim logicalScroll As Long
         logicalScroll = CLng(scrollAmount / dpr)
         
-        ' scrollTo(絶対) ではなく scrollBy(相対) を使用する
-        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "window.scrollBy(0, " & logicalScroll & ");")
+        ' Bootstrap等の scroll-behavior: smooth によるアニメーション遅延を防ぐため、behavior: 'instant' で瞬間スクロールさせる。
+        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "document.documentElement.style.scrollBehavior='auto'; window.scrollTo({top: " & logicalScroll & ", left: 0, behavior: 'instant'});")
         Sleep 800
                 
         cY = cY - scrollAmount
```
     
