# ------------------------------------------------------------------------------#●#✘
<#
  [Robust DOM Utilities]： Lib-WebJSでブラウザに注入された DOMUtils を活用し、高レベルのフェイルセーフ実行機能を提供する。
#>
# ------------------------------------------------------------------------------

# --- 堅牢なCSSセレクタの動的生成 ---
# 曖昧なXPathから、可視状態の要素を厳密に判定し、一意のCSSセレクタを逆算生成する。
function Get-WebCssSelectorHint {
    param ([Parameter(Mandatory=$true)][string]$XPath)
    
    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $xpathB64 = ConvertTo-JsSafeBase64 -Text $XPath

    # 検索ワードの抽出処理： ログ出力時の視認性を高めるため、XPathからテキスト検索部分を抽出する。
    $searchWord = $XPath
    if ($XPath -match "contains\(\.,\s*['`"]([^'`"]+)['`"]\)") {
        $searchWord = $Matches[1]
    }

    $js = @"
        try {
            // Base64化されたXPathをデコードして復元する。
            const xpath = decodeURIComponent(escape(atob('$xpathB64')));
            // 共通ユーティリティを利用して、XPathに合致するすべての要素を取得する。
            const els = window.DOMUtils.getElementsByXPath(xpath, document);
            // 要素が存在しない場合はNOT_FOUNDステータスを返却する。
            if (!els || els.length === 0) return JSON.stringify({ status: 'NOT_FOUND' });

            const results = [];
            let bestSelector = null;

            // 取得した要素群をループ処理し、それぞれのCSSセレクタを生成する。
            for (let i = 0; i < els.length; i++) {
                let sel = '(生成不可)';
                try {
                    // 要素から一意なCSSセレクタを逆算する。
                    sel = window.DOMUtils.generateCssSelector(els[i], document);
                } catch(e){}

                // 判定のみ行うため、探索時のisVisibleはログを出さない(silent: true)設定とする。
                const isVis = window.DOMUtils.isVisible(els[i], { silent: true });

                // 結果配列に候補情報を追加する。
                results.push({ index: i + 1, selector: sel, visible: isVis });

                // 最初に発見した「可視状態」の要素のセレクタを最適セレクタとして保持する。
                if (!bestSelector && isVis) {
                    bestSelector = sel;
                }
            }
            
            // 可視の要素が一つもない場合は、先頭候補を最適セレクタとする。
            if (!bestSelector) bestSelector = results[0].selector;

            // 解析結果をJSON文字列化して返却する。
            return JSON.stringify({ status: 'SUCCESS', count: els.length, data: results, bestSelector: bestSelector });
        } catch(e) {
            return JSON.stringify({ status: 'ERROR', message: e && e.message ? e.message : String(e) });
        }
"@

    # 構築したJSコードをブラウザ側へ送信し、実行結果を受け取る。
    $resJson = Invoke-WebScript -Js $js
    # 戻り値のJSON文字列をPowerShellのオブジェクトに変換する。
    try {
        $resObj = $resJson | ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "セレクタ解析結果(JSON)のデコードに失敗しました" -Details $_.Exception.Message)
    }
    
    # エラーハンドリング： JSONパース後のステータスに基づき例外をスローする。
    if ($resObj.status -eq 'NOT_FOUND') { throw (New-EngineException -Func $func -Type "未発見" -Message "XPathに該当する要素が見つかりません" -Details $XPath) }
    if ($resObj.status -eq 'ERROR') { throw (New-EngineException -Func $func -Type "JSエラー" -Message "セレクタの生成に失敗しました" -Details $resObj.message) }
    
    Write-DebugLog -Message "[Robust DOM] XPath検索: [$searchWord] に合致する要素を $($resObj.count) 件発見しました。" -Level Info
    foreach ($item in $resObj.data) {
        $visMark = if ($item.visible) { "[可視〇]" } else { "[非表示]" }
        Write-DebugLog -Message "  -> 候補 $($item.index):$visMark $($item.selector)" -Level Info
    }
    Write-DebugLog -Message "[Robust DOM] 最適セレクタ: $($resObj.bestSelector)" -Level Info
    # 最も条件の良いセレクタを返却する。
    return $resObj.bestSelector
}

# --- 対象要素の可視化待機と安全なクリック実行 ---
# スクロールと可視性の最終確認を行った上で、隠し要素の誤爆を防ぎ安全にクリックを実行する。
function Invoke-WebSafeClick {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )
    
    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selectorB64 = ConvertTo-JsSafeBase64 -Text $Selector

    # ログ状態の重複出力を防ぐためのトラッキング変数を初期化する。
    $state = @{ LastLogState = "" }
    # Wait-Conditionに渡すための条件スクリプトブロックを定義する。
    $condition = {
        $js = @"
            try {
                // Base64化されたセレクタをデコードして復元する。
                const selector = decodeURIComponent(escape(atob('$selectorB64')));
                // 以前のトラッキングログをクリアする。
                window.DOMUtils.tracker.clear();

                // セレクタに合致するすべての要素を取得する（通常DOMで未発見時はShadow DOMも探索する）。
                // ※ <querySelector> トップ画面およびShadow DOM内の安全クリックを主目的としており、同一オリジンiframe内の要素を対象とする場合は Lib-WebAction の Invoke-WebClickInFrame を使用する。
                let els = Array.from(document.querySelectorAll(selector));
                if (els.length === 0 && typeof deepQuerySelector === 'function') {
                    const shadowEl = deepQuerySelector(selector, document);
                    if (shadowEl) els = [shadowEl];
                }
                window.DOMUtils.tracker.log('探索開始: セレクタに合致する要素 ' + els.length + ' 件');
                
                if (els.length === 0) {
                    window.DOMUtils.tracker.log('DOM内に該当要素が存在しません。');
                } else {
                    // 取得した要素を順に検証する。
                    for (let i = 0; i < els.length; i++) {
                        window.DOMUtils.tracker.log('--- [候補 ' + (i+1) + '/' + els.length + '] の検証を開始 ---');

                        // ここは検証スタート地点なので、isVisibleの成功ログ(幅・高さ)を1回だけ出力させる。
                        if (window.DOMUtils.isVisible(els[i])) {
                            // 可視であることが確認できたら、安全なクリックを実行する。
                            window.DOMUtils.safeClick(els[i]);
                            // クリック成功時、ログを含めてSUCCESSを返却する。
                            return JSON.stringify({ status: 'SUCCESS', logs: window.DOMUtils.tracker.logs });
                        } else {
                            window.DOMUtils.tracker.log('[候補 ' + (i+1) + '] は条件を満たさないためスキップします。');
                        }
                    }
                }
                // すべての候補が条件を満たさない（または未出現）の場合はWAITステータスを返却する。
                return JSON.stringify({ status: 'WAIT', logs: window.DOMUtils.tracker.logs });
            } catch(e) {
                return JSON.stringify({ status: 'ERROR', message: e && e.message ? e.message : String(e), logs: window.DOMUtils.tracker.logs });
            }
"@

        # 構築したJSコードをブラウザ側へ送信し、実行結果を受け取る。
        $resJson = Invoke-WebScript -Js $js
        # 戻り値のJSON文字列をPowerShellのオブジェクトに変換する。
        $resObj = $null
        try {
            $resObj = $resJson | ConvertFrom-Json -ErrorAction Stop
        } catch {
            # JSONパース失敗時は待機ループを継続させるため $false を返す。
            return $false
        }

        if ($null -ne $resObj) {
            # デバッグモード有効時のログ処理： ログが存在する場合に処理を実行する。
            if ($global:IsDebugMode -and $null -ne $resObj.logs) {
                # 現在のログ配列をパイプ(|)で結合し、状態文字列を作成する。
                $currentLogState = $resObj.logs -join "|"
                # 前回の状態から変化があった場合のみ、コンソール/ファイルへ出力する。
                if ($currentLogState -ne $state.LastLogState) {
                    foreach ($logStr in $resObj.logs) {
                        Write-DebugLog -Message "[Robust DOM] $logStr" -Level Info
                    }
                    # トラッキング変数を最新の状態に更新する。
                    $state.LastLogState = $currentLogState
                }
            }
            
            if ($resObj.status -eq 'SUCCESS') { return $true }
            if ($resObj.status -eq 'ERROR') {
                throw (New-EngineException -Func $func -Type "JSエラー" -Message "安全なクリック処理に失敗しました" -Details $resObj.message)
            }
        }
        # 成功しなかった場合は$falseを返し、Wait-Conditionのループを継続させる。
        return $false
    }
    $errMsg = "[$func] タイムアウト: 要素の出現・可視化、またはクリックに失敗しました ($Selector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null
    
    return "SafeClick Completed: $Selector"
}

# --- ピッカー用リアルタイム・ホバーハイライトのON/OFF制御 ---
# マウスカーソル下の要素（Shadow DOM・同一オリジンiframe透過対応）をリアルタイムで赤枠ハイライトする。
function Set-WebHoverHighlight {
    param (
        [Parameter(Mandatory=$true)]
        [Alias("Enable")]
        [string]$EnableFlg
    )

    $func = $MyInvocation.MyCommand.Name
    $enable = $false
    try { $enable = [System.Convert]::ToBoolean($EnableFlg) } catch {}

    if ($enable) {
        $js = @"
            try {
                // 既存のハイライト枠が残っていれば一旦削除する。
                if (window._rpaOutline) { window._rpaOutline.remove(); }
                
                const outline = document.createElement('div');
                outline.style.cssText = 'position:fixed; pointer-events:none; z-index:2147483647; outline:2px solid red; background:rgba(255,0,0,0.1); transition:all 0.04s;';
                document.body.appendChild(outline);
                window._rpaOutline = outline;

                // Shadow DOMおよびiframeを貫通してマウス下の要素とトップ画面基準の矩形を算出する。
                window._rpaUpdateHover = function(clientX, clientY, baseWin) {
                    let curDoc = baseWin.document;
                    let x = clientX;
                    let y = clientY;
                    let offsetX = 0;
                    let offsetY = 0;

                    // baseWinが子iframeの場合、親フレーム群のオフセットを加算する。
                    let w = baseWin;
                    while (w && w !== window.top) {
                        try {
                            const fe = w.frameElement;
                            if (!fe) break;
                            const fr = fe.getBoundingClientRect();
                            offsetX += fr.left + (fe.clientLeft || 0);
                            offsetY += fr.top  + (fe.clientTop  || 0);
                            w = w.parent;
                        } catch(e) { break; }
                    }

                    let el = curDoc.elementFromPoint(x, y);
                    // トップ画面上でiframe自体に当たった場合は内部へ透過する。
                    while (el && (el.tagName === 'IFRAME' || el.tagName === 'FRAME')) {
                        try {
                            const fr = el.getBoundingClientRect();
                            offsetX += fr.left + (el.clientLeft || 0);
                            offsetY += fr.top  + (el.clientTop  || 0);
                            x -= (fr.left + (el.clientLeft || 0));
                            y -= (fr.top  + (el.clientTop  || 0));
                            if (!el.contentDocument) break;
                            curDoc = el.contentDocument;
                            el = curDoc.elementFromPoint(x, y);
                        } catch(e) { break; }
                    }

                    // Shadow DOM内部へ再帰的に透過する。
                    while (el && el.shadowRoot && el.shadowRoot.elementFromPoint) {
                        const inner = el.shadowRoot.elementFromPoint(x, y);
                        if (!inner || inner === el) break;
                        el = inner;
                    }

                    if (el && window._rpaOutline) {
                        const r = el.getBoundingClientRect();
                        window._rpaOutline.style.top    = (r.top + offsetY) + 'px';
                        window._rpaOutline.style.left   = (r.left + offsetX) + 'px';
                        window._rpaOutline.style.width  = r.width + 'px';
                        window._rpaOutline.style.height = r.height + 'px';
                    }
                };
                // トップウィンドウおよび同一オリジンiframeへmousemoveイベントを登録する。
                utilFindInFrames(window, function(win) {
                    try {
                        win._rpaHoverHandler = function(e) { window._rpaUpdateHover(e.clientX, e.clientY, win); };
                        win.document.addEventListener('mousemove', win._rpaHoverHandler);
                    } catch(e) {}
                    return null;
                });
                return true;
            } catch(e) { return false; }
"@
        Invoke-WebScript -Js $js | Out-Null
        return "HoverHighlight Enabled"
    } else {
        $jsOff = @"
            try {
                utilFindInFrames(window, function(win) {
                   try {
                        if (win._rpaHoverHandler) {
                            win.document.removeEventListener('mousemove', win._rpaHoverHandler);
                            win._rpaHoverHandler = null;
                        }
                    } catch(e) {}
                    return null;
                });
                if (window._rpaOutline) {
                    window._rpaOutline.remove();
                    window._rpaOutline = null;
                }
                return true;
            } catch(e) { return false; }
"@
        Invoke-WebScript -Js $jsOff | Out-Null
        return "HoverHighlight Disabled"
    }
}
