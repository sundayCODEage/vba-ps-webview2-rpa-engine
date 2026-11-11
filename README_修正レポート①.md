
#### 🔧 【v303 修正レポート ①】待機ループ内の `DoEvents()` 漏れによるUIスレッド競合について

<p align="right">
  作成日: 2026年11月11日 / 作成者: System beginner & AI Collaborator (Gemini)
</p>

業務システムでの稼働テストにおいて、ポップアップ（仮想タブ）のクローズ直後や高負荷時のマスク解除待機で処理がタイムアウトするケースが確認されました。一部の待機ループにおける [System.Windows.Forms.Application]::DoEvents() の記述漏れが原因であることが判明したため、該当箇所の修正を行いました。

#### 1. 技術的背景（STAモードとメッセージポンプ）
本エンジンは PowerShell を `-STA`（Single-Threaded Apartment）モードで起動し、単一のUIスレッド上で以下を同時稼働させています。
* WinForms ウィンドウ（ $global:BrowserForm ）の描画とOSメッセージ処理
* WebView2 の非同期イベント処理（ NewWindowRequested による仮想タブ捕獲、WindowCloseRequested によるタブ破棄、DownloadStarting によるサイレント保存）
* VBAからのJSONコマンド実行とポーリング待機ループ

PowerShell側で while ループによる待機を行う際、ループ内で [System.Windows.Forms.Application]::DoEvents() を呼び出さないと、**Start-Sleep でスレッドが休止している間、WinFormsおよびWebView2のイベントキュー（メッセージポンプ）が完全に停止**します。

#### 2. なぜローカルsandboxや公開サイトのテストでは見過ごされたのか？
エンジン内のほとんどの待機関数は、Coreモジュールの共通関数 Wait-Condition（内部で DoEvents() を実行）を経由しています。
しかし、Lib-WebXPath_v303.ps1 の Wait-WebXPathElementDisappear のみ、Wait-Condition を使わず独自の while ループ（500ms間隔）を実装しており、ループ末尾の DoEvents() が漏れていました。

* **単体テストで動いていた理由：**
  ループ内部で呼び出している Invoke-WebView2NativeScript の中（ ExecuteScriptAsync の数ミリ秒の応答待ち）に短い DoEvents() が有るため、同一画面内で単純にローディングマスクが消えるだけのテストでは、500msに1回の一瞬のメッセージ処理だけで通過できていました。
* **業務システム実機で顕在化した理由：**
  実機では「ダウンロード子画面（仮想タブ）で『戻る』を押して window.close() を発火させ、直後に親画面へ戻って Wait-WebXPathElementDisappear（マスク消滅待機）に入る」という複合動作が発生します。この時、500msの Start-Sleep がUIスレッドをブロックし、子画面の破棄イベントと親画面のマスク解除処理が競合していました。

#### 3. 放置した場合に発生する不具合（リスク）

1. **仮想タブ（ポップアップ）破棄イベントの滞留とタイムアウト ( Lib-WebXPath_v303.ps1 )**
   * 子画面が window.close() を要求した際、Lib-WebView2_Init_v303.ps1 は $global:BrowserForm.BeginInvoke() を用いてUIスレッドへタブ破棄（Dispose）を依頼します。
   * 直後に Wait-WebXPathElementDisappear に入ると、BeginInvoke の処理がUIスレッドの休止によって待たされ、親画面のロック解除が進行せずタイムアウト（30秒）に至るケースがあります。
2. **長時間待機時のウィンドウ「応答なし」化とフォーカスイベントの不発**
   * サーバー高負荷時など、要素消滅やUIA要素出現に数秒〜十数秒かかる場面でOSの再描画やフォーカス通知が滞り、ブラウザが一時的に「応答なし」状態になったり、onblur` / `onfocus イベントが欠落するリスクがあります。
3. **OSダイアログ操作後のダウンロード完了通知の遅延 ( Lib-DesktopUIA_v303.ps1 )**
   * Invoke-UiaSafeSaveAs のファイル保存完了待機ループ等でUIスレッドが止まると、WebView2本体側のファイルストリーム書き込み完了イベント自体がブロックされ、保存確認がタイムアウトする可能性があります。

#### 4. 修正対象ファイルと差分コード
以下の2ファイル（計4箇所）の while ループ内に [System.Windows.Forms.Application]::DoEvents() を追加し、全モジュールでUIスレッドの応答性を統一しました。

#### ① `Lib-WebXPath_v303.ps1`（ Wait-WebXPathElementDisappear / 150行目付近）
```diff
     while ($sw.Elapsed.TotalSeconds -lt$TimeoutSec) {
         # DOM遷移中等の通信切断を避けるため、CDPではなくネイティブAPIを用いて判定を繰り返す。
         $exists = Invoke-WebView2NativeScript -Js$js -Retries 1
         # 戻り値が false (非表示/消滅) となった場合はループを抜ける。
         if (-not $exists) {
             $isGone =$true
             break
         }
+        [System.Windows.Forms.Application]::DoEvents()
         Start-Sleep -Milliseconds 500
     }
```

#### ② `Lib-DesktopUIA_v303.ps1`（ Invoke-UiaAction 1箇所 / Invoke-UiaSafeSaveAs 2箇所）
```diff
             # 1. Invoke-UiaAction の要素探索ループ内 (170行目付近)
+            [System.Windows.Forms.Application]::DoEvents()
             Start-Sleep -Milliseconds $PollIntervalMs
         }
```
```diff
         # 2. Invoke-UiaSafeSaveAs のダイアログ探索ループ内 (225行目付近)
         if ($hwnd -ne [IntPtr]::Zero -and $null -ne$hwnd) { break }
+        [System.Windows.Forms.Application]::DoEvents()
         Start-Sleep -Milliseconds 200
     }
```
```diff
         # 3. Invoke-UiaSafeSaveAs のファイル保存完了待機ループ内 (305行目付近)
+        [System.Windows.Forms.Application]::DoEvents()
         Start-Sleep -Milliseconds 300
     }
```
