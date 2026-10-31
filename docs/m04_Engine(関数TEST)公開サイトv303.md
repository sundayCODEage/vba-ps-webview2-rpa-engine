
<h4 align="center">総合テスト・関数一覧表（公開サイト 利用） v303</h4>

<p align="right">
  更新日: 2026年10月29日<br>
  作成者: System beginner & AI Collaborator (Gemini)
</p>

#### 1. 位置づけ
本RPAエンジン（VBA × PowerShell × WebView2 ハイブリッド制御基盤）の機能網羅性を確認し、コードのブラッシュアップ時に既存機能が破損していないかを確認する「リグレッションテスト（回帰テスト）」の一覧です。

&emsp;&emsp;&emsp; ├── Mod_TestRun1_Base一般的な.bas ( 公開サイトでのテスト1 ) /&emsp; **Test_Run1_Base**<br>
&emsp;&emsp;&emsp; ├── Mod_TestRun2_Base少し高度.bas ( 公開サイトでのテスト2 ) /&emsp; **Test_Run2_Base**<br>

「テスト」列に ✓ があるものは、公開サイト利用でのテストコードのPart番号です。

**--- VBA: 公開サイトでのテスト：一般的な ---**

```vba
    prompt = "パートを選んでください：" & vbCrLf & _
            " 0. ** 全パート( 10:UIA除く )一括テスト **" & vbCrLf & _
            " ---" & vbCrLf & _
            " 1. エンジン設定とキャッシュ操作" & vbCrLf & _
            " 2. ナビゲーションと取得系" & vbCrLf & _
            "     ＋ CSSセレクタによる標準DOM操作" & vbCrLf & _
            " 4. XPathによる要素操作" & vbCrLf & _
            " 5. ドロップダウン・チェックボックス" & vbCrLf & _
            " 6. JSネイティブ実行およびCDPネイティブ操作" & vbCrLf & _
            " 7. iframeとタブ(Window)管理" & vbCrLf & _
            " 8. エクスポート・証跡保存・テーブル解析" & vbCrLf & _
            " 9. サイレントダウンロードの正常系テスト" & vbCrLf & _
            " ---" & vbCrLf & _
            "10. UIA(デスクトップ操作) の正常系テスト (メモ帳を使用)"  
    ans = InputBox(prompt, "テストパート選択")
```

**--- VBA: 公開サイトでのテスト：少し高度 ---**

```vba
    prompt = "パートを選んでください：" & vbCrLf & _
            " 0. ** 全パート( 21:PDF除く )一括テスト **" & vbCrLf & _
            " ---" & vbCrLf & _
            " 11. Web上の Shadow DOM ページ (テキスト取得)" & vbCrLf & _
            " 12. Virtual Shadow DOM での操作テスト" & vbCrLf & _
            " 13.（マスク）、スピナー・ローディング待機" & vbCrLf & _
            " 14.（PDF） UIA保存テスト（アプローチ検証）" & vbCrLf & _
            " 15. Fetch API によるバイナリのサイレントDL" & vbCrLf & _
            " 16. Robust DOM(SafeAction) ＆ 物理フォールバック検証" & vbCrLf & _
            " ---" & vbCrLf & _
            " 21. PDFをﾌﾞﾗｳｻﾞ内蔵ビューアで（Adobe Acrobat等）" & vbCrLf & _
            " 22. 仮想タブ（ポップアップ）でのサイレントDL自動継承" & vbCrLf & _
            " 99. 【 v303 】エッジケース検証"         
    ans = InputBox(prompt, "テストパート選択")
```

#### 2. テストシナリオ検証

| 関数名 (Command) | テスト | 実行パート | テスト概要・機能説明 |
| :--- | :---: | :--- | :--- |
| **【1.システム設定・キャッシュ制御】** | | | |
| Set-EngineConfig | ✓ | 1 | エンジンのハイライト設定等の内部フラグ書き換え |
| Write-DebugTextFile | ✓ | 1 | PowerShell側からのテキストファイル直接出力 |
| Clear-WebCache | ✓ | 1 | Cookieやキャッシュの初期化（全クリア） |
| **【2.ナビゲーション・待機制御】** | | | |
| Invoke-WebNavigation | ✓ | 2, 4, 5, 6, 7, 8, 9, 11... | URLへの画面遷移とロード発火 |
| Wait-WebPageLoad | ✓ | 2, 4, 5, 6, 7, 8, 9, 11... | ドキュメントおよび同一オリジンiframeの読み込み完了待機 |
| Wait-WebDocumentReady | ✓ | 2, 13, 16 | DOMツリーの解析完了（軽量待機） |
| Wait-WebUrlContains | ✓ | 2 | 遷移後のURLに特定の文字列が含まれるまでの待機 |
| Wait-WebTitleContains | ✓ | 2 | 画面タイトルが指定文字列に変化するまでの待機 |
| **【3.要素の取得・状態判定】** | | | |
| Get-WebUrl | ✓ | 2 | 現在表示中の絶対URLの取得 |
| Get-WebTitle | ✓ | 2 | 現在のウィンドウタイトルの取得 |
| Get-WebText | ✓ | 3, 7, 11, 12, 13 | CSS指定によるテキスト抽出 |
| Get-WebXPathText | ✓ | 4, 13 | XPath指定によるテキスト抽出 |
| Get-WebAttribute | ✓ | 3, 5, 9, 12, 15 | HTML要素の属性値(href, name, value等)の動的取得 |
| Get-WebCssSelectorHint | ✓ | 16 | XPath指定に合致する要素群から、可視要素のCSSセレクタを逆生成 |
| Get-WebEmbedPdfUrl | ✓ | 21 | 要素の属性を解析し、PDFの絶対URLを動的取得 |
| **【4. DOM操作 (CSS/XPath)】** | | | |
| Wait-WebElement / Wait-WebXPathElement | ✓ | 3, 4, 8, 11, 12, 14, 15 | 対象要素が画面上に出現(可視化)するまでの待機 |
| Test-WebElement | ✓ | 15, 21 | 指定要素の存在確認（例外を投げず True/False の文字列を返す） |
| Wait-WebElementInvisible / Wait-WebXPathElementDisappear| ✓ | 3, 4, 13 | 対象要素が画面から消える（非表示化）までの待機 |
| Wait-WebScreenUnlock | ✓ | 13 | 画面全体を覆うオーバーレイマスク(ロック)解除の自動検知・待機 |
| Set-WebTextInput / Set-WebXPathTextInput | ✓ | 3, 4, 12, 16 | テキストボックスへの文字列入力（失敗時は物理入力へフォールバック） |
| Invoke-WebClick / Invoke-WebXPathClick | ✓ | 3, 4, 7, 9, 12, 13, 99 | ボタンやリンクへのクリック発火（失敗時は物理クリックへフォールバック） |
| Invoke-WebSafeClick | ✓ | 16 | 要素の可視性とサイズを判定し、画面中央へスクロール後にクリック |
| Set-WebCheckbox | ✓ | 5 | チェックボックスのON/OFF状態の指定 |
| Select-WebDropdown | ✓ | 5, 12 | ドロップダウン（select）のvalue値による選択 |
| **【5. iframe・タブ(ウィンドウ)管理】** | | | |
| Wait-WebElementInFrame | ✓ | 7, 99 | iframe内部の要素が出現するまでの透過待機 |
| Invoke-WebClickInFrame | ✓ | 7 | iframe内部の要素への直接クリック操作 |
| List-Tabs | ✓ | 7, 99 | 現在開かれている全タブ（仮想ID含む）のリスト取得 |
| Set-ActiveTab | ✓ | 7, 99 | 指定タブID（P01等）へのアクティブタブ切り替え |
| Switch-Tab / Switch-TabByTitle | ✓ | 7, 99 | タブIDおよびタイトル名によるウィンドウ切り替え（CDP再接続含む） |
| **【6. JavaScript / CDPネイティブ通信】** | | | |
| Invoke-WebView2NativeScript | ✓ | 6 | WebView2標準APIを利用したJS実行と戻り値取得 |
| Invoke-CdpScript | ✓ | 6, 99 | CDP通信を経由したJS実行（Runtime.evaluate） |
| Invoke-WebScript | ✓ | 5, 6, 7, 16, 21, 99 | JS実行の共通窓口（CDP接続時はCDP、未接続・失敗時はNativeへルーティング） |
| Invoke-CdpCommand | ✓ | 6 | CDPの生メソッド（Browser.getVersion等）の直接実行 |
| Set-CdpNativeTextInput / Invoke-CdpNativeClick| ✓ | 6 | DOMイベントを介さないCDPレイヤーからの直接操作 |
| **【7.証跡出力・テーブル解析・ダウンロード】** | | | |
| Export-WebTableToCsv | ✓ | 8 | HTMLテーブルの2次元配列化およびメモリ展開 |
| Export-WebHtml / Export-WebScreenshot | ✓ | 8, 11, 12 | 画面のDOMソースおよび全体キャプチャの保存 |
| Export-WebElementsToCsv / Export-WebFrameTreeToCsv| ✓ | 8, 11 | 操作可能な全要素リスト・iframe階層ツリーの出力 |
| Export-WebDomSnapshot / Export-WebLayoutDump | ✓ | 8, 99 | CDP経由でのDOMツリー・座標スナップショットとレイアウト情報のJSON保存 |
| Enable-SilentDownload | ✓ | 9, 15, 21, 22 | OSの保存ダイアログを抑止したサイレントDL予約 |
| Invoke-WebFetchDownload | ✓ | 15, 21 | Fetch APIを利用した裏側からのバイナリ直接ダウンロード |
| Wait-FileDownload | ✓ | 9, 15, 21, 22 | .crdownload の監視によるファイル保存完了待機 |
| **【8.デスクトップ(UIA)操作・その他】** | | | |
| Switch-AppWindow | ✓ | 10, 14 | 外部アプリ（メモ帳等）の最前面化とフォーカス確保 |
| Invoke-DesktopCenterClick | ✓ | 14 | アクティブウィンドウの中央を物理クリック |
| Invoke-DesktopCoordinateClick | ✓ | 16, 99 | 指定したOS絶対スクリーン座標（物理ピクセル）への物理クリック |
| Invoke-UiaAction | ✓ | 10, 14 | UIAによるUI要素（ボタン・入力欄）の取得と操作 |
| Invoke-DesktopSendKeys | ✓ | 10, 14, 16 | 物理キーボードエミュレーションによるキー・ショートカット送信 |
| Invoke-UiaSafeSaveAs | ✓ | 10, 14 | 「名前を付けて保存」ダイアログに対する安全なファイル保存 |
| Export-WindowScreenshot / Export-WindowHierarchyToCsv | ✓ | 8 | デスクトップウィンドウのキャプチャおよびUI階層出力 |
| Build-RpaStepCode | ✓ | 99 | デスクトップUIA操作用の実行コマンド文字列（Invoke-UiaAction）を生成 |
| **【9.内部関数】** | | | |
| Write-DebugLog | ✓ | | コンソール出力とファイル出力 |
| New-EngineException | ✓ | | 例外情報の標準フォーマット化 |
| Get-ActiveWebView | ✓ |  | 現在アクティブなタブのWebView2インスタンスを取得 |
| Wait-Condition | ✓ | | UIフリーズを防止しつつ、指定条件がTrueになるまで待機 |
| Connect-CdpSession | ✓ | | /json エンドポイントからTargetIdを探査し、WebSocketセッションを確立 |
| Normalize-XPath | ✓ | | XPathの表記揺れ（改行・空白）を自動補正 |
| ConvertTo-JsSafeBase64 | ✓ | | JSインジェクション防止のため、文字列をUTF-8 Base64へエンコード |
| Convert-BoundingBoxToScreenPoint | ✓ | 16, 99 | 論理座標（BoundingBox）を自動スクロール・DPR補正してOSスクリーン座標へ変換 |

---
下記モジュール内の関数は除きます。（画面情報エクスポート＆要素ピッカー:  Mod_TestDevTools でテスト）

##### 🛡️ [SafeAction] 可視性判定・安全クリック
* **Set-WebHoverHighlight:** ピッカー選択待機中、マウスカーソル下の要素（Shadow DOM・同一オリジンiframe透過対応）を赤枠でハイライト表示。

##### 🐛 [Debug] デバッグ・証跡モジュール
* (Export-WebDomSnapshot: 画面上のDOMツリーとスタイルをCDP経由で一括取得し、JSONとして保存。)
* (Export-WebLayoutDump: Page.getLayoutMetrics を用い、DPRやViewportを含むレイアウト情報を保存。)
* **Test-SelectorStability:** 指定したセレクタで要素が継続して取得できるかを検証。
* **Convert-DomSnapshotFlat:** CDPの1次元配列および layout.bounds を、オブジェクトリストに展開。`[内部関数]`
* **Get-DomSnapshotWithBoxModel:** インタラクティブ要素に限定し、DOM構造と座標（layout.bounds / BoxModel）を取得。`[内部関数]`

##### 🔍 [DevTools] 要素ピッカー・コード自動生成モジュール
* **Invoke-UiaRecordStep:** マウスポインタ下の要素を特定し、セレクタとRPAコマンド情報を生成（ピッカー機能のメイン窓口）。
* **Test-WebPointReverse:** 指定したClient座標からの要素取得を検証（自動テスト用）。
* **Test-WebSelectorResolution:** UIA要素からのセレクタ解決（位置照合・Fuzzy検索へのフォールバック挙動）を検証・ログ出力する診断関数。
* **Process-RecordAction:** 取得した要素情報をもとに、DOM逆引き・UIA照合・フォールバック等のルート分岐を制御。`[内部関数]` (他 Get-WebSelectorFromUia, Get-WebSelectorFromUiaFuzzy, Invoke-BrowserPointReverse, Get-CorePointReverseJs, Build-RpaStepCode 等)
