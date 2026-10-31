
<h3 align="center">汎用RPA操作エンジン 開発・運用仕様書</h3>
<p align="center"><b>【 VBA × PowerShell × WebView2 ハイブリッド制御基盤 】</b></p>

<p align="right">
  版数: 第2.2版( <b>v303対応</b> )<br>
  最終更新日: 2026年10月28日<br>
  作成者: System beginner & AI Collaborator ( <b>Gemini</b> )
</p>

### 第1章 システム概要と設計思想
本システムは、レガシーな業務システムから、Shadow DOMや多段iframe、動的生成UIなどのDOM構造を持つWebアプリケーションまでを操作対象とする汎用RPAエンジンです。<br>
環境構築の手順を簡略化し、操作の再現性を確保するため、以下の設計方針に基づいています。

#### 1.1 基本設計方針
* **インストーラー不要の稼働構成 (Portable Operation)**
外部のWebDriver（Selenium等）や専用のRPAクライアントソフトを使用せず、Windows OSに標準搭載されている **PowerShell 5.1** と **Microsoft Edge WebView2 ランタイム**（および同梱のWebView2 DLL群）で動作します。インストーラーの実行や管理者権限を必要とせず、フォルダの配置とExcelの参照設定のみで稼働します。
* **「基幹」と「拡張」のモジュール分離 (Core & Extension Architecture)**
システム起動と通信を担う「基幹モジュール」と、用途に応じて読み込む「拡張モジュール」に構成を分離しています。Win32 APIやアセンブリのロードは司令塔（Ps_Engine_Core）に一元化し、二重ロードによる競合を防ぐ疎結合な構造としています。
* **複数制御方式の併用 (Native + CDP + UIA)**
通常のDOM API（JavaScript）による操作に加え、CDP（Chrome DevTools Protocol）を用いた通信や、UIAutomation / Win32 API を用いたOSレベルの物理操作（マウスクリック・キーボード入力）へフォールバックする構成を備えています。
* **JSONによる型と状態の維持 (Type Retention)**
言語間（VBA ⇔ PowerShell ⇔ JavaScript）のプロセス通信で生じる「型の消失」を防ぐため、通信プロトコルをJSONペイロードに統一しています。文字列の部分一致判定に依存せず、データ型に基づいたエラーハンドリングを行います。

#### 1.2 アーキテクチャ図解
プロセス間の独立性を保ちながら、標準入出力（パイプ）を介して同期をとる非同期・疎結合なアーキテクチャです。

```text
  +-------------------------------------------------------------+
  |                         VBA (Excel)                         |
  |  - JSONペイロード送信 ／ 応答待ち・タイムアウト監視             |
  +------------------------------+------------------------------+
                                 | (標準入出力パイプ通信)
                                 v
  +-------------------------------------------------------------+
  |               PowerShell 5.1 (Ps_Engine_Core)               |
  |  - Win32 API一括ロード ／ JSONルーティング / 例外パース        |
  +------------------------------+------------------------------+
                   (ドットソース結合モジュール群)
              | (基幹モジュール)                  | (拡張モジュール)
              v                                  v
  +-----------------------+              +----------------------+
  | WebView2_Init/Native  |              | WebAction / WebXPath |
  | WebCDP / WebJS        |              | DesktopUIA / DevTools|
  +-----------------------+              +----------------------+
                 WebView2 フォーム・ブラウザコア
              |                                  |
              | (経路A: Native制御)               | (経路B: CDP / UIA制御)
              | - ExecuteScriptAsync             | - WebSocket通信 / Win32 API
              v                                  v
  +-------------------------------------------------------------+
  |    ターゲットWebページ (DOM / iframe / Shadow DOM) / OS       |
  +-------------------------------------------------------------+
```

---
### 第2章 動作環境とデータ管理
#### 2.1 動作要件・必須ライブラリ
*  Framework: PowerShell 5.1 (Windows OS標準)
*  ランタイム: Microsoft Edge WebView2 ランタイム
*  VBA依存関係: VBA-JSON v2.3.1 (JsonConverter)

> [!NOTE]
> プロセス間通信におけるJSON解析のため、VBE（Visual Basic Editor）の参照設定にて Microsoft Scripting Runtime の有効化が必須です。<br>
> &emsp;&emsp; [ツール] -> `[参照設定]` -> `Microsoft Scripting Runtime` に✅<sub>(チェック)</sub>を入れてください。

#### 2.2 標準ディレクトリ構成とモジュール配置
システム一式を任意のローカルフォルダに配置して使用します。

```text
 [RPA実行ルートフォルダ]
 ├─ Ps_Engine_Core_v303.ps1          (司令塔・ルーター・Win32API/UIA一括ロード)
 │
 ├─✨ [基幹モジュール] (ブラウザ起動と通信に使用する4ファイル)
 │   ├─ Lib-WebJS_v303.ps1           (ブラウザ全画面へのJS自動注入ユーティリティ)
 │   ├─ Lib-WebView2_Init_v303.ps1   (ブラウザ画面生成・仮想タブ管理・クラッシュ保護)
 │   ├─ Lib-WebView2_Native_v303.ps1 (Native API通信・指数バックオフリトライ)
 │   └─ Lib-WebCDP_v303.ps1          (WebSocket・CDP通信)
 │
 ├─✨ [拡張モジュール] (用途に応じた機能群)
 │   ├─ Lib-WebAction_v303.ps1       (標準DOM操作・待機・物理フォールバック・Fetch DL)
 │   ├─ Lib-WebXPath_v303.ps1        (XPath操作・表記揺れ正規化)
 │   ├─ Lib-WebSafeAction_v303.ps1   (可視性確認クリック・透過ホバーハイライト)
 │   ├─ Lib-DesktopUIA_v303.ps1      (OSネイティブ・Win32API物理操作・ダイアログ操作)
 │   ├─ Lib-DevTools_v303.ps1        (GUIピッカー・座標逆引き・セレクタ生成)
 │   └─ Lib-WebDebug_v303.ps1        (DOMスナップショット・結合対応テーブルCSV・証跡保存)
 │
 ├─ 実行用マクロファイル.xlsm        (VBA司令塔: Ps_Engine.cls / Ps_Bridge.bas / UserForm_Picker)
 │
 ├─ [ Libs ]                         (WebView2関連DLL: Core, WinForms, Loader)
 │  ├─ Microsoft.Web.WebView2.Core.dll
 │  ├─ Microsoft.Web.WebView2.WinForms.dll
 │  └─ WebView2Loader.dll
 │
 ├─ [ sandbox ]                      (ローカル検証用テストHTML群: Shadow DOM, マスク待機等)
 ├─ [ Logs ]                         (実行ログ・スクショ等の出力先: 5世代自動ローテーション)
 └─ [ UserData ]                     (WebView2独立ユーザーデータフォルダ)
```

#### 2.3 環境分離と画面サイズの固定 (UDFとMaximized)
* **環境分離 (UDF: \UserData\EBWebView):**
WebView2はセッションデータやキャッシュ、Cookieを管理するために、実行フォルダ直下に独立したユーザーデータフォルダ（UserData/EBWebView）を生成します。OS標準のEdgeブラウザやユーザーの通常セッションとは分離されたブラウザ環境が作成されるため、手作業のブラウザと干渉しません。アプリケーションを終了してもCookieやLocalStorageは保持され、次回起動時に引き継がれます。
【具体例】 業務で通常ログインしているOffice 365や社内ポータルがあっても、RPAエンジンは別の独立したブラウザ環境として起動します。
* **画面の最大化 (Maximized):**
開発時（ピッカー取得時）と本番実行時における「レスポンシブWebデザインによるDOM構造の変化（モバイル表示への切り替わり等）」を防ぐため、RPAブラウザウィンドウは Ps_Engine_Core でOSの有効作業領域を取得し、最大化（Maximized）状態で起動します。これにより、物理クリック時の座標の差異も抑止します。

#### 2.4 キャッシュのクリア
長期間の運用時に古いキャッシュが影響するのを防ぐため、テスト開始時やジョブの区切りで Clear-WebCache を実行し、キャッシュや履歴を削除することができます。

```vba
' モードを指定してキャッシュを削除
' CacheOnly: 一時ファイルや画像キャッシュのみ削除
' All: Cookieやローカルストレージを含めた全履歴をリセット

rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")
```

---
### 第3章 モジュール・コア関数リファレンス
本エンジンは、VBA（呼び出し元）とPowerShell（実行エンジン）の独立したプロセスで構成され、PowerShell側は機能ごとに分割されたモジュール群を動的に結合（ドットソース読み込み）して稼働します。

#### 3.1 VBA側モジュール（操作指示）： 構成例

```text
 ├ 📊xxx_rpa.xlsm       # VBA ( シナリオ )
 ├── Ps_Engine.cls       ( プロセス通信とAPI実行を担うRPAエンジンのコアクラス )
 ├── JsonConverter.bas   ( VBA-JSON: VBAでのJSON解析 )
 ├── Ps_Bridge.bas       ( JSONパラメータ生成・エラー変換などのVBA側ユーティリティ )
 │
 ├── UserForm_Picker.frm ( ピッカー・エクスポート操作用フォーム )
 └── Mod_TestDevTools.bas( ピッカー制御および自動テスト用モジュール )
```

* **Ps_Engine (クラスモジュール):**
PowerShellプロセスの非同期起動、JSONペイロードの構築（BuildPayload）、および標準入出力を介した同期通信（SendAndReceive）を担うクラスです。複数行の戻り値（RESULT）の連結、UI応答性の維持とプロセス終了検知、ミリ秒単位のタイムアウト監視を含みます。
* **Ps_Bridge (標準モジュール):**
プロセス間通信を補助するユーティリティ群です。PowerShellから返された例外文字列を正規表現で解析し、VBAのエラー情報として構造化する ParseRpaError などを提供します。
* **Mod_TestDevTools & UserForm_Picker (フォーム):**
操作対象要素を画面上から取得し、CSSセレクタやUIA識別子を解析してRPAコマンドを生成するためのGUIピッカー機能と制御ロジックです。

#### 3.2 PowerShell コアモジュール群 (Ps_Engine_Core)
* **Ps_Engine_Core (司令塔・ルーティング):** メインループ、JSONルーティング、例外フォーマット生成（New-EngineException）、およびタブ管理を行います。
* **Lib-WebJS (共通ユーティリティ):** ブラウザ初期化時に自動注入されるグローバルJS関数群（多段iframe透過 utilFindInFrames、Shadow DOM貫通 deepQuerySelector、可視性判定 isVisible、セレクタ生成 DOMUtils 等）を定義します。

#### 3.3 ブラウザ制御・通信モジュール
* **Lib-WebView2_Init:** WinFormsの生成、WebView2エンジンの非同期起動、およびJSの自動注入を実施します。ポップアップ（別ウィンドウ要求）をインターセプトし、仮想タブとして保持・破棄する機構を持ちます。
* **Lib-WebView2_Native:** WebView2標準の ExecuteScriptAsync を使用したJS実行基盤です。JSONアンエスケープと指数バックオフによるリトライ機構を含みます。
* **Lib-WebCDP:** WebSocketセッションの確立とJSON-RPCメッセージの送受信を管理します。DOMの評価（Runtime.evaluate）およびOSレベルのマウス・キーボード入力イベント（Input.dispatchMouseEvent 等）の発火を行います。

#### 3.4 DOMアクション・UIA連携モジュール
* **Lib-WebAction:** Invoke-WebClick や Set-WebTextInput など、DOMベースの標準操作を提供します。JS操作失敗時には、座標データを用いた物理操作（フォールバック）へ移行します。Fetch APIを用いたバックグラウンドダウンロード機能も含みます。
* **Lib-WebSafeAction:** 非表示要素への誤操作を防ぐため、可視性の判定とスクロールを伴うクリック（Invoke-WebSafeClick）およびピッカー用のホバーハイライト機能を提供します。
* **Lib-WebXPath:** document.evaluate を用いたXPath探索と、空白・改行の表記揺れに対する正規化（Normalize-XPath）を提供します。
* **Lib-DesktopUIA:** UIAutomationを用いたOSネイティブ操作モジュールです。「名前を付けて保存」ダイアログの操作（クリップボード経由のパス入力）や、Win32 APIを用いた物理クリック・キー送信を担います。
* **Lib-DevTools (ピッカー):** UIA（OS座標）とWebView2（DOM座標）を連携させ、マウスポインタ下の要素からCSSセレクタまたはXPathを逆算します。レーベンシュタイン距離を用いたFuzzy検索によるフォールバック機構を備えています。
* **Lib-WebDebug:** 画面全体のHTML保存、多段iframeのツリー解析、テーブルのCSV・VBA配列出力、CDPによるDOMスナップショットのエクスポート機能等を提供します。

---
### 第4章 プロセス間通信とエラーハンドリング設計
VBA（呼び出し元）とPowerShell（実行エンジン）は独立したプロセスとして稼働するため、標準入出力（パイプ）を介したメッセージプロトコルで同期をとります。

#### 4.1 JSONペイロードによる通信プロトコル
VBAからPowerShellへのコマンド送信は、以下のJSON構造にシリアライズして標準入力（StdIn）へ送信します。

```json
{
  "Command": "Invoke-WebClick",
  "Parameters": {
    "Selector": "button.submit"
  }
}
```

#### 4.2 異常系のパースと伝播
PowerShellからVBAへ状態を返す際は、行頭に以下の識別子を付与し、VBA側はこれを読み取って同期処理を行います。
* [RESULT] ...&nbsp;&nbsp;: 関数の戻り値（データ、取得文字列など。改行を含む複数行データは [SUCCESS] の直前まで連結して取得します）。
* [SUCCESS] ..: コマンドが正常終了したことを示すシグナル。
* [ERROR] ...&nbsp;&nbsp;&nbsp;: PowerShell内で発生した例外メッセージ。

PowerShell側で New-EngineException を用いて生成されたエラー文字列は、VBA側の正規表現パーサー（ParseRpaError）によって RpaExceptionInfo 構造体に変換されます。これにより、「関数名」「エラー種別」「メッセージ」「詳細情報」がVBA側へ伝達されます。

#### 例外文字列の標準フォーマット (New-EngineException)
PowerShell内の各モジュールで例外をスローする際は、New-EngineException 関数を用いて改行を除去した1行の標準フォーマット文字列を生成します。

```text
# 出力フォーマット（Detailsがある場合 / ない場合）
[ERROR][<Func>] [<Type>]: <Message> (<Details>)
[ERROR][<Func>] [<Type>]: <Message>
```

VBA側の正規表現パーサー（Ps_Bridge.bas の ParseRpaError）は、この文字列を解析して RpaExceptionInfo 構造体の各フィールドへ格納します。

| PowerShell側引数<br> (New-EngineException) | VBA側構造体<br> (RpaExceptionInfo) | 内容 |
| :--- | :---: | :---: |
| $Func | FunctionName | 例外が発生したPowerShell関数名（例: Wait-WebElement） |
| $Type | ErrorType | ValidateSet で定義された10種類のエラー種別（下表参照） |
| $Message | Message | エラーの概要メッセージ（改行は半角スペースに正規化） |
| $Details (任意) | Details | 対象のセレクタや下層例外メッセージ等の補足情報（改行は " \| " に正規化） |

定義済みエラー種別 (ValidateSet 全10種)

* **Timeout:** 要素の出現・消滅待機、ページ読み込み、ファイル保存確認などの制限時間超過
* **未発見:** 指定されたセレクタ、XPath、タブ、外部ウィンドウ、テーブル等が存在しない場合
* **JSエラー:** ブラウザ側でのJavaScript実行時例外、または返却されたJSONデータのパース失敗
* **CDPエラー:** CDPターゲット未発見、WebSocket接続・再接続失敗、JSON-RPC通信エラー
* **ネイティブエラー:** WebView2標準API（ExecuteScriptAsync, Navigate, CapturePreviewAsync 等）の実行失敗
* **初期化エラー:** WebView2インスタンスやアクティブタブが未初期化の状態での操作要求
* **ファイルエラー:** HTML、スクリーンショット、CSV、DOMスナップショット等のファイル保存失敗
* **UIAエラー:** UIAutomationまたはWin32 APIによる要素取得・物理クリック・キー送信の失敗
* **引数エラー:** 必須パラメータ（Selector や Attribute 等）の未指定、または未定義のアクション指定
* **内部エラー:** 上記に分類されない処理中の例外、またはフォーマット外のエラー発生時のフォールバック

```vba
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Select Case rpaErr.ErrorType
            Case "Timeout"
                Debug.Print "【待機超過】画面の応答がありません (" & rpaErr.FunctionName & ")"
            Case "未発見"
                Debug.Print "【スキップ】要素が見つかりません: " & rpaErr.Details
                Resume Next ' エラーを無視して次へ進む運用も可能
            Case "JSエラー"
                MsgBox "システムエラー: " & rpaErr.Message, vbCritical
        End Select
    End If
```

#### 4.3 VBA側の通信監視（SendAndReceive）
プロセスの応答停止を防ぐため、VBA側では以下の3つの処理を行っています。
* **プロセスの生存確認:** psProcess.Status を監視し、PowerShellプロセスの終了を検知します。
* **日跨ぎ対応タイムアウト:** Windows API GetTickCount とビットマスク演算を使用し、ミリ秒単位でタイムアウト判定を実施します。
* **UI応答性の維持:** Sleep と DoEvents を組み合わせ、レスポンス待機中のExcel画面のフリーズを防ぎます。

#### 4.4 異言語間連携における型統一ラッパーとコーディング規約
PowerShell（制御側）とJavaScript（ブラウザ側）の通信において発生する「型の消失」を防ぐため、JSONラッパー構造を採用しています。

* **データ型での Return (JS側):** 成功/失敗の判定は return true; または return false; を使用します。
* **型評価 (PS側):** 文字列としての比較（if ($res -eq "true")）は行わず、展開されたBoolean型を直接評価（if ($res) または if (-not $res)）します。
* **末尾のコメント禁止:** JavaScriptをヒアドキュメント（@"..."@）で渡す際、スクリプトの最終行に // による単一行コメントを記述しないでください。ラッパーの閉じカッコ )(); がコメントアウトされ、構文エラーとなります。

---
### 第5章 要素の特定と操作アプローチ
Web自動化における「要素の待機」と「要素の特定・操作」に関する仕様です。

#### 5.1 待機処理 (Wait) の使い分け
Web画面の挙動に合わせて、待機コマンドを選択します。
* **Wait-WebPageLoad (標準待機)**
  * **用途:** 画面遷移時の標準待機。
  * **特徴:** ドキュメントおよび同一オリジンの子iframeの readyState === 'complete' を再帰的に確認します。
* **Wait-WebDocumentReady (軽量待機)**
  * **用途:** 背景で通信が継続する画面や、file:// 環境などでの待機。
  * **特徴:** utilFindInFrames を用いてドキュメントの readyState === 'complete' を確認した時点で復帰します。
* **Wait-WebScreenUnlock / Wait-WebXPathElementDisappear (マスク待機)**
  * **用途:** 「処理中」を示す半透明のオーバーレイ（画面全体を覆うマスク要素）の解除待機。
  * **特徴:** z-index が100以上で画面の80%以上を覆うマスク要素が非表示になるまで待機します。

#### 5.2 CSSセレクタとXPathの指定・正規化
CSSセレクタを用いた操作:（Set-WebTextInput 等）は同一オリジンのiframeおよびShadow DOMを自動探索するため、フレーム階層を個別に指定する必要はありません。
* **ID指定:** &emsp;&emsp;&nbsp; #login-button
* **クラス指定:** &nbsp; .submit-btn
* **属性指定:** &emsp;&nbsp; input\[name='username']

```vba
' テキストボックスへの入力
rpaEngine.RunAction "Set-WebTextInput", CreateParams("Selector", "input[name='username']", "Value", "user123")

' ボタンのクリック
rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", ".submit-btn")
```

**XPathの正規化 (Normalize-XPath):** 特定のテキストを持つ要素を指定する場合はXPathを使用します。<br>
本エンジンでは、//button[text()='送信'] のような完全一致指定が入力された場合、内部で [contains(normalize-space(.), '送信')] に変換（正規化）され、HTMLソース上の改行や空白による不一致を回避します。

```vba
' 「承認」というテキストを持つボタンをXPathで指定してクリック
rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//button[text()='承認']")
```

#### 5.3 非表示要素を除外する操作 (Robust DOM)
Webサイトで用いられる「透明化された入力要素」や「モバイル表示用の非表示メニュー」などを操作対象から除外するための機能です。
* **フェーズ1:** セレクタ逆生成 (Get-WebCssSelectorHint)
XPath指定（例: //a[contains(., 'ニュース')]）に合致する要素群から、可視状態の要素を抽出し、優先度に従ってCSSセレクタを生成します。
* **フェーズ2: 可視性確認付きクリック (Invoke-WebSafeClick)**
  * **可視性の判定 (isVisible):** display: none や visibility: hidden に加え、CSSの透明度（opacity）や要素のサイズ（幅・高さが0より大きいか）を評価し、画面上に表示されているかを判定します。
  * **スクロールとクリック (safeClick):** 対象要素を画面中央へスクロール（scrollIntoView({block: 'center'})）させた後、再度可視性を確認してからクリックを実行します。

#### 5.4 多段iframeとShadow DOMの透過探索（グローバル初期化）
多段iframeとShadow DOMの内部を探索するため、WebView2の初期化フェーズにて CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync() を使用しています。これにより、ページやiframeが生成される際に以下の共通関数群が自動注入されます。
* **deepQuerySelector:** Shadow DOMを再帰的に探索する関数
* **utilFindInFrames:** 同一オリジンの多段iframeを再帰的に探索する関数
* **DOMUtils (isVisible, safeClick, generateCssSelector 等):** 可視性判定およびセレクタ生成を行うユーティリティ

```powershell
# 【実装例 (Set-WebTextInput の場合)】
function Set-WebTextInput {
    param ($Selector, $Value)
    # 事前にBase64でエンコード
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
    $valB64 = ConvertTo-JsSafeBase64 -Text $Value
    
    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const val = decodeURIComponent(escape(atob('$valB64')));
        
        // 注入済みの utilFindInFrames / deepQuerySelector を呼び出す
        const found = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                if (el) {
                    el.focus();
                    el.value = val;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@
    $res = Invoke-WebScript -Js $js
    if (-not $res) { throw (New-EngineException -Func "Set-WebTextInput" -Type "未発見" -Message "入力に失敗しました" -Details $Selector) }
}
```

---
### 第6章 OS・ブラウザ間の座標変換メカニズム
本エンジンでは、ピッカーによる要素取得や、JS操作失敗時の物理操作フォールバックにおいて、OS座標とブラウザ座標を相互に変換する処理を実装しています。

#### 6.1 座標系とDPRの定義
* **Screen座標 (OS絶対座標):** モニター全体の左上を起点とする物理ピクセル座標です。マウスの移動や物理クリック（Win32 API）で使用します。
* **Client座標 (ウィンドウ相対座標):** WebView2コントロールの左上を起点とする物理ピクセル座標です。
* **DPR (Device Pixel Ratio):** Windowsの「ディスプレイの拡大率（125%時は 1.25 など）」です。OSの物理ピクセルとブラウザのCSS論理ピクセルの比率を表します。

#### 6.2 パターンA: 【ピッカー実行時】 OS座標からDOM要素を逆引きする (PointReverse / UIA-DOM)
画面上の要素にマウスを合わせた際、その物理位置からブラウザ内のDOM要素を特定する手順です。
1. **OS座標の取得:** マウスカーソル（またはUIA要素）の絶対物理座標（Screen X/Y）を取得します。
2. **Client座標への変換:** PointToClient メソッドを使用し、ブラウザウィンドウ内での相対物理座標（Client X/Y）に変換します。
3. **ブラウザ論理座標への変換（DPRとスクロール補正）:** 取得した Client X/Y を DPRで除算して論理ピクセルに変換し、必要に応じてページのスクロール量（scrollX / scrollY）を加算します。

$$\text{uiaCenterX} = \left(\frac{\text{clientX}}{\text{dpr}}\right) + \text{scrollX}$$

4. **DOM要素の特定:**
* **PointReverseルート:** elementsFromPoint でShadow DOMや多段iframeを貫通し、要素を特定します（同一オリジンiframe内要素の場合は、親フレームの枠オフセットを再帰加算して絶対論理座標を算出します）。
* **UIA-DOMルート:** 上記計算結果と、DOMSnapshot（layout.bounds）で取得した BoundingBox 座標を比較し、ユークリッド距離で最も近い要素を特定します。

#### 6.3 パターンB: 【操作フォールバック時】 DOM座標からOS物理クリックを実行する
JavaScriptでの操作（.click() 等）が失敗した要素に対し、マウスカーソルを移動させて物理クリックを行う手順（Convert-BoundingBoxToScreenPoint）です。
1. **論理座標の取得:** ピッカー等で記録した対象要素の BoundingBox（ページ絶対論理座標）を受け取ります。
2. **スクロール量の減算:** 絶対論理座標から現在のスクロール量を引き、現在の表示領域内での中心座標（Client論理座標）を算出します。
3. **物理座標への変換（DPRの乗算）:** 算出した論理座標に DPRを乗算し、物理的な Client X/Y に変換します。
4. **Screen座標への変換:** PointToScreen メソッドを使用して、モニター上の絶対物理座標（Screen X/Y）に変換します。
5. **物理クリックの実行:** 算出した Screen X/Y に対して、Win32 API (SetCursorPos と mouse_event) を用いてマウスクリック（または SendKeys によるキー送信）を実行します。

---
### 第7章 応用制御機能 (Advanced Controls)
#### 7.1 Fetch API を用いたファイルのバックグラウンド保存
Edge内蔵のPDFビューア画面などでは、OSの「名前を付けて保存」ダイアログに対するUIA操作がタイムアウトする場合があります。本エンジンでは、現在のセッション（Cookie等のログイン状態）を維持したまま Fetch API でバイナリデータを取得し、保存する手法を用意しています。
1. 対象のURLへバックグラウンドでリクエストを送り、データを取得します。
2. 取得したデータを「Blob（バイナリデータ）」としてメモリに保持します。
3. メモリ上のデータにアクセスするための一時的な内部URL（ObjectURL）を生成します。
4. 非表示のアンカータグ（`<a>`）を生成し、download 属性を付与してクリックイベントを発火させます。

```JavaScript
fetch(target)
    // 取得したデータをバイナリオブジェクト(Blob)へ変換する。
    .then(r => r.blob())
    .then(b => {
        // Blobから一時的なオブジェクトURLを生成する。
        const url = window.URL.createObjectURL(b);
        // アンカータグ(a)を動的に生成し、ダウンロード属性を付与する。
        const a = document.createElement('a');
        a.href = url;
        // 実際のファイル名はエンジン側(Enable-SilentDownload)で上書きする。
        a.download = 'auto_download_temp';
        // DOMへ追加し、クリックイベントを発火させてダウンロードを開始する。
        document.body.appendChild(a);
        a.click();
    });
```
これを Enable-SilentDownload（事前保存先指定）および Wait-FileDownload と組み合わせることで、保存ダイアログを表示せずに指定パスへファイルを保存します。

#### 7.2 仮想タブ管理とプロセスクリーンアップ
業務システムが発行する window.open などの新規ウィンドウ要求を NewWindowRequested イベントで捕捉し、別ウィンドウではなくエンジン内部の「仮想タブ（追加のWebView2コントロール）」として生成・管理します。
親画面との window.opener の参照関係を維持したまま処理を進め、JS側からの window.close 要求、またはプロセス異常終了（ProcessFailed）を検知した際は、該当コントロールを破棄（Dispose）して親画面へ復帰します。

#### 7.3 バッチ監視とポーリング制御
非同期で更新されるバッチ処理状態は、以下の手順で監視します。
1. Wait-WebXPathElementDisappear または Wait-WebScreenUnlock で画面のローディングマスク解除を待機します。
2. Export-WebTableToCsv で FileName に空文字（`""`）を指定し、CSVファイルのディスク保存をスキップしてテーブルの配列文字列データ（セル区切り `<T>`、行区切り `<R>`）のみをメモリ上に取得します。
3. VBAのループ内で目的のステータスに変化するまでポーリングを行います。

```vba
' ファイル出力を行わず、VBAのメモリ上にテーブルデータのみを取得する例
Dim tableData As String
tableData = rpaEngine.RunAction("Export-WebTableToCsv", CreateParams("Selector", "#sampleTable", "FileName", ""))
```

---
### 第8章 開発・拡張時のコーディング規約
本エンジンを保守・拡張する際は、プロセス間通信の互換性を維持するため、以下のコーディング規約を適用します。

#### 8.1 命名規則 (PowerShell)
変数スコープと役割を区別するため、大文字・小文字の使い分けを行います。

|項目|適用ルール|例|備考|
|-|-|-|-|
|**パラメータ・グローバル変数**|PascalCase|$TimeoutSec, $global:Tabs|パラメータとして受け取るもの|
|**ローカル変数**|camelCase|$timeoutSec, $xpathNorm, $xpathB64, $selB64|関数内および初期化作業用変数|
|**時間・間隔変数**|単位を接尾辞化|$WaitTimeMs, $TimeoutSec|Timeout等の曖昧な命名は禁止し、単位（Sec/Ms）を必須とする|
|**VBAからの真偽値引数**|Flg接尾辞|$IsDebugModeFlg, $EnableHighlightFlg|プロセス間通信で文字列("true"/"false")として受ける引数|

#### 8.2 JavaScript インジェクションに関する安全基準
PowerShellからJSのヒアドキュメント（@"..."@）に動的な文字列（CSSセレクタやXPath、入力値など）を展開する際、クォート文字等による構文エラーを防ぐため、以下のエンコード方式を標準とします。
* **PowerShell側でのエンコード:** 変数を埋め込む前に、ConvertTo-JsSafeBase64 を用いてBase64文字列に変換します。
* **JavaScript側でのデコード:** JSコード内で decodeURIComponent(escape(atob('...'))) によりデコードして使用します。
* **変数のキャメルケース化:** パラメータが PascalCase であっても、JSへ埋め込むための作業変数は camelCase（例: $XPath → $xpathB64、$Selector → $selB64）に変換してから使用します。

```powershell
# 【標準実装パターン】
$selB64 = ConvertTo-JsSafeBase64 -Text $Selector
$js = @"
    const selector = decodeURIComponent(escape(atob('$selB64')));
    const el = document.querySelector(selector);
    // ...
"@
```

#### 8.3 JavaScript実装時の制約事項
* **変数宣言:** JSコード内において var による変数宣言は原則使用せず、再代入が不要な変数は const、必要な変数は let を使用します。
* **JS内での変数命名:** PowerShellのPascalCaseをJSコンテキストに持ち込まず、JavaScriptの規約に従い camelCase を使用します。
* **末尾コメントの禁止:** JSヒアドキュメントの最終行（"@ の直前行）に // によるコメントを記述しないでください。ラッパーの閉じカッコ )(); がコメントアウトされ、構文エラーが発生します。

```JavaScript
// 【NGパターンの例】
const el = document.querySelector('.btn');
return el !== null; // ここで要素の有無を返す
// ↑ 最終行に // コメントがあると、ラッパーの閉じカッコ )(); までコメントアウトされます。
```

#### 8.4 型統一ラッパー構造とプロセス間通信の運用ルール
JS側のデータ型（Boolean, Number, Null等）をPowerShell側へ引き継ぐため、下層のJS実行部（Invoke-WebView2NativeScript, Invoke-CdpScript）には以下のラッパーを実装しています。

```JavaScript
// 【JS側ラッパーの構造】
const result = (function() { /* 動的生成されたJSコード */ })();
return JSON.stringify({
    status: "success",
    data: result !== undefined ? result : null
});
```

**スクリプト実行の共通窓口 (Invoke-WebScript):**<br>
DOM操作や値の取得、要素の待機処理において、JavaScriptを実行する際の共通窓口（ルーター）となるのが Invoke-WebScript です。
* **ルーティング:** CDPのWebSocketセッションがOpen状態の場合は Invoke-CdpScript を呼び出し、それ以外の場合は Invoke-WebView2NativeScript を呼び出します。
* **フォールバック:** CDP通信で例外が発生した際は、自動的に Invoke-WebView2NativeScript へ切り替えて実行します。
【PowerShell側の評価ルール】
戻り値は展開された型として評価します。

```PowerShell
$res = Invoke-WebScript -Js $js

# 成功判定 (推奨)
if ($res) { return $true }
# 失敗判定 (推奨)
if (-not $res) { throw "要素が見つかりません" }

# 文字列比較 (非推奨)
if ($res -eq "true") { ... }
```

---
### 第9章 デバッグおよび証跡エクスポート
テストエラー時や動作確認用として、画面の状態をファイルへ出力できます（保存先は実行時セッションごとの Logs フォルダ内で世代管理されます）。
* **Export-WebHtml:** 画面全体（同一オリジンのiframe含む）のHTMLを保存します。
* **Export-WebScreenshot:** 画面のPNGスクリーンショットを保存します。
* **Export-WebElementsToCsv:** 画面内の操作対象要素の属性値と可視状態をCSVに出力します。
* **Export-WebFrameTreeToCsv:** iframe/frameのネスト階層構造をツリー形式でCSVに出力します。
* **Export-WindowHierarchyToCsv:** OS上のウィンドウ階層（プロセス・ハンドル）をCSVに出力します。
* **Export-WebDomSnapshot:** CDPを利用し、画面のDOMツリー構造とスタイル情報・座標をJSON形式で出力します。

---
### 第10章 要素ピッカーの自動生成メカニズムと処理フロー (DevTools)
画面上の要素をマウスポインタで指定し、RPAコマンドを生成する「ピッカー機能」は、VBAからの呼び出し窓口となる関数を起点とし、複数の内部関数が連携して動作します。

#### 10.1 ピッカー処理の全体フロー図
処理は「UIA要素と座標の取得」「ルート別探索（オンデマンドDOMスナップショット）」「コマンド組み立て」「セレクタ安定性テスト」の4フェーズで進行します。

```text
 [VBAからの呼び出し窓口]
 ├─ Invoke-UiaRecordStep     (カーソル位置から取得)
 └─ Test-WebPointReverse     (指定座標から直接取得検証)
      │
      ▽
 [フェーズ1: UIA要素と座標の取得]
 ├─ Get-UiaTargetInfoFromCursor   (最前面要素と親コンテナの探索)
 ├─ Get-UiaElementFromPoint       (物理座標からのUIA要素取得)
 └─ Get-UiaWindowFromElement      (所属するウィンドウ情報の特定)
      │
      ▽
 [フェーズ2: オーケストレーターによる探索 (Process-RecordAction)]
 ├─ ルートA: 座標ベースのDOM逆引き (UsePointReverse有効時)
 │    ├─ Invoke-BrowserPointReverse
 │    └─ Get-CorePointReverseJs (elementsFromPoint / ShadowDOM・iframe貫通 / DOMUtils階層CSS生成・絶対座標算出)
 │
 ├─ ルートB: UIAメタデータによるDOM照合 (代替)
 │    ├─ Get-DomSnapshotWithBoxModel  (CDP layout.bounds による要素座標の一括取得)
 │    ├─ Invoke-UiaDomResolution
 │    ├─ Get-WebSelectorFromUia       (位置・テキストによる照合)
 │    └─ Get-WebSelectorFromUiaFuzzy  (レーベンシュタイン距離を用いた類似検索)
 │
 └─ ルートC: デスクトップ物理操作へのフォールバック
      └─ Build-FallbackStepCode       (OSダイアログ等、ブラウザ外の場合)
      │
      ▽
 [フェーズ3: RPAコマンドの組み立て & フェーズ4: 安定性検証]
 ├─ Build-RpaStepCode            (操作対象のメタデータ構築)
 ├─ ConvertTo-RpaCommandCode     (Invoke-WebClick 等のコマンド情報出力)
 └─ Test-SelectorStability       (deepQuerySelector対応の連続存在確認テスト)
```

#### 10.2 メイン処理フロー
#### 1. エントリーポイント (VBA連携)
* **Invoke-UiaRecordStep:** カーソル下の要素解析を開始します。UsePointReverse（座標逆引き）が有効な場合は事前のスナップショット取得をスキップし、UIA照合（ルートB）が必要な場合のみオンデマンドで Get-DomSnapshotWithBoxModel を呼び出します。
* **Get-DomSnapshotWithBoxModel:** CDPの DOMSnapshot.captureSnapshot（layout.bounds）を使用し、ブラウザ画面内の操作対象要素の「DOMツリー構造」と「BoundingBox（絶対座標）」を取得してメモリ上に保持します。
* **Test-WebPointReverse:** VBAから直接「Client X/Y座標」を渡し、その位置にある要素が取得できるかをテストするための関数です。
#### 2. ターゲット要素の特定 (Get-UiaTargetInfoFromCursor)
現在のマウスカーソルのOS物理座標（X/Y）から、その下にあるUIA要素を取得します。Window や Pane などのコンテナを除外し、親要素をスコアリングして対象要素を特定します。
#### 3. コマンド生成オーケストレーター (Process-RecordAction)
特定されたUIA要素の情報をもとに、以下のルートでDOM要素への逆引きとコマンド生成を行います。
* **ルートA (物理座標ベース逆引き):** マウスカーソルのOS物理座標を Invoke-BrowserPointReverse へ渡し、ブラウザの論理座標へ変換します。Get-CorePointReverseJs によって構築されたJS（elementsFromPoint）を用いてShadow DOMや同一オリジンiframeを再帰的に貫通して要素を特定し、iframe枠オフセットを加味した絶対論理座標（BoundingBox）を算出します。
  * **Shadow DOM内のCSSセレクタ生成:** ブラウザ仕様上、XPath（document.evaluate）はShadow DOMの境界を越えられないため、対象要素がShadowRoot配下（rootNode.host が存在）にある場合は、IDやクラスを持たない要素であってもXPATHへフォールバックせず、注入済みの window.DOMUtils.generateCssSelector(el, root) を呼び出してShadowRoot基準の階層CSSセレクタ（:nth-of-type 等）を生成します。これにより、ピッカー取得後の安定性テスト（Test-SelectorStability）から本番実行（deepQuerySelector）までCSSセレクタで一貫して動作します。
  * **透過・隠蔽UIの判別:** ラジオボタン等で本来の `<input>` が透明・サイズ0で隠蔽されている場合は、可視サイズ判定（isSubstitutable）によってすり替えを抑止し、画面上に表示されている装飾要素（span 等）を捕捉します。
* **ルートB (UIAメタデータベース逆引き):** 座標から特定できなかった場合（または UsePointReverse 無効時）、Invoke-UiaDomResolution へ移行し、UIAの Name や AutomationId などのテキスト情報と DOMSnapshot を照合してDOM要素を特定します。
* **ルートC (デスクトップ操作へのフォールバック):** 対象要素がWebブラウザ外（OSのダイアログなど）であった場合、Build-FallbackStepCode を用いて、UIA操作用のコマンドを生成します。

#### 10.3 セレクタ安定性テスト (Test-SelectorStability)
ピッカーで要素を取得した直後、VBA側から Test-SelectorStability による「セレクタの安定性テスト」が実行されます。
検証の目的:
* **動的IDの検知:** アクセスや操作のたびにIDが変化する要素において、取得直後のセレクタで継続して要素が特定できるかを確認します。
* **仮想DOMの再描画対策:** ReactやVueなどのフレームワークにおいて、取得したセレクタが一定回数（デフォルト3回）連続して要素を捕捉できるかを検証します。

---
### 第11章 モジュール・コア関数リファレンス一覧
本エンジンに実装されている、モジュール別の全関数一覧です。（※ [内部関数] と記載のあるものはモジュール内で呼び出される補助関数です）

#### 11.1 PowerShell コアモジュール群
🛠️ **[Core] 司令塔・ルーティングモジュール (Ps_Engine_Core_v303.ps1)**
* **Write-DebugLog:** コンソール出力とファイル出力（世代管理対応）を行うロギング関数。
* **New-EngineException:** [ERROR] プレフィックスでVBAへ返す例外文字列をフォーマット生成。
* **Get-ActiveWebView:** 現在アクティブなタブのWebView2インスタンスを取得。
* **Set-ActiveTab:** タブIDを直接指定してアクティブタブを切り替え（前面化）。
* **List-Tabs:** 起動中の全タブ情報（ID、URL、タイトル）をJSONで取得。
* **Switch-Tab:** タブIDによる切り替え。CDPの再接続処理を含む。
* **Switch-TabByTitle:** タイトルの部分一致検索によるタブ切り替え。
* **Wait-Condition:** UI応答性を維持しつつ、指定条件がTrueになるまで待機（汎用）。
* **Invoke-WebScript:** JS実行のルーティング。CDPが有効ならCDP、失敗時はNativeへフォールバック。
* **ConvertTo-JsSafeBase64:** JSインジェクション防止のため、文字列をUTF-8 Base64へエンコード。[内部関数]
* **Set-EngineConfig:** 実行時のエンジン設定（要素ハイライトのON/OFF等）を動的に変更。

🚀 **[Init] & [Native] ブラウザ初期化・ネイティブ通信 (Lib-WebView2_Init_v303.ps1 / Lib-WebView2_Native_v303.ps1)**
* **Clear-WebCache:** UDFのキャッシュ、Cookie、LocalStorage等を非同期で削除。
* **Invoke-WebView2NativeScript:** ExecuteScriptAsync を使用したJS実行。JSONアンエスケープとリトライ機構を内包。

⚡ **[CDP] WebSocket通信モジュール (Lib-WebCDP_v303.ps1)**
* **Connect-CdpSession:** json エンドポイントからTargetIdを探査し、WebSocketセッションを確立。[内部関数]
* **Invoke-CdpCommand:** JSON-RPCメッセージの送受信。タイムアウトと自動再接続を管理。
* **Invoke-CdpScript:** CDP経由でのJS評価 (Runtime.evaluate)。戻り値のJSONデコードを含む。
* **Invoke-CdpNativeClick:** CDPを使用し、マウスダウン/アップイベントを座標指定で発火。
* **Set-CdpNativeTextInput:** CDPを使用し、キーボード入力イベントを発火（Input.insertText）。

#### 11.2 DOMアクション・UIA連携・拡張モジュール群
🛡️ **[Action] Web標準操作モジュール (Lib-WebAction_v303.ps1)**
* **Invoke-WebNavigation:** 指定URLへのページ遷移を実行。
* **Wait-WebPageLoad:** DOMの readyState=complete を同一オリジンiframe含めて再帰的に待機。
* **Wait-WebDocumentReady:** ドキュメントの読み込みステータス完了を待機。
* **Wait-WebUrlContains / Wait-WebTitleContains:** URLやタイトルに指定文字列が含まれるまで待機。
* **Wait-WebElement:** 指定要素がDOM上に出現し、かつ画面上に可視化されるまで待機。
* **Wait-WebElementInFrame:** 指定したiframe内の要素が出現・可視化されるまで待機。
* **Invoke-WebClickInFrame:** 指定したiframe内の要素をスクロールしてクリック。
* **Wait-WebElementInvisible:** 指定要素が非表示になる、またはDOMから消滅するまで待機。
* **Wait-WebScreenUnlock:** ローディングマスク（透過レイヤー）の解除を待機。
* **Invoke-WebClick:** 多段iframe・Shadow DOMを透過探索し、対象要素をクリック（失敗時は物理クリックへフォールバック）。
* **Set-WebTextInput:** テキストボックスに値を入力し、input / change イベントを発火（失敗時は物理キー送信へフォールバック）。
* **Select-WebDropdown:** ドロップダウン（select）の指定値を選択し、change イベントを発火。
* **Set-WebCheckbox:** チェックボックスの状態（True/False）を判定し、差異があれば切り替え。
* **Get-WebText:** 要素の innerText または value を取得。
* **Get-WebUrl / Get-WebTitle:** 現在のURL、およびページタイトルを取得。
* **Enable-SilentDownload:** DLダイアログを抑制し、指定フォルダ・ファイル名でのバックグラウンドダウンロードを設定。
* **Wait-FileDownload:** .crdownload の消失および排他ロック解除を確認し、DL完了を待機。
* **Get-WebEmbedPdfUrl:** 埋め込みPDF（embed）の絶対URLを取得。
* **Test-WebElement:** 指定要素の存在確認を行い、例外を投げず "True" / "False" の文字列を返却。
* **Get-WebAttribute:** 指定したWeb要素の属性値（Attribute）を多段iframe・Shadow DOMを透過して取得。
* **Invoke-WebFetchDownload:** Fetch APIを利用したバックグラウンドでのダウンロード発火。
* **Convert-BoundingBoxToScreenPoint:** ブラウザ内の論理座標（BoundingBox）を、DPR・スクロール補正を経てOS絶対スクリーン座標（物理ピクセル）へ変換。[内部関数]

🛡️ **[SafeAction] 可視性判定・安全クリック (Lib-WebSafeAction_v303.ps1)**
* **Get-WebCssSelectorHint:** XPathから可視状態の要素を判定し、CSSセレクタを逆生成。
* **Invoke-WebSafeClick:** スクロールと可視性の確認を行った上で、要素をクリック。
* **Set-WebHoverHighlight:** ピッカー選択待機中、マウスカーソル下の要素（Shadow DOM・同一オリジンiframe透過対応）を赤枠でハイライト表示。

🎯 **[XPath] XPath操作モジュール (Lib-WebXPath_v303.ps1)**
* **Normalize-XPath:** XPathのテキスト指定の表記揺れ（改行・空白）を補正。[内部関数]
* **Wait-WebXPathElement:** XPath指定で要素の可視化を待機。設定有効時は赤枠ハイライトを実行。
* **Wait-WebXPathElementDisappear:** XPath要素の非表示・消滅を待機。
* **Invoke-WebXPathClick:** XPath要素に対し、一連のマウスイベントを発火（失敗時は物理クリックへフォールバック）。
* **Set-WebXPathTextInput:** XPath要素へフォーカスし、テキスト入力とイベント発火を実行。
* **Get-WebXPathText:** XPath要素のタグを判別し、テキスト（value または innerText）を取得。

🖥️ **[UIA] デスクトップ操作モジュール (Lib-DesktopUIA_v303.ps1)**
* **Switch-AppWindow:** Win32 APIを用いて指定した外部ウィンドウを最前面へ移動。
* **Invoke-UiaAction:** UIAutomationを用い、パターン操作または物理キー送信でOS要素を操作。
* **Invoke-UiaSafeSaveAs:** 「名前を付けて保存」ダイアログを捕捉し、クリップボード経由でパスを入力・保存。
* **Invoke-DesktopSendKeys:** 対象のウィンドウへ物理キー（SendKeys）を送信。
* **Invoke-DesktopCenterClick:** アクティブウィンドウの中央を物理クリック。
* **Invoke-DesktopCoordinateClick:** 指定されたX/Y座標（OS絶対スクリーン座標）に対して物理マウスクリックを実行。

🐛 **[Debug] デバッグ・証跡モジュール (Lib-WebDebug_v303.ps1)**
* **Export-WebHtml:** クロスオリジンを考慮し、同一オリジンiframeを含むHTMLスナップショットを保存。
* **Export-WebScreenshot:** CDP、またはネイティブAPIへフォールバックして画面のPNGスクショを保存。
* **Export-WebTableToCsv:** テーブル要素を解析し、VBA取込用の配列文字列データ（`<T>` / `<R>` 区切り）とCSVを生成（FileName に `""` を指定した場合はCSV保存をスキップして配列データのみを取得。ExpandMergedCells 指定により rowspan/colspan の結合セル展開に対応）。
* **Export-WebElementsToCsv:** 画面内の操作可能要素（input, a, button 等）の属性を抽出してCSV化。
* **Export-WebFrameTreeToCsv:** 多段iframeのネスト構造をツリー形式で解析しCSV化。
* **Export-WindowScreenshot:** Win32 API等を使用し、ブラウザの枠を含むウィンドウ全体のスクショを保存。
* **Export-WindowHierarchyToCsv:** OS上で起動している全プロセスのハンドルとタイトル一覧をCSV出力。
* **Write-DebugTextFile:** 任意の文字列をデバッグ用テキストファイルへ追記保存。
* **Export-WebDomSnapshot:** 画面上のDOMツリーとスタイルをCDP経由で取得し、JSONとして保存。
* **Export-WebLayoutDump:** Page.getLayoutMetrics を用い、DPRやViewportを含むレイアウト情報を保存。
* **Test-SelectorStability:** 指定したセレクタで要素が継続して取得できるかを検証（Shadow DOM・iframe透過対応）。
* **Convert-DomSnapshotFlat:** CDPの1次元配列および layout.bounds を、オブジェクトリストに展開。[内部関数]
* **Get-DomSnapshotWithBoxModel:** インタラクティブ要素に限定し、DOM構造と座標（layout.bounds / BoxModel）を取得。[内部関数]

🔍 **[DevTools] 要素ピッカー・コード自動生成モジュール (Lib-DevTools_v303.ps1)**
* **Invoke-UiaRecordStep:** マウスポインタ下の要素を特定し、セレクタとRPAコマンド情報を生成（ピッカー機能のメイン窓口）。
* **Test-WebPointReverse:** 指定したClient座標からの要素取得を検証（自動テスト用）。
* **Test-WebSelectorResolution:** UIA要素からのセレクタ解決（位置照合・Fuzzy検索）を検証・ログ出力する診断関数。
* **Process-RecordAction:** 取得した要素情報をもとに、DOM逆引き・UIA照合・フォールバック等のルート分岐を制御。[内部関数] (他 Get-WebSelectorFromUia, Get-WebSelectorFromUiaFuzzy, Invoke-BrowserPointReverse, Get-CorePointReverseJs 等)
