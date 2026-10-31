Attribute VB_Name = "Ps_Bridge"
Option Explicit

' --- 特定ウィンドウの最前面移動用API宣言 ---
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Const SW_RESTORE As Long = 9
    
' --- 例外情報格納用構造体 ---
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
    
    ' 引数のペア不一致判定： 奇数個の場合は空の辞書を返却する。
    If (UBound(args) - LBound(args) + 1) Mod 2 <> 0 Then
        Set CreateParams = dict
        Exit Function
    End If
    
    ' キーと値のペア登録： 2つずつループして辞書へ追加する。
    For i = LBound(args) To UBound(args) - 1 Step 2
        dict.Add CStr(args(i)), args(i + 1)
    Next i
    
    Set CreateParams = dict
End Function

' --- PowerShellエラー文字列の構造体パース ---
Public Function ParseRpaError(ByVal errString As String) As RpaExceptionInfo
    Dim result As RpaExceptionInfo
    result.HasError = False
    result.RawText = errString
    
    ' 正規表現オブジェクトの生成： 遅延バインディングでVBScript.RegExpを生成する。
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' パターン定義（[関数名] [種別]: メッセージ (詳細)）： 詳細は省略される場合があるため、末尾の括弧はオプショナルとする。
    regEx.Pattern = "^\[(.*?)\] \[(.*?)\]:\s*(.+?)(?:\s*\(([^\(\)]*)\))?$"
    regEx.IgnoreCase = True
    regEx.Global = False
    
    Dim matches As Object
    Set matches = regEx.Execute(errString)
    
    If matches.Count > 0 Then
        Dim subMatches As Object
        Set subMatches = matches(0).subMatches
        
        result.HasError = True
        result.FunctionName = subMatches(0)
        result.ErrorType = subMatches(1)
        result.Message = subMatches(2)
        
        ' 詳細情報（Details）が存在する場合のみ取得
        If subMatches.Count > 3 Then
            result.Details = subMatches(3)
        End If
    Else
        ' フォーマット外の予期せぬエラーの場合（フォールバック）
        result.HasError = True
        result.FunctionName = "Unknown"
        result.ErrorType = "内部エラー"
        result.Message = errString
    End If
    
    ParseRpaError = result
End Function

' --- Excelを強制的に最前面に持ってくる関数 ---
Public Sub ForceFocusExcel()
    Dim hwnd As LongPtr
    hwnd = Application.hwnd
    
    ' Excelが最小化されている場合は元に戻す
    If IsIconic(hwnd) Then
        ShowWindow hwnd, SW_RESTORE
    End If
    
    ' 最前面に移動
    SetForegroundWindow hwnd
    ' VBAのAppActivateも併用（念のため）
    AppActivate Application.Caption
End Sub

