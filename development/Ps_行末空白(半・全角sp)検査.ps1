# =========================================================
# 行末空白（半角・全角スペース）チェック ＆ リストアップスクリプト
# ※ 読み取り専用（ファイルの変更・上書き保存は一切行いません）
# =========================================================

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " 同一階層PSモジュールの、" -ForegroundColor Cyan
Write-Host "         行末（改行直前）の空白（半角・全角）検査スクリプト" -ForegroundColor Cyan
Write-Host " "
Write-Host " ・空白のみの行（コード間のインデント行）は無視します。" -ForegroundColor Gray
Write-Host " ・ファイルの書き換えは一切行いません（リストアップのみ）。" -ForegroundColor Gray
Write-Host " ・表示記号: [␣] = 半角スペース / [□] = 全角スペース" -ForegroundColor Yellow
Write-Host "==========================================================`n" -ForegroundColor Cyan

# スクリプトの保存先フォルダ（同一階層）を取得
$targetPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$selfName   = if ($MyInvocation.MyCommand.Name) { $MyInvocation.MyCommand.Name } else { "" }
Write-Host "検索対象フォルダ（同一階層のみ）: $targetPath`n" -ForegroundColor Cyan

# 同一階層のPS関連ファイルを取得（自身を除く）
$targetFiles = Get-ChildItem -Path $targetPath -File | Where-Object {
    $_.Extension -in '.ps1', '.psm1', '.psd1' -and $_.Name -ne $selfName
}

$detectedFileCount = 0
$totalSpaceLines   = 0

foreach ($file in $targetFiles) {
    $filePath = $file.FullName
    $rawText  = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::UTF8)

    # 空ファイルの場合はスキップ
    if ([string]::IsNullOrEmpty($rawText)) { continue }

    # 改行（CRLF または LF）で行分割
    $lines = $rawText -split "\r?\n"
    $hitList = [System.Collections.Generic.List[PSObject]]::new()

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lineNum  = $i + 1
        $lineText = $lines[$i]

        # 1. 空行、または空白（半角・全角・タブ）のみのインデント行は無視してスキップ
        if ([string]::IsNullOrWhiteSpace($lineText)) { continue }

        # 2. 文字が存在する行で、末尾（改行直前）に半角スペース( )または全角スペース( )があるか正規表現で判定
        if ($lineText -match '^(.*?[\S])([  ]+)$') {
            $bodyPart      = $Matches[1]
            $trailingSpace = $Matches[2]

            # 内訳のカウント
            $halfCount = ([regex]::Matches($trailingSpace, ' ')).Count
            $fullCount = ([regex]::Matches($trailingSpace, ' ')).Count

            # コンソール上で視認できるよう記号に置換
            $visualizedSpace = $trailingSpace.Replace(' ', '[␣]').Replace(' ', '[□]')

            $hitList.Add([PSCustomObject]@{
                LineNum    = $lineNum
                Body       = $bodyPart
                VisualTail = $visualizedSpace
                HalfCount  = $halfCount
                FullCount  = $fullCount
            })
        }
    }

    # 該当行が存在したファイルのみ詳細を出力
    if ($hitList.Count -gt 0) {
        $detectedFileCount++
        $totalSpaceLines += $hitList.Count

        Write-Host "--------------------------------------------------" -ForegroundColor Yellow
        Write-Host "ファイル : $($file.Name) (検出: $($hitList.Count) 行)" -ForegroundColor Green
        Write-Host "パス     : $filePath" -ForegroundColor Gray

        foreach ($item in $hitList) {
            $detail = @()
            if ($item.HalfCount -gt 0) { $detail += "半角:$($item.HalfCount)" }
            if ($item.FullCount -gt 0) { $detail += "全角:$($item.FullCount)" }
            $detailStr = $detail -join ", "

            # 行番号・内訳・本体コードを表示し、末尾の空白部分のみ赤色で強調表示
            Write-Host ("  行 {0,4} ({1,-11}): {2}" -f $item.LineNum, $detailStr, $item.Body) -NoNewline
            Write-Host $item.VisualTail -ForegroundColor Red
        }
        Write-Host ""
    }
}

# 最終サマリー表示
Write-Host "==========================================================" -ForegroundColor Cyan
if ($detectedFileCount -eq 0) {
    Write-Host "行末に余分な空白（半角・全角スペース）を含む行はありませんでした。" -ForegroundColor Green
} else {
    Write-Host "検査完了: 計 $detectedFileCount 件のファイル（合計 $totalSpaceLines 行）で行末の空白を検出しました。" -ForegroundColor Yellow
}
Write-Host "==========================================================" -ForegroundColor Cyan