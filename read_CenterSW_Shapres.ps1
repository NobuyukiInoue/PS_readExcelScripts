param($Path, $Pattern, $Exclude, $Depth)

##--------------------------------------------------------##
## 引数チェック
##--------------------------------------------------------##

if (-Not $Path) {
    $Path = "../"
}

if (-Not $Pattern) {
    $pattern = "*接続構成図*.xls*"
}

if (-Not $Exclude) {
    $Exclude = "*org*"
}

if (-Not $Depth) {
    $Depth = 2
}


##--------------------------------------------------------##
## ShapeオブジェクトのNameプロパティとTextプロパティを取得
##--------------------------------------------------------##

function readTextRangeText($T_WB, $i_WS)
{
    $resultStr = ""
    for (${i} = 1; ${i} -le ${T_WB}.WorkSheets.item(${i_WS}).Shapes.Count; ${i}++) {
        # $resultStr += "`t" + ${i} + ":" + ${T_WB}.WorkSheets.item(${i_WS}).Shapes(${i}).Name + "`n"

        if (${T_WB}.WorkSheets.item(${i_WS}).Shapes(${i}).TextFrame2 -ne $NULL) {
            if ($T_WB.WorkSheets.item($i_WS).Shapes(${i}).TextFrame2.TextRange.Text -ne $NULL) {
                $ShapeName = ${T_WB}.WorkSheets.item(${i_WS}).Shapes(${i}).Name
                $ShapeText = $T_WB.WorkSheets.item($i_WS).Shapes(${i}).TextFrame2.TextRange.Text
                $ShapeText = $ShapeText.Replace("`n", "`t")
                $resultStr += "`t" + ${ShapeName} + ":`t" + ${ShapeText} + "`n"
            }
        }
    }
    
    Write-Output $resultStr
}


##--------------------------------------------------------##
## キー入力の確認
##--------------------------------------------------------##

function Pause
{
    if ($psISE) {
        $null = Read-Host 'Press Enter Key...'
    }
    else {
        Write-Output "Press Any Key..."
        (Get-Host).UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
    }
}


##--------------------------------------------------------##
## Main関数
##--------------------------------------------------------##
function Main([string]$Path, [string]$Pattern, [string]$Exclude, [int]$Depth) {
    $list = Get-ChildItem -Recurse -Path $Path -Filter $Pattern -Depth $Depth | ForEach {$_.FullName}

    foreach ($fname in $list) {
        Write-Output $fname

        # Excelオブジェクトを取得
        $excel = New-Object -ComObject Excel.Application

        # 表示を有効化
        $excel.Visible = $TRUE

        # 既存のExcelファイルを読み取り専用で開く
        $T_WB = $excel.Workbooks.Open($fname, $FALSE, $TRUE)

        for (${i_WS} = 1; ${i_WS} -le ${T_WB}.WorkSheets.Count; ${i_WS}++) {
            Write-Output "----------------------------------------------"
            Write-Output ${T_WB}.WorkSheets.item(${i_WS}).Name
            Write-Output "----------------------------------------------"
            readTextRangeText ${T_WB} ${i_WS}
        }

        $T_WB.Close($FALSE)
        $excel.Quit()
    }
}


##--------------------------------------------------------##
## Main呼び出し
##--------------------------------------------------------##

Main $Path $Pattern $Exclude $Depth
