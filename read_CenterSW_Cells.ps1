param($Path, $Pattern, $Exclude, $Depth)

##--------------------------------------------------------##
## 引数チェック
##--------------------------------------------------------##

if (-Not $Path) {
    $Path = "../"
}

if (-Not $Pattern) {
    $pattern = "*実装表*.xls*"
}

if (-Not $Exclude) {
    $Exclude = "*org*"
}

if (-Not $Depth) {
    $Depth = 2
}


##--------------------------------------------------------##
## センターSWの設定ポートの行を出力
##--------------------------------------------------------##
function printCenterSW($T_WB, [int]$SheetNum, [array]$hostnameInStr, [int]$LineCountMAX)
{
    $isFound = $FALSE

    for ($i = 1; $i -le $LineCountMAX; $i++) {

        $portnum = $T_WB.worksheets.item($SheetNum).Range("E18").Cells($i, 1).text
        $hostname = $T_WB.worksheets.item($SheetNum).Range("E18").Cells($i, 4).text
        $vrf = $T_WB.worksheets.item($SheetNum).Range("E18").Cells($i, 7).text

        if (($portnum.Length -eq 0) -And $isFound) {
            readCells $T_WB $SheetNum "F18" $i 12
            continue
        }

        $isFound = $FALSE

        if (-Not $vrf.Equals("10")) {
            # VRFが"10"でない場合は次の行へ
            continue
        }

        foreach ($targetStr in $hostnameInStr) {

            $pos = $hostname.IndexOf($targetStr)

            if ($pos -ge 0) {
                # ホスト名に "-c" または "-PN-c" が含まれている場合
                $isFound = $TRUE
                readCells $T_WB $SheetNum "F18" $i 12
                break
            }
        }
    }
}

##--------------------------------------------------------##
## 指定セルの内容を表示
##--------------------------------------------------------##
function readCells($T_WB, [int]$SheetNum, [string]$T_Address, [int]$row, [int]$count)
{
    $resultStr = $T_WB.worksheets.item($SheetNum).Range($T_Address).Cells($row, $j).address() + ":"
    for ($j = 0; $j -lt $count; $j++) {
        $val = $T_WB.worksheets.item($SheetNum).Range($T_Address).Cells($row, $j).text
        $val = $val.Replace("`n", "")
        $resultStr += [String]::Format("{0,-15}", $val)
    }
    
    Write-Output $resultStr
}


##--------------------------------------------------------##
## キー入力を確認
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
    $hostnameInStr = @("-c", "-PN-c")

    # Excelオブジェクトを取得
    $excel = New-Object -ComObject Excel.Application

    # 表示を有効化
    $excel.Visible = $TRUE

    foreach ($fname in $list) {

        # 既存のExcelファイルを読み取り専用で開く
        $T_WB = $excel.Workbooks.Open($fname, $FALSE, $TRUE)

        if ($T_WB.WorkSheets.Count -ge 2) {
            #--------------------------------------------------#
            # 統合ルータA
            #--------------------------------------------------#
            Write-Output "---------------------------------------"
            Write-Output $fname
            Write-Output $T_WB.WorkSheets.item(2).Name
            Write-Output "---------------------------------------"
        
            #-----------------------------------------------------#
            # WorkSheets.item(2) 新旧センターSWへの接続ポート情報
            #-----------------------------------------------------#
            printCenterSW $T_WB 2 $hostnameInStr 40 
        }

        if ($T_WB.WorkSheets.Count -ge 3) {
            #--------------------------------------------------#
            # 統合ルータB
            #--------------------------------------------------#
            Write-Output "---------------------------------------"
            Write-Output $fname
            Write-Output $T_WB.WorkSheets.item(3).Name
            Write-Output "---------------------------------------"

            #-----------------------------------------------------#
            # WorkSheets.item(3) 新旧センターSWへの接続ポート情報
            #-----------------------------------------------------#
            printCenterSW $T_WB 3 $hostnameInStr 40 
        }

        $T_WB.Close()
    }
    $excel.Quit()
}


##--------------------------------------------------------##
## Main呼び出し
##--------------------------------------------------------##

Main $Path $Pattern $Exclude $Depth
