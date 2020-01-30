param($Path, $Pattern, $Exclude, $Depth)

##--------------------------------------------------------##
## �����`�F�b�N
##--------------------------------------------------------##

if (-Not $Path) {
    $Path = "../"
}

if (-Not $Pattern) {
    $pattern = "*�����\*.xls*"
}

if (-Not $Exclude) {
    $Exclude = "*org*"
}

if (-Not $Depth) {
    $Depth = 2
}


##--------------------------------------------------------##
## �Z���^�[SW�̐ݒ�|�[�g�̍s���o��
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
            # VRF��"10"�łȂ��ꍇ�͎��̍s��
            continue
        }

        foreach ($targetStr in $hostnameInStr) {

            $pos = $hostname.IndexOf($targetStr)

            if ($pos -ge 0) {
                # �z�X�g���� "-c" �܂��� "-PN-c" ���܂܂�Ă���ꍇ
                $isFound = $TRUE
                readCells $T_WB $SheetNum "F18" $i 12
                break
            }
        }
    }
}

##--------------------------------------------------------##
## �w��Z���̓��e��\��
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
## �L�[���͂��m�F
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
## Main�֐�
##--------------------------------------------------------##
function Main([string]$Path, [string]$Pattern, [string]$Exclude, [int]$Depth) {
    $list = Get-ChildItem -Recurse -Path $Path -Filter $Pattern -Depth $Depth | ForEach {$_.FullName}
    $hostnameInStr = @("-c", "-PN-c")

    # Excel�I�u�W�F�N�g���擾
    $excel = New-Object -ComObject Excel.Application

    # �\����L����
    $excel.Visible = $TRUE

    foreach ($fname in $list) {

        # ������Excel�t�@�C����ǂݎ���p�ŊJ��
        $T_WB = $excel.Workbooks.Open($fname, $FALSE, $TRUE)

        if ($T_WB.WorkSheets.Count -ge 2) {
            #--------------------------------------------------#
            # �������[�^A
            #--------------------------------------------------#
            Write-Output "---------------------------------------"
            Write-Output $fname
            Write-Output $T_WB.WorkSheets.item(2).Name
            Write-Output "---------------------------------------"
        
            #-----------------------------------------------------#
            # WorkSheets.item(2) �V���Z���^�[SW�ւ̐ڑ��|�[�g���
            #-----------------------------------------------------#
            printCenterSW $T_WB 2 $hostnameInStr 40 
        }

        if ($T_WB.WorkSheets.Count -ge 3) {
            #--------------------------------------------------#
            # �������[�^B
            #--------------------------------------------------#
            Write-Output "---------------------------------------"
            Write-Output $fname
            Write-Output $T_WB.WorkSheets.item(3).Name
            Write-Output "---------------------------------------"

            #-----------------------------------------------------#
            # WorkSheets.item(3) �V���Z���^�[SW�ւ̐ڑ��|�[�g���
            #-----------------------------------------------------#
            printCenterSW $T_WB 3 $hostnameInStr 40 
        }

        $T_WB.Close()
    }
    $excel.Quit()
}


##--------------------------------------------------------##
## Main�Ăяo��
##--------------------------------------------------------##

Main $Path $Pattern $Exclude $Depth
