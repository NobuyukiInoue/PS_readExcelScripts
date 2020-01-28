param($Path, $Pattern, $Exclude, $Depth)

##--------------------------------------------------------##
## �����`�F�b�N
##--------------------------------------------------------##

if (-Not $Path) {
    $Path = "../"
}

if (-Not $Pattern) {
    $pattern = "*�ڑ��\���}*.xls*"
}

if (-Not $Exclude) {
    $Exclude = "*org*"
}

if (-Not $Depth) {
    $Depth = 2
}


##--------------------------------------------------------##
## Shape�I�u�W�F�N�g��Name�v���p�e�B��Text�v���p�e�B���擾
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
## �L�[���͂̊m�F
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

    foreach ($fname in $list) {
        Write-Output $fname

        # Excel�I�u�W�F�N�g���擾
        $excel = New-Object -ComObject Excel.Application

        # �\����L����
        $excel.Visible = $TRUE

        # ������Excel�t�@�C����ǂݎ���p�ŊJ��
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
## Main�Ăяo��
##--------------------------------------------------------##

Main $Path $Pattern $Exclude $Depth
