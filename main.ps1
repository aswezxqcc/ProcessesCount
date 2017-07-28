$appname = "tim" #进程简称
$xc = 4 #CPU线程数
$line = 10#行数
$zhouqi = 2 #取样周期   
$data = 1..$line
$path=[Environment]::GetFolderPath("Desktop")#保存路径，此路径为桌面

function getDate () {
   return   Get-Date -Format('hhmm')
   
}
function getCPU ([string]$iProcess) {
    $z = "*$iProcess*" #带'*'用于模糊查找，比如一些列进程名字不通，eg：MicrosoftEdge，MicrosoftEdgeCP
    $process1 = Get-Process $z
 
    $a = $process1.CPU  
    sleep $zhouqi
    $process2 = Get-Process $z
    $b = $process2.CPU  
    $d = $process2.ProcessName 
    $f = $process2.WS
 
    $c = New-Object object
    Add-Member -InputObject $c -Name prea -Value "$a" -MemberType NoteProperty;
    Add-Member -InputObject $c -Name preb -Value "$b" -MemberType NoteProperty;    
    Add-Member -InputObject $c -Name pro -Value "$d" -MemberType NoteProperty;
    Add-Member -InputObject $c -Name mem -Value "$f" -MemberType NoteProperty;   

    return $c
    
}
function  getData($a, $b) {
    $dataobject = New-Object object 
    $time=getDate
    Add-Member -InputObject $dataobject -Name time -Value " $time " -MemberType NoteProperty;    
    Add-Member -InputObject $dataobject -Name mem -Value " $a " -MemberType NoteProperty;
    Add-Member -InputObject $dataobject -Name pre -Value " $b " -MemberType NoteProperty;
  
    return $dataobject
}
for ($m = 0; $m -le ($line - 1); $m++) {
    $return = (getCPU $appname)
    $arraya = $return.prea
    $arrayb = $return.preb
    $pro = ($return.pro -split " ")[0]
    $mem = ($return.mem -split " ")
    $newarraya = ($arraya -split " ") 
    $newarrayb = ($arrayb -split " ") 
    $pre = 0..($newarraya.Count - 1)

    for ($i = 0; $i -le ($newarraya.Count - 1); $i++) {
        $pre[$i] = ($newarrayb[$i] - $newarraya[$i]) / $zhouqi * 100 / $xc;
    }
    $preT = 0
    for ($n = 0; $n -le ($pre.Count - 1); $n++) {
        $tmp = $pre[$n]
        $preT += $tmp

    }
 
    $memT = 0
    for ($n = 0; $n -le ($mem.Count - 1); $n++) {
        $tmp = $mem[$n]
        $memT += $tmp

    }  
    $data[$m] = getData $memT $preT 
    $data[$m] #可注释掉，用来显示当前内存和CPU使用率
}
    


# excel start
$excel = New-Object -ComObject Excel.Application 
$workbook = $excel.Workbooks.add()
$sheet = $workbook.worksheets.Item(1)
$workbook.WorkSheets.item(1).Name = $pro 



$x = 2

$lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
$colorIndex = "microsoft.office.interop.excel.xlColorIndex" -as [type]
$borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
$chartType = "microsoft.office.interop.excel.xlChartType" -as [type]

for ($b = 1 ; $b -le 3 ; $b++) {
    $sheet.cells.item(1, $b).font.bold = $true
    $sheet.cells.item(1, $b).borders.LineStyle = $lineStyle::xlDashDot
    $sheet.cells.item(1, $b).borders.ColorIndex = $colorIndex::xlColorIndexAutomatic
    $sheet.cells.item(1, $b).borders.weight = $borderWeight::xlMedium
}

$sheet.cells.item(1, 1) = "title   time"
$sheet.cells.item(1, 2) = "WS memory(MB)"
$sheet.cells.item(1, 3) = "CPU(%)" 


foreach ($process in $data) {
    $sheet.cells.item($x, 1) = $process.time
    $sheet.cells.item($x, 2) = $process.mem/1024/1024
    $sheet.cells.item($x, 3) = $process.pre

    $x++
} 

$range = $sheet.usedRange
$range.EntireColumn.AutoFit() | out-null
$excel.Visible = $true
$filename= $appname+'-'+(Get-Date -Format 'MMddhhmm')+'.xlsx' #excel文件名
$excel.ActiveWorkBook.SaveAs($path+'/'+$filename)#保存excel
#$excel.quit()