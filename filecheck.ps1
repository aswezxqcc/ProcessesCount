$path = "C:\Program Files (x86)\CUT\"
$bakPath = "C:\Program Files (x86)\CUT.bak\"
$filesName = "AutoUpdate.exe", "AutoUpdateCopy.exe", "CFETS.NGCNYTS.CM.DownloadManager.dll", "CFETS.NTP.CLT.ClientBridge.dll", "CTAFShell.AutoUpdate.dll", "CTAFShell.ClientBridge.dll", "CTAFShell.Common.dll", "CTAFShell.Common.Plugin.dll", "CTAFShell.DownloadManager.dll", "CTAFShell.exe", "CTAFShell.Utility.dll", "CTAFShell.WidgetManager.dll", "CTAFShell.WinformException.exe", "CTAFShellCore.dll", "log4net.dll", "Newtonsoft.Json.dll", "Microsoft.QualityTools.Testing.Fakes.dll","Plugins\CTAFShell.Plugin.Excel.dll","Plugins\CTAFShell.Plugin.LogContent.dll","Plugins\CTAFShell.Plugin.USBKey.dll","Plugins\Interop.CryptoKitLib.dll"
$result = 1..$filesName.Count

function getMD5 ([string]$filePath, [array]$fileName) {
    for ($n = 0; $n -le ($fileName.Count - 1); $n++) {
        $path1=$filePath+$fileName[$n]
        $result[$n] = (certutil -hashfile "$path1" MD5)[1]
    }
    return $result;
}
$a=getMD5 $path  $filesName
$b=getMD5 $bakPath  $filesName

for ($m=0;$m -le ($a.Count-1);$m++){
    if($a[$m] -eq $b[$m]){
    $true
    }else{
     $filesName[$m]+":"+$false
    }
}