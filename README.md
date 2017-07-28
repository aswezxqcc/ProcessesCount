## ProcessesCount
使用powershell获取Windows进程的CPU 内存使用情况，并输出到excel中
```powershell
$appname = "TIM" #进程简称
$xc = 4 #CPU线程数
$line = 50#行数
$zhouqi = 2  #检测的周期  单位s
```
## 环境要求
#### 1.powershell 4.0以上    
查看当前版本  
```powershell
$psversiontable.psversion
``` 
win8.1以及Windows server 2012 R2自带4.0 可以不升级 
win10自带5.0 不用升级，  
win7用户需要升级到5.0   
##### powershell 5.0 需要Windows Management Framework 5.0  
* https://www.microsoft.com/en-us/download/details.aspx?id=50395   
#### 2.完整版Microsoft office
一般情况下安装的正版office即可  
验证方式 ：运行 
```powershell
$excelApp= New-Object -ComObject Excel.Application 
```
没有报错 OK!
