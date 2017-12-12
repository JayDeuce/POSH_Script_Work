$name = $env:COMPUTERNAME
$process = Get-Process | Select-Object -Property Name,id,Company, Description,path | Select-Object -First 20 | ConvertTo-Html
#$CpuUsage = Get-Counter -ComputerName localhost '\Process(*)\% Processor Time' | Select-Object -Property countersamples | Select-Object -Property instancename, cookedvalue | Sort-Object -Property cookedvalue -Descending | Select-Object -First 20 | ConvertTo-Html
$cpuUsage = Get-WmiObject win32_processor | Select-Object LoadPercentage | ConvertTo-Html
#$memUsage =

$head = Get-Content .\head.html
$main = @"
<body><h2>$name</h2>
     <p> CPU Usage: <br><br> $($cpuUsage+'%') </p>
     <hr>
     <p> Processes: <br><br> $($process) </p>
</body>
"@
$tail = Get-Content .\tail.html
















$html = $head + $main + $tail

$html | Out-File -FilePath C:\temp\test.html