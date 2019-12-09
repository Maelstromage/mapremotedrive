$letter1 = "T"
$path1 = "\\server\share"
$wshell = new-object -ComObject wscript.shell

$MapDrive1 = "cmd /c start PowerShell -ExecutionPolicy Unrestricted -Command " + '"& {{}Start-Process PowerShell -ArgumentList ' + "'-ExecutionPolicy Unrestricted -command " + '""New-PSDrive -Name ' + $letter1 + ' -PSProvider FileSystem -Root ' + "'" + '"' + "'" + $path1 + "'"+'"' + "' -Persist" + '""' +"'" + '{}}";'


$wshell.run("explorer.exe Shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}")
sleep 1
$wshell.SendKeys($keystosend)
sleep 1
$wshell.sendkeys('~')
