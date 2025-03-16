Dim shell
Set shell = CreateObject("WScript.Shell")

' Defina o caminho do seu script PowerShell
Dim caminhoScript
caminhoScript = shell.ExpandEnviromentStrings("%USERPROFILE%\Desktop\powershell\baixarGitAwk.ps1")

' Comando para executar o script PowerShell em segundo plano
shell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & caminhoScript & """", 0, False

' Limpeza
Set shell = Nothing
