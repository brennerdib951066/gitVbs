Dim shell
Set shell = CreateObject("WScript.Shell")

' Obtém o diretório do usuário dinamicamente
Dim caminhoScript
caminhoScript = shell.ExpandEnvironmentStrings("%USERPROFILE%\Desktop\powershell\abriPowershell.ps1")

' Comando para executar o script PowerShell em segundo plano
shell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & caminhoScript & """", 0, False

' Limpeza
Set shell = Nothing
