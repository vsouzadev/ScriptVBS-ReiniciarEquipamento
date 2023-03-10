'=========================================================================

'Nome: Reiniciar Equipamento
'Função: Reiniciar um computador remotamente, apenas inserindo o nome de rede do equipamento.
'Data: 10/03/2023
'Versão: 1.0
'Autor: Vinicius de Almeida Souza
'GitHub: https://github.com/vsouzadev

'=========================================================================

Set objShell = CreateObject("WScript.Shell")

strComputer = InputBox("Insira o nome do computador a ser reiniciado:", "Reiniciar Equipamento")


If strComputer = "" Then
    MsgBox "O nome do computador e invalido.", vbCritical, "Reiniciar Equipamento"
    WScript.Quit
End If


intAnswer = MsgBox("Tem certeza de que deseja reiniciar o computador " & strComputer & "?", vbQuestion + vbYesNo, "Reiniciar Equipamento")


If intAnswer = vbNo Then
    MsgBox "O comando de reinicializacao foi cancelado pelo usuario.", vbInformation, "Reiniciar Equipamento"
    WScript.Quit
End If


Set objShell = WScript.CreateObject("WScript.Shell")


objShell.Run "shutdown -r -t 0 -f -m \\" & strComputer, 1, True


MsgBox "O computador " & strComputer & " sera reiniciado em breve.", vbInformation, "Reiniciar Equipamento"
