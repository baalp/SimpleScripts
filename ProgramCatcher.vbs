Option Explicit

Dim Shell, ShellApp, FSO, WMI

Set Shell = CreateObject("WScript.Shell")
Set ShellApp = CreateObject("Shell.Application")
Set FSO   = CreateObject("Scripting.FileSystemObject")
Set WMI   = GetObject("winmgmts:\\.\root\cimv2")

Dim ScriptName
ScriptName = "ProgramCatcher"

'Настройки скрипта
Dim ProgramName, CheckFrequency, RunningTime
'Отслеживаемая программа
ProgramName = "notepad.exe"
'Частота проверки (в ms)
CheckFrequency = 1000
'Длительность работы (в s)
'0 = Бесконечно
RunningTime = 0

'---------------------------------------------------------
'Начало главного алгоритма

'Подтверждение запуска
Dim result, message
message = "Данная программа будет, в фоновом режиме, следить за запущенными приложениями. "
message = message + "Она среагирует на запуск " + ProgramName + ". "
message = message + "Чтобы завершить программу, запустите " + ProgramName
message = message + " или завершите задачу wscript.exe из Диспетчера задач. " + vbCr + vbCr
message = message + "Нажмите ОК, чтобы запустить программу." + vbCr + "Нажмите ОТМЕНА или закройте окно для отмены."
result = msgbox(message, 1, ScriptName)
    'OK = 1, Cancel = 2
select case result
    case 1
        'WScript.Sleep 100
    case 2
        WScript.Quit
end select

'Цикл отслеживания
Dim i, iterations
i = 0
if RunningTime > 0 Then
    iterations = int(RunningTime / (CheckFrequency / 1000))
Else
    iterations = 0
End If
Do
    i = i + 1
    WScript.Sleep CheckFrequency
    ProgramCatcher(ProgramName)
    if i >= iterations and iterations > 0 Then
        WScript.Quit
    End If
Loop

'Call SendToast("Shield",5000,"Добро пожаловать!","Мы ждали тебя","None",5)
'Call SendToast("Shield",5000,"Добро пожаловать!","Чувствуй себя, как дома","None",5)

'---------------------------------------------------------
'Поиск запущенного приложения
Sub ProgramCatcher(programName)
    Dim colItems
    Set colItems = WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & ProgramName & "'")
    If colItems.Count > 0 Then
        'Если программа запущена
        ShellApp.MinimizeAll
        'Call SendToast("Shield",5000,"Отслеживание",ProgramName + " запущен","None",5)
        WScript.Quit
    Else
        'WScript.Sleep 100
        'Если программа не запущена
        'Call SendToast("Error",5000,"Отслеживание",ProgramName + " не запущен","None",5)
    End If
End Sub

'---------------------------------------------------------
'Отправка Tosat-уведомления

'Внешняя функция вызова уведомления
' notifyIcon - Иконка около названия приложения
' delayMs - Длительность показа уведомления (Windows может это проигнорировать)
' title - Заголовок уведомления
' text - Текст уведомления
' tipIcon - Иконка левее самого уведомления
' timeoutSec- Длительность жизни скрипта, вызывающего уведомление 
Sub SendToast(notifyIcon, delayMs, title, text, tipIcon, timeoutSec)
    Dim PSPath, Cmd
    'Формируем PS-скрипт рядом с VBS-скриптом
    PSPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, ".")) & "ps1"
    'Формируем текст PS-скрипта
    Call WritePSFile(PSPath, notifyIcon, delayMs, title, text, tipIcon, timeoutSec)
    'Формируем команду запуска PowerShell.exe с созданным PS-скриптом
    Cmd = "cmd /c PowerShell.exe -STA -ExecutionPolicy Bypass -NoProfile -File """ & PSPath & """"
    'Запускаем готовую команду
    Shell.Run Cmd, 0, True
End Sub

'Формирование PS-скрипта для уведомления
Sub WritePSFile(psPath, notifyIcon, delayMs, title, text, tipIcon, timeoutSec)
    Dim f
    Set f = FSO.CreateTextFile(psPath, True)
    f.WriteLine "Add-Type -AssemblyName System.Windows.Forms"
    f.WriteLine "Add-Type -AssemblyName System.Drawing"
    f.WriteLine ""
    f.WriteLine "$notify = New-Object System.Windows.Forms.NotifyIcon"
    f.WriteLine "$notify.Icon = [System.Drawing.SystemIcons]::" & notifyIcon
    f.WriteLine "$notify.Visible = $true"
    f.WriteLine ""
    f.WriteLine "$notify.ShowBalloonTip(" & delayMs & ", """ & EscapePS(title) & """, """ & EscapePS(text) & """, [System.Windows.Forms.ToolTipIcon]::" & tipIcon & ")"
    f.WriteLine ""
    f.WriteLine "[System.Windows.Forms.Application]::DoEvents()"
    f.WriteLine "Start-Sleep -Seconds " & timeoutSec
    f.WriteLine ""
    f.WriteLine "$notify.Dispose()"
    f.WriteLine "$self = $MyInvocation.MyCommand.Path"
    f.WriteLine "Start-Sleep -Milliseconds 200"
    f.WriteLine "Remove-Item -LiteralPath $self -Force"
    f.Close
End Sub

'Функция корректного форматирования строк в команду PS
Function EscapePS(s)
    EscapePS = Replace(s, """", "`""")
End Function

'---------------------------------------------------------