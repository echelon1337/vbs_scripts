' terminator.vbs
' skrypt zabija procesy zajmujace powyzej x% procesora
' czasami trzeba troche poczeka zanim zacznie dzialac 
' -------------------------------------------------------'

strComputer = "."    ' bez tego mozna sie obejsc
dim proc_to_kill(50) ' tablica z ID procesow do zabicia
proc_number = 0

pow_consumption = InputBox("Enter value:", "Power consumption, default 25%", 25)

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")

' wyciagam wszystkie procesy
Set perfProcessList = objWMIService.ExecQuery(_
    "SELECT * FROM Win32_PerfFormattedData_PerfProc_Process")

For Each process in perfProcessList
    If process.PercentProcessorTime > pow_consumption and _
        process.Name <> "Idle" and process.Name <> "System Idle Process" and _
        process.Name <> "_Total" then

        If proc_number < 50 then
            proc_to_kill(proc_number) = process.IDProcess
            proc_number = proc_number + 1
        End if

    End if
Next

' wyciagam tylko proces o danym ID
' z tej czesci nie jestem dumy

For index = 0 to proc_number
    Set ProcessList = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_Process WHERE ProcessId ='"& proc_to_kill(index) &"'" )

    For Each proc in ProcessList
        proc.terminate()
    Next
Next

' End of List Process 
Wscript.Quit

' https://msdn.microsoft.com/en-us/library/aa394323(v=vs.85).aspx
' https://msdn.microsoft.com/en-us/library/aa394599(v=vs.85).aspx
' https://msdn.microsoft.com/en-us/library/aa394372(v=vs.85).aspx
' https://www.tutorialspoint.com/vbscript/index.htm
' https://www.w3schools.com/asp/asp_looping.asp
