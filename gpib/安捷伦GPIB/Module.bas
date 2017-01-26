Attribute VB_Name = "ModuleTimeDelay"
 Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 'Download by http://www.codefans.net
 Public Function MytimeOUT(interval)  '—” ±≥Ã–Ú
    Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
        Sleep (1)
        DoEvents
    Loop
    
End Function
