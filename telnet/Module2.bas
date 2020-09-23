Attribute VB_Name = "Module2"
Public Sub ReadList(list As ListBox, Filename As String, Optional ClearList As Boolean)
    On Error GoTo Err
    Open Filename For Input As #1
    Do While Not EOF(1)
        Input #1, lstinpuT
        list.AddItem lstinpuT
    Loop
    Close #1
    Exit Sub
Err:
    Exit Sub
End Sub
Public Sub WriteList(list As ListBox, Filename As String)
    If list.ListCount <= 0 Then
        Exit Sub
        End
    End If
    On Error GoTo Err
    Open Filename For Output As #1
    For i = 0 To list.ListCount - 1
        Print #1, list.list(i)
    Next
    Close #1
    Exit Sub
Err:
    MsgBox "Error In WriteList" & Chr(13) & Chr(13) & Err.number _
    & " - " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub



