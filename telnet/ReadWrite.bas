Attribute VB_Name = "Module1"

Public Sub WriteFile(Text As TextBox, FileName As String, Optional aPPpaTH As Boolean)
On Error Resume Next
    If aPPpaTH = True Then
        Open App.Path & "/" & FileName For Output As #1
            Print #1, Text
        Close #1
    Else
        Open FileName For Output As #1
            Print #1, Text
        Close #1
    End If
End Sub
Public Sub ReadFile(Text As TextBox, FileName As String, Optional aPPpaTH As Boolean, Optional Clearbox As Boolean)
On Error Resume Next
    If Clearbox = True Then
        Text.Text = ""
    End If
    If aPPpaTH = True Then
        Open App.Path & "/" & FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, info
            Text = info
        Loop
    Close #1
Else:
        Open FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, info
            Text = info
        Loop
    Close #1
End If
End Sub
Public Sub Ran(Text As TextBox, List As ListBox, Combo As ComboBox, number As Integer, bY As Integer, Optional X As Boolean)
    If X = True Then Combo.Clear: List.Clear: Text.Text = ""
        Text = (Int(Rnd * number) + bY)
            List.AddItem (Int(Rnd * number) + bY)
              Combo.AddItem (Int(Rnd * number) + bY)
End Sub
Public Sub DisableCTRL(Optional Answer As Boolean)
    If Answer = True Then
        App.TaskVisible = True
            Else
                App.TaskVisible = False
            End If
End Sub
Public Sub DeleteFile(FileName As String, Optional AreYousure As Boolean)
    If AreYousure = True Then
        kill FileName
            Else
        End If
End Sub
Public Sub Rename(Named1 As String, Named2 As String, Optional AreYousure As Boolean)
    If AreYousure = True Then
        Name Named1 As Named2
            Else
    End If
End Sub
Public Sub MakeDIR(DIR As String, Optional aPPpaTH As Boolean, Optional AreYousure As Boolean)
    If AreYousure = True Then
        If aPPpaTH = True Then
            MkDir App.Path & DIR
                ElseIf aPPpaTH = False Then
            MkDir DIR
        End If
    End If
End Sub
Public Sub Opened(Msg As String, Title As String, Optional vbC As Boolean, Optional vbI As Boolean)
    If App.PrevInstance = True Then
        If vbC = True Then
            vbI = False
                MsgBox Msg, vbCritical, Title
                    End
        If vbI = True Then
            vbC = False
                MsgBox Msg, vbInformation, Title
                    End
            End If
        End If
    End If
End Sub
