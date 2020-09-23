Attribute VB_Name = "mdlList"
Option Explicit

Public Sub LoadList(sLocation As String, lstListBox As ListBox)
On Error GoTo dlgerror
Dim sCurrent As String
Dim I As Integer



lstListBox.Clear

Open sLocation For Input As #1

I = 0




Do Until EOF(1)

Line Input #1, sCurrent

lstListBox.AddItem sCurrent, I

I = I + 1


Loop

Close #1

MsgBox "List Loaded"


Exit Sub


dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub

Public Sub SaveList(sLocation As String, lstListBox As ListBox)
On Error GoTo dlgerror

Dim sCurrent As String
Dim I As Integer

Open sLocation For Output As #1


I = 0

Do Until I = lstListBox.ListCount

sCurrent = lstListBox.LIST(I)

Print #1, sCurrent

I = I + 1


Loop


Close #1

MsgBox "List Saved"

Exit Sub

dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub
