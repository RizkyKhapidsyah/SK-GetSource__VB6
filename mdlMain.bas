Attribute VB_Name = "mdlMain"
Option Explicit

Public msAppPath As String
Public Const sSeperator = " ***** "
Public sFileName As String

Sub CenterForm(frm As Form)
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
End Sub

Sub SelectAll(tb As TextBox)
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
End Sub

Sub EndApp(Message As Boolean)
    If Message = True Then
        If MsgBox("Are you sure you want to exit GetSource?", vbYesNo + vbQuestion, "Exit") = vbYes Then
            UnloadAllForms
        End If
    Else
        UnloadAllForms
    End If
End Sub

Sub UnloadAllForms()
Dim i As Integer

    On Error Resume Next
    For i = 0 To Forms.Count - 1
        Unload Forms(i)
    Next
End Sub

Sub AddList()
Dim nLeft As Integer
Dim nRight As Integer
Dim nFirstWordLeft As Integer
Dim nFirstWordRight As Integer
Dim nFirstWordLength As Integer
Dim nSecondWordLeft As Integer
Dim nSecondWordRight As Integer
Dim nSecondWordLength As Integer
Dim sWrap As String
Dim sLineOfText As String
Dim sAllText As String
Dim sFileName As String
    

    sFileName = msAppPath & "\" & "lst0"
    If Dir(sFileName) <> "" Then
        On Error Resume Next
        sWrap = Chr(13) + Chr(10)  'create wrap character
        Open sFileName For Input As #1
        Do Until EOF(1)          'then read lines from file
            Line Input #1, sLineOfText
            
            ' Search the string
            '---------------------------
            ' String Start Position
            nLeft = InStr(1, sLineOfText, sSeperator, vbTextCompare)
            ' String width
            nRight = nLeft + Len(sSeperator) - 1
            
            nFirstWordLeft = 1
            nFirstWordRight = nLeft - 1
            nFirstWordLength = nLeft - 1
            
            nSecondWordLeft = nRight + 1
            nSecondWordRight = Len(sLineOfText)
            nSecondWordLength = Len(sLineOfText) - nRight
            
            With frmMain.lsvAddresses
                .ListItems.Add 1, , Left(sLineOfText, nFirstWordLength)
                .ListItems.Item(1).ListSubItems.Add , , Right(sLineOfText, nSecondWordLength)
            End With
            
            sAllText = sAllText & sLineOfText & sWrap
        Loop
        
        Close #1
    End If
End Sub
