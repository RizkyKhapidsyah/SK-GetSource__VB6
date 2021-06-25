VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GetSource"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   7440
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet netNet 
      Left            =   7320
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   7440
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7435
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":000C
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsvAddresses 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "name"
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Address"
         Text            =   "Address"
         Object.Width           =   10583
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    ' Show Add Form
    frmAdd.Show vbModal, Me
End Sub

Private Sub cmdClear_Click()
    lsvAddresses.ListItems.Clear
    
    ' Enable/Disable Download, Modify, Remove and Clear buttons
    EnableBtns
End Sub

Private Sub cmdDownload_Click()
Dim sAddress As String

    On Error GoTo ErrHandler

    sAddress = lsvAddresses.SelectedItem.ListSubItems(1).Text
    txtSource.Text = netNet.OpenURL(sAddress)
    
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox "Error Opening URL", vbCritical, Err.Description
    End If
End Sub

Private Sub cmdExit_Click()
    EndApp True
End Sub

Private Sub cmdModify_Click()
Dim i As Integer

    For i = 1 To lsvAddresses.ListItems.Count
        If lsvAddresses.ListItems(i).Selected = True Then
            With frmModify
                .txtName.Text = lsvAddresses.ListItems(i).Text
                .txtAddress.Text = lsvAddresses.ListItems(i).ListSubItems(1).Text
            End With
            frmModify.Show vbModal, Me
        End If
    Next
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    
    lsvAddresses.ListItems.Remove lsvAddresses.SelectedItem.Index
    
    ' Enable/Disable Download, Modify, Remove and Clear buttons
    EnableBtns
End Sub

Private Sub cmdSave_Click()
    dlgSave.Filter = "HTML Documents (*.html)|*.html|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgSave.DialogTitle = "Save Source"
    dlgSave.ShowSave
    If Len(dlgSave.FileName) > 0 Then
        Open dlgSave.FileName For Output As #1
        Print #1, txtSource.Text
        Close #1
    End If
End Sub

Private Sub Form_Load()
    ' Center the Form
    CenterForm Me
    
    ' Get the Application Path
    msAppPath = App.Path
    
    ' Add the List Items
    AddList
    
    
    ' Enable/Disable Download, Modify, Remove and Clear buttons
    EnableBtns
End Sub

Sub EnableBtns()
    ' Enable/Disable Download, Modify, Remove and Clear buttons
    If lsvAddresses.ListItems.Count > 0 Then
        cmdDownload.Enabled = True
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
    Else
        cmdDownload.Enabled = False
        cmdModify.Enabled = False
        cmdRemove.Enabled = False
        cmdClear.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sFileName As String
Dim sContents As String
Dim nList As Integer

    sFileName = msAppPath & "\" & "lst0"
    
    On Error Resume Next
    For nList = 1 To lsvAddresses.ListItems.Count
        While nList <= lsvAddresses.ListItems.Count
            sContents = sContents & lsvAddresses.ListItems(nList).Text & sSeperator
            sContents = sContents & lsvAddresses.ListItems(nList).ListSubItems(1).Text & vbCrLf
            nList = nList + 1
        Wend
    Next
    
    Open sFileName For Output As #1
    Print #1, sContents
    Close #1
End Sub
