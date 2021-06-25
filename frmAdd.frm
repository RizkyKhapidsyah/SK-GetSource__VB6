VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "http://"
      Top             =   810
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   210
      Width           =   3495
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sName As String
Dim sAddress As String

    sName = txtName.Text
    sAddress = txtAddress.Text

    ' Add the new Item
    With frmMain
        .lsvAddresses.ListItems.Add 1, , sName
        .lsvAddresses.ListItems.Item(1).ListSubItems.Add , , sAddress
    
        ' Enable/Disable Download, Modify, Remove and Clear buttons
        If .lsvAddresses.ListItems.Count > 0 Then
            .cmdDownload.Enabled = True
            .cmdModify.Enabled = True
            .cmdRemove.Enabled = True
            .cmdClear.Enabled = True
        Else
            .cmdDownload.Enabled = False
            .cmdModify.Enabled = False
            .cmdRemove.Enabled = False
            .cmdClear.Enabled = False
        End If
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    ' Center Form
    CenterForm Me
    
    ' Enabe/Disable OK button
    txtName_Change
End Sub

Private Sub txtAddress_Change()
    txtName_Change
End Sub

Private Sub txtAddress_GotFocus()
    SelectAll txtAddress
End Sub

Private Sub txtName_Change()
    ' Enabe/Disable OK button
    If txtName.Text = vbNullString Or txtAddress.Text = vbNullString Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectAll txtName
End Sub
