VERSION 5.00
Begin VB.Form frmModify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmModify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1470
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   270
      Width           =   855
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   870
      Width           =   855
   End
End
Attribute VB_Name = "frmModify"
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
    
    If (Len(sName) > 0) And (Len(sAddress) > 0) Then
        With frmMain
            .lsvAddresses.SelectedItem.Text = sName
            .lsvAddresses.SelectedItem.ListSubItems(1).Text = sAddress
        End With
    End If
    
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

