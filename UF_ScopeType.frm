VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ScopeType
   Caption         =   "Add Scope Line"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblPrompt
      Caption         =   "Choose type for this scope line:"
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   4560
   End
   Begin MSForms.Label lblDesc
      Caption         =   ""
      Height          =   660
      Left            =   120
      Top             =   420
      Width           =   4560
   End
   Begin MSForms.CommandButton btnNormal
      Caption         =   "Normal Scope Line"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   4560
   End
   Begin MSForms.CommandButton btnException
      Caption         =   "Exception / Exclusion  (red row)"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4560
   End
   Begin MSForms.CommandButton btnCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel -- Don't Add This Line"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4560
   End
End
Attribute VB_Name = "UF_ScopeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 0 = cancel/no action, 1 = normal scope line, 2 = exception/exclusion
Public Result As Integer

Public Sub SetDesc(ByVal desc As String)
    lblDesc.Caption = desc
End Sub

Private Sub UserForm_Initialize()
    lblDesc.WordWrap = True
    Result = 0
End Sub

Private Sub btnNormal_Click()
    Result = 1
    Me.Hide
End Sub

Private Sub btnException_Click()
    Result = 2
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Result = 0
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Result = 0
        Cancel = 1
        Me.Hide
    End If
End Sub
