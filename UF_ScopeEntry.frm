VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ScopeEntry
   Caption         =   "Scope Entry"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBidder
      Caption         =   ""
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   5160
   End
   Begin MSForms.Label lblDesc
      Caption         =   ""
      Height          =   840
      Left            =   120
      Top             =   420
      Width           =   5160
      WordWrap        =   -1  'True
   End
   Begin MSForms.Frame fraResponse
      Caption         =   "Bidder's response:"
      Height          =   1680
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5160
      Begin MSForms.OptionButton optIncluded
         Caption         =   "Included  (I)"
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   0   'False
         Width           =   4920
      End
      Begin MSForms.OptionButton optExcluded
         Caption         =   "Excluded  (E)"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Value           =   0   'False
         Width           =   4920
      End
      Begin MSForms.OptionButton optAmount
         Caption         =   "Dollar Amount:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   0   'False
         Width           =   1440
      End
      Begin MSForms.TextBox txtAmount
         Enabled         =   0   'False
         Height          =   270
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   3360
      End
      Begin MSForms.OptionButton optSkip
         Caption         =   "Unconfirmed / Skip"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Value           =   -1  'True
         Width           =   4920
      End
   End
   Begin MSForms.CommandButton btnOK
      Caption         =   "OK -- Next Item"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   3060
      Width           =   2520
   End
   Begin MSForms.CommandButton btnStop
      Caption         =   "Stop -- Done Entering"
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      Top             =   3060
      Width           =   2520
   End
End
Attribute VB_Name = "UF_ScopeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' "I"=included, "E"=excluded, "$"=dollar amount, "U"=unconfirmed/skip, "STOP"=done
Public Result As String
Public AmountText As String

Private Sub UserForm_Initialize()
    Result = "U"
    AmountText = ""
    optSkip.Value = True
    txtAmount.Enabled = False
End Sub

' Call before each Show to load the description for this item
Public Sub SetItem(ByVal bidName As String, ByVal desc As String)
    lblBidder.Caption = "Bidder: " & bidName
    lblDesc.Caption = desc
    optSkip.Value = True
    txtAmount.Enabled = False
    txtAmount.Text = ""
    Result = "U"
    AmountText = ""
End Sub

Private Sub optIncluded_Click()
    txtAmount.Enabled = False
End Sub

Private Sub optExcluded_Click()
    txtAmount.Enabled = False
End Sub

Private Sub optAmount_Click()
    txtAmount.Enabled = True
    txtAmount.SetFocus
End Sub

Private Sub optSkip_Click()
    txtAmount.Enabled = False
End Sub

Private Sub btnOK_Click()
    If optIncluded.Value Then
        Result = "I"
    ElseIf optExcluded.Value Then
        Result = "E"
    ElseIf optAmount.Value Then
        AmountText = Trim$(txtAmount.Text)
        If Len(AmountText) = 0 Then
            MsgBox "Enter an amount, or select a different option.", vbExclamation, "Amount Required"
            txtAmount.SetFocus
            Exit Sub
        End If
        Result = "$"
    Else
        Result = "U"
    End If
    Me.Hide
End Sub

Private Sub btnStop_Click()
    Result = "STOP"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Result = "STOP"
        Cancel = 1
        Me.Hide
    End If
End Sub
