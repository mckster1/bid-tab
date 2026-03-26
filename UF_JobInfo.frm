VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_JobInfo
   Caption         =   "Generate Trade Tabs -- Job Info"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblProjName
      Caption         =   "Project Name:"
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   1980
   End
   Begin MSForms.TextBox txtProjName
      Height          =   270
      Left            =   2220
      TabIndex        =   0
      Top             =   120
      Width           =   3540
   End
   Begin MSForms.Label lblEstimator
      Caption         =   "Estimator:"
      Height          =   240
      Left            =   120
      Top             =   480
      Width           =   1980
   End
   Begin MSForms.TextBox txtEstimator
      Height          =   270
      Left            =   2220
      TabIndex        =   1
      Top             =   480
      Width           =   3540
   End
   Begin MSForms.Label lblBidDate
      Caption         =   "Bid Date:"
      Height          =   240
      Left            =   120
      Top             =   840
      Width           =   1980
   End
   Begin MSForms.TextBox txtBidDate
      Height          =   270
      Left            =   2220
      TabIndex        =   2
      Top             =   840
      Width           =   3540
   End
   Begin MSForms.Label lblJobGSF
      Caption         =   "Job GSF:"
      Height          =   240
      Left            =   120
      Top             =   1200
      Width           =   1980
   End
   Begin MSForms.TextBox txtJobGSF
      Height          =   270
      Left            =   2220
      TabIndex        =   3
      Top             =   1200
      Width           =   3540
   End
   Begin MSForms.Label lblTradeSF
      Caption         =   "Trade SF  (this run only -- not saved):"
      Height          =   240
      Left            =   120
      Top             =   1560
      Width           =   1980
   End
   Begin MSForms.TextBox txtTradeSF
      Height          =   270
      Left            =   2220
      TabIndex        =   4
      Top             =   1560
      Width           =   3540
   End
   Begin MSForms.Frame fraOverwrite
      Caption         =   "Existing tabs -- overwrite scope lines?"
      Height          =   720
      Left            =   120
      TabIndex        =   5
      Top             =   1980
      Width           =   5640
      Begin MSForms.OptionButton optOverwriteNo
         Caption         =   "Seed only NEW tabs  (default)"
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   5400
      End
      Begin MSForms.OptionButton optOverwriteYes
         Caption         =   "Clear + re-seed ALL selected tabs"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Value           =   0   'False
         Width           =   5400
      End
   End
   Begin MSForms.CommandButton btnGenerate
      Caption         =   "Generate Tabs"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2700
   End
   Begin MSForms.CommandButton btnCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2940
      TabIndex        =   7
      Top             =   2760
      Width           =   2700
   End
End
Attribute VB_Name = "UF_JobInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelled As Boolean
Public OverwriteScope As Long  ' 1=yes overwrite, 2=no (new only)

Public Property Get ProjName() As String
    ProjName = Trim$(txtProjName.Text)
End Property

Public Property Get Estimator() As String
    Estimator = Trim$(txtEstimator.Text)
End Property

Public Property Get BidDate() As String
    BidDate = Trim$(txtBidDate.Text)
End Property

Public Property Get JobGSF() As String
    JobGSF = Trim$(txtJobGSF.Text)
End Property

Public Property Get TradeSF() As String
    TradeSF = Trim$(txtTradeSF.Text)
End Property

' Call before Show to pre-fill saved values from Config_CSI
Public Sub Prefill(ByVal projName As String, ByVal estimator As String, _
                   ByVal bidDate As String, ByVal jobGSF As String)
    txtProjName.Text = projName
    txtEstimator.Text = estimator
    txtBidDate.Text = bidDate
    txtJobGSF.Text = jobGSF
End Sub

Private Sub UserForm_Initialize()
    Cancelled = True
    OverwriteScope = 2
    optOverwriteNo.Value = True
End Sub

Private Sub btnGenerate_Click()
    Cancelled = False
    OverwriteScope = IIf(optOverwriteYes.Value, 1, 2)
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Cancel = 1
        Me.Hide
    End If
End Sub
