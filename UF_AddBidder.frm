VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddBidder
   Caption         =   "Add Bidder"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblCompany
      Caption         =   "Company:"
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   1440
   End
   Begin MSForms.TextBox txtCompany
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3960
   End
   Begin MSForms.Label lblContact
      Caption         =   "Contact:"
      Height          =   240
      Left            =   120
      Top             =   480
      Width           =   1440
   End
   Begin MSForms.TextBox txtContact
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   3960
   End
   Begin MSForms.Label lblPhone
      Caption         =   "Phone:"
      Height          =   240
      Left            =   120
      Top             =   840
      Width           =   1440
   End
   Begin MSForms.TextBox txtPhone
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   3960
   End
   Begin MSForms.Label lblEmail
      Caption         =   "Email:"
      Height          =   240
      Left            =   120
      Top             =   1200
      Width           =   1440
   End
   Begin MSForms.TextBox txtEmail
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   3960
   End
   Begin MSForms.Label lblDate
      Caption         =   "Date Received:"
      Height          =   240
      Left            =   120
      Top             =   1560
      Width           =   1440
   End
   Begin MSForms.TextBox txtDate
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   3960
   End
   Begin MSForms.Label lblNotes
      Caption         =   "Notes:"
      Height          =   240
      Left            =   120
      Top             =   1920
      Width           =   1440
   End
   Begin MSForms.TextBox txtNotes
      Height          =   270
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   3960
   End
   Begin MSForms.Label lblBaseBid
      Caption         =   "Base Bid:"
      Height          =   240
      Left            =   120
      Top             =   2280
      Width           =   1440
   End
   Begin MSForms.TextBox txtBaseBid
      Height          =   270
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   3960
   End
   Begin MSForms.CheckBox chkReEnter
      Caption         =   "Re-enter scope + alternates for this bidder now"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   2700
      Value           =   0   'False
      Width           =   5520
   End
   Begin MSForms.CommandButton btnAdd
      Caption         =   "Add Bidder"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   3060
      Width           =   2700
   End
   Begin MSForms.CommandButton btnCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2940
      TabIndex        =   9
      Top             =   3060
      Width           =   2700
   End
End
Attribute VB_Name = "UF_AddBidder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelled As Boolean

Public Property Get Company() As String
    Company = Trim$(txtCompany.Text)
End Property

Public Property Get Contact() As String
    Contact = Trim$(txtContact.Text)
End Property

Public Property Get Phone() As String
    Phone = Trim$(txtPhone.Text)
End Property

Public Property Get Email() As String
    Email = Trim$(txtEmail.Text)
End Property

Public Property Get DateReceived() As String
    DateReceived = Trim$(txtDate.Text)
End Property

Public Property Get Notes() As String
    Notes = Trim$(txtNotes.Text)
End Property

Public Property Get BaseBid() As String
    BaseBid = Trim$(txtBaseBid.Text)
End Property

Public Property Get ReEnterScope() As Boolean
    ReEnterScope = chkReEnter.Value
End Property

Private Sub UserForm_Initialize()
    Cancelled = True
    chkReEnter.Value = False
End Sub

Private Sub btnAdd_Click()
    Cancelled = False
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
