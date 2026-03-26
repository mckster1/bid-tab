Attribute VB_Name = "TSC_v10_CreateForms"
Option Explicit

'=============================================================================
' Run CreateUserForms_v10 once to build all 6 UserForms.
' Delete this module afterward.
'=============================================================================

Private mCode As String
Private mStep As String          ' tracks progress for error messages

Private Sub L(ByVal s As String)
    mCode = mCode & s & vbCrLf
End Sub

' ---- control / form helpers ----

Private Sub Zap(ByVal vbp As Object, ByVal n As String)
    On Error Resume Next: vbp.VBComponents.Remove vbp.VBComponents(n): On Error GoTo 0
End Sub

Private Function MkForm(ByVal vbp As Object, ByVal n As String, _
                         ByVal cap As String, ByVal W As Single, ByVal H As Single) As Object
    mStep = "Create form: " & n
    Zap vbp, n
    Dim c As Object: Set c = vbp.VBComponents.Add(3)
    c.Name = n
    On Error Resume Next           ' Width/Height not settable in all Excel builds
    c.Designer.Caption = cap
    c.Designer.Width = W
    c.Designer.Height = H
    On Error GoTo 0
    Set MkForm = c
End Function

Private Sub Lbl(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal lft As Single, ByVal tp As Single, ByVal W As Single, ByVal H As Single)
    mStep = "Add label: " & n
    Dim c As Object: Set c = d.Controls.Add("Forms.Label.1", n)
    c.Caption = cap: c.Left = lft: c.Top = tp: c.Width = W: c.Height = H
End Sub

Private Sub Txt(ByVal d As Object, ByVal n As String, _
                ByVal lft As Single, ByVal tp As Single, ByVal W As Single, ByVal H As Single)
    mStep = "Add textbox: " & n
    Dim c As Object: Set c = d.Controls.Add("Forms.TextBox.1", n)
    c.Left = lft: c.Top = tp: c.Width = W: c.Height = H
End Sub

Private Sub Btn(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal lft As Single, ByVal tp As Single, ByVal W As Single, ByVal H As Single)
    mStep = "Add button: " & n
    Dim c As Object: Set c = d.Controls.Add("Forms.CommandButton.1", n)
    c.Caption = cap: c.Left = lft: c.Top = tp: c.Width = W: c.Height = H
End Sub

Private Sub Opt(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal lft As Single, ByVal tp As Single, ByVal W As Single, ByVal H As Single)
    mStep = "Add option: " & n
    Dim c As Object: Set c = d.Controls.Add("Forms.OptionButton.1", n)
    c.Caption = cap: c.Left = lft: c.Top = tp: c.Width = W: c.Height = H
End Sub

Private Sub Chk(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal lft As Single, ByVal tp As Single, ByVal W As Single, ByVal H As Single)
    mStep = "Add checkbox: " & n
    Dim c As Object: Set c = d.Controls.Add("Forms.CheckBox.1", n)
    c.Caption = cap: c.Left = lft: c.Top = tp: c.Width = W: c.Height = H
End Sub

Private Sub SetCode(ByVal comp As Object)
    mStep = "Write code to: " & comp.Name
    Dim m As Object: Set m = comp.CodeModule
    If m.CountOfLines > 0 Then m.DeleteLines 1, m.CountOfLines
    m.AddFromString mCode
    mCode = ""
End Sub

'=============================================================================
' ENTRY POINT
'=============================================================================
Public Sub CreateUserForms_v10()
    On Error GoTo ErrHandler
    Dim vbp As Object: Set vbp = ThisWorkbook.VBProject
    BuildScopeType vbp
    BuildScopeEntry vbp
    BuildAddBidder vbp
    BuildAddBidder2 vbp
    BuildJobInfo vbp
    BuildJobInfo2 vbp
    MsgBox "6 UserForms created." & vbCrLf & _
           "You may now delete the TSC_v10_CreateForms module.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Failed at step: " & mStep, vbCritical
End Sub

'=============================================================================
' UF_ScopeType  (5 controls)
'=============================================================================
Private Sub BuildScopeType(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_ScopeType", "Add Scope Line", 306, 190)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblPrompt",   "Choose type for this scope line:", 6,   6, 288, 14
    Lbl d, "lblDesc",     "",                                 6,  24, 288, 36
    Btn d, "btnNormal",   "Normal Scope Line",                6,  78, 288, 24
    Btn d, "btnException","Exception / Exclusion  (red row)", 6, 108, 288, 24
    Btn d, "btnCancel",   "Cancel -- Don't Add This Line",    6, 138, 288, 24

    mCode = ""
    L "Option Explicit"
    L ""
    L "Public Result As Integer"
    L ""
    L "Public Sub SetDesc(ByVal desc As String)"
    L "    lblDesc.Caption = desc"
    L "    lblDesc.WordWrap = True"
    L "End Sub"
    L ""
    L "Private Sub UserForm_Initialize()"
    L "    Result = 0"
    L "    btnNormal.Default = True"
    L "    btnCancel.Cancel = True"
    L "End Sub"
    L ""
    L "Private Sub btnNormal_Click()"
    L "    Result = 1"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnException_Click()"
    L "    Result = 2"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnCancel_Click()"
    L "    Result = 0"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    L "    If CloseMode = vbFormControlMenu Then"
    L "        Result = 0"
    L "        Cancel = 1"
    L "        Me.Hide"
    L "    End If"
    L "End Sub"
    SetCode comp
End Sub

'=============================================================================
' UF_ScopeEntry  (9 controls — no Frame)
'=============================================================================
Private Sub BuildScopeEntry(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_ScopeEntry", "Scope Entry", 330, 240)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblBidder",   "",                    6,   6, 318, 14
    Lbl d, "lblDesc",     "",                    6,  24, 318, 42
    Opt d, "optIncluded", "Included  (I)",       6,  78, 318, 16
    Opt d, "optExcluded", "Excluded  (E)",       6, 100, 318, 16
    Opt d, "optAmount",   "Dollar Amount:",      6, 122, 144, 16
    Txt d, "txtAmount",                        156, 122, 162, 16
    Opt d, "optSkip",     "Unconfirmed / Skip",  6, 144, 318, 16
    Btn d, "btnOK",   "OK -- Next Item",         6, 180, 150, 24
    Btn d, "btnStop", "Stop -- Done Entering",  162, 180, 150, 24

    mCode = ""
    L "Option Explicit"
    L ""
    L "Public Result As String"
    L "Public AmountText As String"
    L ""
    L "Private Sub UserForm_Initialize()"
    L "    lblDesc.WordWrap = True"
    L "    Result = ""U"""
    L "    AmountText = """""
    L "    optSkip.Value = True"
    L "    txtAmount.Enabled = False"
    L "    btnOK.Default = True"
    L "End Sub"
    L ""
    L "Public Sub SetItem(ByVal bidName As String, ByVal desc As String)"
    L "    lblBidder.Caption = ""Bidder: "" & bidName"
    L "    lblDesc.Caption = desc"
    L "    optSkip.Value = True"
    L "    txtAmount.Enabled = False"
    L "    txtAmount.Text = """""
    L "    Result = ""U"""
    L "    AmountText = """""
    L "End Sub"
    L ""
    L "Private Sub optIncluded_Click()"
    L "    txtAmount.Enabled = False"
    L "End Sub"
    L ""
    L "Private Sub optExcluded_Click()"
    L "    txtAmount.Enabled = False"
    L "End Sub"
    L ""
    L "Private Sub optSkip_Click()"
    L "    txtAmount.Enabled = False"
    L "End Sub"
    L ""
    L "Private Sub optAmount_Click()"
    L "    txtAmount.Enabled = True"
    L "    txtAmount.SetFocus"
    L "End Sub"
    L ""
    L "Private Sub btnOK_Click()"
    L "    If optIncluded.Value Then"
    L "        Result = ""I"""
    L "    ElseIf optExcluded.Value Then"
    L "        Result = ""E"""
    L "    ElseIf optAmount.Value Then"
    L "        AmountText = Trim$(txtAmount.Text)"
    L "        If Len(AmountText) = 0 Then"
    L "            MsgBox ""Enter an amount, or choose a different option."", vbExclamation, ""Amount Required"""
    L "            txtAmount.SetFocus"
    L "            Exit Sub"
    L "        End If"
    L "        Result = ""$"""
    L "    Else"
    L "        Result = ""U"""
    L "    End If"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnStop_Click()"
    L "    Result = ""STOP"""
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    L "    If CloseMode = vbFormControlMenu Then"
    L "        Result = ""STOP"""
    L "        Cancel = 1"
    L "        Me.Hide"
    L "    End If"
    L "End Sub"
    SetCode comp
End Sub

'=============================================================================
' UF_AddBidder  page 1/2  (10 controls) — Company, Contact, Phone, Email
'=============================================================================
Private Sub BuildAddBidder(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_AddBidder", "Add Bidder (1 of 2)", 318, 168)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblCompany", "Company:", 6,  6, 84, 14
    Txt d, "txtCompany",            96,  6, 216, 16
    Lbl d, "lblContact", "Contact:", 6, 30, 84, 14
    Txt d, "txtContact",            96, 30, 216, 16
    Lbl d, "lblPhone",   "Phone:",   6, 54, 84, 14
    Txt d, "txtPhone",              96, 54, 216, 16
    Lbl d, "lblEmail",   "Email:",   6, 78, 84, 14
    Txt d, "txtEmail",              96, 78, 216, 16
    Btn d, "btnNext",   "Next ->",   6, 114, 144, 24
    Btn d, "btnCancel", "Cancel",  162, 114, 144, 24

    mCode = ""
    L "Option Explicit"
    L ""
    L "Public Cancelled As Boolean"
    L ""
    L "Public Property Get Company() As String"
    L "    Company = Trim$(txtCompany.Text)"
    L "End Property"
    L ""
    L "Public Property Get Contact() As String"
    L "    Contact = Trim$(txtContact.Text)"
    L "End Property"
    L ""
    L "Public Property Get Phone() As String"
    L "    Phone = Trim$(txtPhone.Text)"
    L "End Property"
    L ""
    L "Public Property Get Email() As String"
    L "    Email = Trim$(txtEmail.Text)"
    L "End Property"
    L ""
    L "Private Sub UserForm_Initialize()"
    L "    Cancelled = True"
    L "    btnNext.Default = True"
    L "    btnCancel.Cancel = True"
    L "End Sub"
    L ""
    L "Private Sub btnNext_Click()"
    L "    Cancelled = False"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnCancel_Click()"
    L "    Cancelled = True"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    L "    If CloseMode = vbFormControlMenu Then"
    L "        Cancelled = True"
    L "        Cancel = 1"
    L "        Me.Hide"
    L "    End If"
    L "End Sub"
    SetCode comp
End Sub

'=============================================================================
' UF_AddBidder2  page 2/2  (9 controls) — Date, Notes, BaseBid, ReEnter
'=============================================================================
Private Sub BuildAddBidder2(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_AddBidder2", "Add Bidder (2 of 2)", 318, 168)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblDate",    "Date Received:", 6,  6, 84, 14
    Txt d, "txtDate",                     96,  6, 216, 16
    Lbl d, "lblNotes",   "Notes:",         6, 30, 84, 14
    Txt d, "txtNotes",                    96, 30, 216, 16
    Lbl d, "lblBaseBid", "Base Bid:",      6, 54, 84, 14
    Txt d, "txtBaseBid",                  96, 54, 216, 16
    Chk d, "chkReEnter", "Re-enter scope + alternates for this bidder now", 6, 84, 300, 16
    Btn d, "btnAdd",    "Add Bidder", 6,  114, 144, 24
    Btn d, "btnCancel", "Cancel",   162,  114, 144, 24

    mCode = ""
    L "Option Explicit"
    L ""
    L "Public Cancelled As Boolean"
    L ""
    L "Public Property Get DateReceived() As String"
    L "    DateReceived = Trim$(txtDate.Text)"
    L "End Property"
    L ""
    L "Public Property Get Notes() As String"
    L "    Notes = Trim$(txtNotes.Text)"
    L "End Property"
    L ""
    L "Public Property Get BaseBid() As String"
    L "    BaseBid = Trim$(txtBaseBid.Text)"
    L "End Property"
    L ""
    L "Public Property Get ReEnterScope() As Boolean"
    L "    ReEnterScope = chkReEnter.Value"
    L "End Property"
    L ""
    L "Private Sub UserForm_Initialize()"
    L "    Cancelled = True"
    L "    chkReEnter.Value = False"
    L "    btnAdd.Default = True"
    L "    btnCancel.Cancel = True"
    L "End Sub"
    L ""
    L "Private Sub btnAdd_Click()"
    L "    Cancelled = False"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnCancel_Click()"
    L "    Cancelled = True"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    L "    If CloseMode = vbFormControlMenu Then"
    L "        Cancelled = True"
    L "        Cancel = 1"
    L "        Me.Hide"
    L "    End If"
    L "End Sub"
    SetCode comp
End Sub

'=============================================================================
' UF_JobInfo  page 1/2  (10 controls) — ProjName, Estimator, BidDate, JobGSF
'=============================================================================
Private Sub BuildJobInfo(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_JobInfo", "Generate Trade Tabs (1 of 2)", 348, 168)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblProjName",  "Project Name:", 6,  6, 114, 14
    Txt d, "txtProjName",                 126,  6, 210, 16
    Lbl d, "lblEstimator", "Estimator:",    6, 30, 114, 14
    Txt d, "txtEstimator",                126, 30, 210, 16
    Lbl d, "lblBidDate",   "Bid Date:",     6, 54, 114, 14
    Txt d, "txtBidDate",                  126, 54, 210, 16
    Lbl d, "lblJobGSF",    "Job GSF:",      6, 78, 114, 14
    Txt d, "txtJobGSF",                   126, 78, 210, 16
    Btn d, "btnNext",   "Next ->",          6, 114, 144, 24
    Btn d, "btnCancel", "Cancel",         162, 114, 144, 24

    mCode = ""
    L "Option Explicit"
    L ""
    L "Public Cancelled As Boolean"
    L ""
    L "Public Property Get ProjName() As String"
    L "    ProjName = Trim$(txtProjName.Text)"
    L "End Property"
    L ""
    L "Public Property Get Estimator() As String"
    L "    Estimator = Trim$(txtEstimator.Text)"
    L "End Property"
    L ""
    L "Public Property Get BidDate() As String"
    L "    BidDate = Trim$(txtBidDate.Text)"
    L "End Property"
    L ""
    L "Public Property Get JobGSF() As String"
    L "    JobGSF = Trim$(txtJobGSF.Text)"
    L "End Property"
    L ""
    L "Public Sub Prefill(ByVal pName As String, ByVal pEst As String, ByVal pDate As String, ByVal pGSF As String)"
    L "    txtProjName.Text  = pName"
    L "    txtEstimator.Text = pEst"
    L "    txtBidDate.Text   = pDate"
    L "    txtJobGSF.Text    = pGSF"
    L "End Sub"
    L ""
    L "Private Sub UserForm_Initialize()"
    L "    Cancelled = True"
    L "    btnNext.Default = True"
    L "    btnCancel.Cancel = True"
    L "End Sub"
    L ""
    L "Private Sub btnNext_Click()"
    L "    Cancelled = False"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnCancel_Click()"
    L "    Cancelled = True"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    L "    If CloseMode = vbFormControlMenu Then"
    L "        Cancelled = True"
    L "        Cancel = 1"
    L "        Me.Hide"
    L "    End If"
    L "End Sub"
    SetCode comp
End Sub

'=============================================================================
' UF_JobInfo2  page 2/2  (6 controls) — TradeSF, overwrite options, Generate
'=============================================================================
Private Sub BuildJobInfo2(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_JobInfo2", "Generate Trade Tabs (2 of 2)", 348, 150)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblTradeSF",      "Trade SF  (this run only -- not saved):", 6,  6, 210, 14
    Txt d, "txtTradeSF",                                                222,  6, 114, 16
    Opt d, "optOverwriteNo",  "Seed only NEW tabs  (default)",           6, 30, 330, 16
    Opt d, "optOverwriteYes", "Clear + re-seed ALL selected tabs",       6, 54, 330, 16
    Btn d, "btnGenerate", "Generate Tabs", 6,  90, 144, 24
    Btn d, "btnCancel",   "Cancel",       162,  90, 144, 24

    mCode = ""
    L "Option Explicit"
    L ""
    L "Public Cancelled As Boolean"
    L "Public OverwriteScope As Long"
    L ""
    L "Public Property Get TradeSF() As String"
    L "    TradeSF = Trim$(txtTradeSF.Text)"
    L "End Property"
    L ""
    L "Private Sub UserForm_Initialize()"
    L "    Cancelled = True"
    L "    OverwriteScope = 2"
    L "    optOverwriteNo.Value = True"
    L "    btnGenerate.Default = True"
    L "    btnCancel.Cancel = True"
    L "End Sub"
    L ""
    L "Private Sub btnGenerate_Click()"
    L "    Cancelled = False"
    L "    OverwriteScope = IIf(optOverwriteYes.Value, 1, 2)"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub btnCancel_Click()"
    L "    Cancelled = True"
    L "    Me.Hide"
    L "End Sub"
    L ""
    L "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    L "    If CloseMode = vbFormControlMenu Then"
    L "        Cancelled = True"
    L "        Cancel = 1"
    L "        Me.Hide"
    L "    End If"
    L "End Sub"
    SetCode comp
End Sub
