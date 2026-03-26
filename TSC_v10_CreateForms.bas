Attribute VB_Name = "TSC_v10_CreateForms"
Option Explicit

'=============================================================================
' Run CreateUserForms_v10 once to build all 6 UserForms.
' Delete this module after running.
'=============================================================================
Public Sub CreateUserForms_v10()
    Dim vbp As Object: Set vbp = ThisWorkbook.VBProject
    BuildScopeType vbp
    BuildScopeEntry vbp
    BuildAddBidder vbp
    BuildAddBidder2 vbp
    BuildJobInfo vbp
    BuildJobInfo2 vbp
    MsgBox "6 UserForms created." & vbCrLf & _
           "You may now delete the TSC_v10_CreateForms module.", vbInformation
End Sub

' ---- shared helpers ----

Private Sub Zap(ByVal vbp As Object, ByVal n As String)
    On Error Resume Next: vbp.VBComponents.Remove vbp.VBComponents(n): On Error GoTo 0
End Sub

Private Function MkForm(ByVal vbp As Object, ByVal n As String, _
                         ByVal cap As String, ByVal W As Single, ByVal H As Single) As Object
    Zap vbp, n
    Dim c As Object: Set c = vbp.VBComponents.Add(3)   ' 3 = vbext_ct_MSForm
    c.Name = n
    c.Designer.Caption = cap
    c.Designer.Width = W
    c.Designer.Height = H
    Set MkForm = c
End Function

Private Sub Lbl(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single)
    Dim c As Object: Set c = d.Controls.Add("Forms.Label.1", n, True)
    c.Caption = cap: c.Left = L: c.Top = T: c.Width = W: c.Height = H
End Sub

Private Sub Txt(ByVal d As Object, ByVal n As String, _
                ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single)
    Dim c As Object: Set c = d.Controls.Add("Forms.TextBox.1", n, True)
    c.Left = L: c.Top = T: c.Width = W: c.Height = H
End Sub

Private Sub Btn(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, _
                Optional ByVal isDef As Boolean = False, Optional ByVal isCan As Boolean = False)
    Dim c As Object: Set c = d.Controls.Add("Forms.CommandButton.1", n, True)
    c.Caption = cap: c.Left = L: c.Top = T: c.Width = W: c.Height = H
    If isDef Then c.Default = True
    If isCan Then c.Cancel = True
End Sub

Private Sub Opt(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single)
    Dim c As Object: Set c = d.Controls.Add("Forms.OptionButton.1", n, True)
    c.Caption = cap: c.Left = L: c.Top = T: c.Width = W: c.Height = H
End Sub

Private Sub Chk(ByVal d As Object, ByVal n As String, ByVal cap As String, _
                ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single)
    Dim c As Object: Set c = d.Controls.Add("Forms.CheckBox.1", n, True)
    c.Caption = cap: c.Left = L: c.Top = T: c.Width = W: c.Height = H
End Sub

Private Sub SetCode(ByVal comp As Object, ByVal src As String)
    Dim m As Object: Set m = comp.CodeModule
    If m.CountOfLines > 0 Then m.DeleteLines 1, m.CountOfLines
    m.AddFromString src
End Sub

' ============================================================
' UF_ScopeType  (5 controls)
' Prompt + desc label + 3 buttons
' ============================================================
Private Sub BuildScopeType(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_ScopeType", "Add Scope Line", 306, 190)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblPrompt", "Choose type for this scope line:", 6, 6, 288, 14
    Lbl d, "lblDesc", "", 6, 24, 288, 36
    Btn d, "btnNormal",    "Normal Scope Line",              6,  78, 288, 24, True,  False
    Btn d, "btnException", "Exception / Exclusion  (red row)", 6, 108, 288, 24, False, False
    Btn d, "btnCancel",    "Cancel -- Don't Add This Line",  6, 138, 288, 24, False, True
    SetCode comp, Join(Array( _
        "Option Explicit", _
        "", _
        "Public Result As Integer", _
        "", _
        "Public Sub SetDesc(ByVal desc As String)", _
        "    lblDesc.Caption = desc", _
        "    lblDesc.WordWrap = True", _
        "End Sub", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Result = 0", _
        "End Sub", _
        "", _
        "Private Sub btnNormal_Click()", _
        "    Result = 1", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnException_Click()", _
        "    Result = 2", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnCancel_Click()", _
        "    Result = 0", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", _
        "    If CloseMode = vbFormControlMenu Then", _
        "        Result = 0", _
        "        Cancel = 1", _
        "        Me.Hide", _
        "    End If", _
        "End Sub" _
    ), vbCrLf)
End Sub

' ============================================================
' UF_ScopeEntry  (9 controls — no Frame)
' ============================================================
Private Sub BuildScopeEntry(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_ScopeEntry", "Scope Entry", 330, 240)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblBidder",  "", 6,   6, 318, 14
    Lbl d, "lblDesc",    "", 6,  24, 318, 42
    Opt d, "optIncluded",  "Included  (I)",       6,  78, 318, 16
    Opt d, "optExcluded",  "Excluded  (E)",       6, 100, 318, 16
    Opt d, "optAmount",    "Dollar Amount:",       6, 122, 144, 16
    Txt d, "txtAmount",                          156, 122, 162, 16
    Opt d, "optSkip",      "Unconfirmed / Skip",  6, 144, 318, 16
    Btn d, "btnOK",   "OK -- Next Item",         6, 180, 150, 24, True,  False
    Btn d, "btnStop", "Stop -- Done Entering",  162, 180, 150, 24, False, False
    SetCode comp, Join(Array( _
        "Option Explicit", _
        "", _
        "Public Result As String", _
        "Public AmountText As String", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    lblDesc.WordWrap = True", _
        "    Result = ""U""", _
        "    AmountText = """"", _
        "    optSkip.Value = True", _
        "    txtAmount.Enabled = False", _
        "End Sub", _
        "", _
        "Public Sub SetItem(ByVal bidName As String, ByVal desc As String)", _
        "    lblBidder.Caption = ""Bidder: "" & bidName", _
        "    lblDesc.Caption = desc", _
        "    optSkip.Value = True", _
        "    txtAmount.Enabled = False", _
        "    txtAmount.Text = """"", _
        "    Result = ""U""", _
        "    AmountText = """"", _
        "End Sub", _
        "", _
        "Private Sub optIncluded_Click()", _
        "    txtAmount.Enabled = False", _
        "End Sub", _
        "", _
        "Private Sub optExcluded_Click()", _
        "    txtAmount.Enabled = False", _
        "End Sub", _
        "", _
        "Private Sub optAmount_Click()", _
        "    txtAmount.Enabled = True", _
        "    txtAmount.SetFocus", _
        "End Sub", _
        "", _
        "Private Sub optSkip_Click()", _
        "    txtAmount.Enabled = False", _
        "End Sub", _
        "", _
        "Private Sub btnOK_Click()", _
        "    If optIncluded.Value Then", _
        "        Result = ""I""", _
        "    ElseIf optExcluded.Value Then", _
        "        Result = ""E""", _
        "    ElseIf optAmount.Value Then", _
        "        AmountText = Trim$(txtAmount.Text)", _
        "        If Len(AmountText) = 0 Then", _
        "            MsgBox ""Enter an amount, or choose a different option."", vbExclamation, ""Amount Required""", _
        "            txtAmount.SetFocus", _
        "            Exit Sub", _
        "        End If", _
        "        Result = ""$""", _
        "    Else", _
        "        Result = ""U""", _
        "    End If", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnStop_Click()", _
        "    Result = ""STOP""", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", _
        "    If CloseMode = vbFormControlMenu Then", _
        "        Result = ""STOP""", _
        "        Cancel = 1", _
        "        Me.Hide", _
        "    End If", _
        "End Sub" _
    ), vbCrLf)
End Sub

' ============================================================
' UF_AddBidder  page 1 of 2  (10 controls)
' Company, Contact, Phone, Email
' ============================================================
Private Sub BuildAddBidder(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_AddBidder", "Add Bidder (1 of 2)", 318, 168)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblCompany", "Company:",      6,  6, 84, 14
    Txt d, "txtCompany",                 96,  6, 216, 16
    Lbl d, "lblContact", "Contact:",      6, 30, 84, 14
    Txt d, "txtContact",                 96, 30, 216, 16
    Lbl d, "lblPhone",   "Phone:",        6, 54, 84, 14
    Txt d, "txtPhone",                   96, 54, 216, 16
    Lbl d, "lblEmail",   "Email:",        6, 78, 84, 14
    Txt d, "txtEmail",                   96, 78, 216, 16
    Btn d, "btnNext",   "Next ->",       6, 114, 144, 24, True,  False
    Btn d, "btnCancel", "Cancel",       162, 114, 144, 24, False, True
    SetCode comp, Join(Array( _
        "Option Explicit", _
        "", _
        "Public Cancelled As Boolean", _
        "", _
        "Public Property Get Company() As String:    Company = Trim$(txtCompany.Text):  End Property", _
        "Public Property Get Contact() As String:    Contact = Trim$(txtContact.Text):  End Property", _
        "Public Property Get Phone() As String:      Phone   = Trim$(txtPhone.Text):    End Property", _
        "Public Property Get Email() As String:      Email   = Trim$(txtEmail.Text):    End Property", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Cancelled = True", _
        "End Sub", _
        "", _
        "Private Sub btnNext_Click()", _
        "    Cancelled = False", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnCancel_Click()", _
        "    Cancelled = True", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", _
        "    If CloseMode = vbFormControlMenu Then", _
        "        Cancelled = True", _
        "        Cancel = 1", _
        "        Me.Hide", _
        "    End If", _
        "End Sub" _
    ), vbCrLf)
End Sub

' ============================================================
' UF_AddBidder2  page 2 of 2  (9 controls)
' Date Received, Notes, Base Bid, Re-enter checkbox
' ============================================================
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
    Btn d, "btnAdd",    "Add Bidder", 6,   114, 144, 24, True,  False
    Btn d, "btnCancel", "Cancel",    162,  114, 144, 24, False, True
    SetCode comp, Join(Array( _
        "Option Explicit", _
        "", _
        "Public Cancelled As Boolean", _
        "", _
        "Public Property Get DateReceived() As String: DateReceived = Trim$(txtDate.Text):    End Property", _
        "Public Property Get Notes() As String:        Notes        = Trim$(txtNotes.Text):   End Property", _
        "Public Property Get BaseBid() As String:      BaseBid      = Trim$(txtBaseBid.Text): End Property", _
        "Public Property Get ReEnterScope() As Boolean: ReEnterScope = chkReEnter.Value:      End Property", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Cancelled = True", _
        "    chkReEnter.Value = False", _
        "End Sub", _
        "", _
        "Private Sub btnAdd_Click()", _
        "    Cancelled = False", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnCancel_Click()", _
        "    Cancelled = True", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", _
        "    If CloseMode = vbFormControlMenu Then", _
        "        Cancelled = True", _
        "        Cancel = 1", _
        "        Me.Hide", _
        "    End If", _
        "End Sub" _
    ), vbCrLf)
End Sub

' ============================================================
' UF_JobInfo  page 1 of 2  (10 controls)
' Project Name, Estimator, Bid Date, Job GSF
' ============================================================
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
    Btn d, "btnNext",   "Next ->",         6, 114, 144, 24, True,  False
    Btn d, "btnCancel", "Cancel",        162, 114, 144, 24, False, True
    SetCode comp, Join(Array( _
        "Option Explicit", _
        "", _
        "Public Cancelled As Boolean", _
        "", _
        "Public Property Get ProjName() As String:  ProjName  = Trim$(txtProjName.Text):  End Property", _
        "Public Property Get Estimator() As String: Estimator = Trim$(txtEstimator.Text): End Property", _
        "Public Property Get BidDate() As String:   BidDate   = Trim$(txtBidDate.Text):   End Property", _
        "Public Property Get JobGSF() As String:    JobGSF    = Trim$(txtJobGSF.Text):    End Property", _
        "", _
        "Public Sub Prefill(ByVal pName As String, ByVal pEst As String, ByVal pDate As String, ByVal pGSF As String)", _
        "    txtProjName.Text  = pName", _
        "    txtEstimator.Text = pEst", _
        "    txtBidDate.Text   = pDate", _
        "    txtJobGSF.Text    = pGSF", _
        "End Sub", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Cancelled = True", _
        "End Sub", _
        "", _
        "Private Sub btnNext_Click()", _
        "    Cancelled = False", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnCancel_Click()", _
        "    Cancelled = True", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", _
        "    If CloseMode = vbFormControlMenu Then", _
        "        Cancelled = True", _
        "        Cancel = 1", _
        "        Me.Hide", _
        "    End If", _
        "End Sub" _
    ), vbCrLf)
End Sub

' ============================================================
' UF_JobInfo2  page 2 of 2  (6 controls)
' Trade SF + overwrite options + Generate/Cancel buttons
' ============================================================
Private Sub BuildJobInfo2(ByVal vbp As Object)
    Dim comp As Object: Set comp = MkForm(vbp, "UF_JobInfo2", "Generate Trade Tabs (2 of 2)", 348, 150)
    Dim d As Object: Set d = comp.Designer
    Lbl d, "lblTradeSF",       "Trade SF  (this run only -- not saved):", 6,  6, 210, 14
    Txt d, "txtTradeSF",                                                 222,  6, 114, 16
    Opt d, "optOverwriteNo",   "Seed only NEW tabs  (default)",           6, 30, 330, 16
    Opt d, "optOverwriteYes",  "Clear + re-seed ALL selected tabs",       6, 54, 330, 16
    Btn d, "btnGenerate", "Generate Tabs", 6,  90, 144, 24, True,  False
    Btn d, "btnCancel",   "Cancel",       162,  90, 144, 24, False, True
    SetCode comp, Join(Array( _
        "Option Explicit", _
        "", _
        "Public Cancelled As Boolean", _
        "Public OverwriteScope As Long", _
        "", _
        "Public Property Get TradeSF() As String: TradeSF = Trim$(txtTradeSF.Text): End Property", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Cancelled = True", _
        "    OverwriteScope = 2", _
        "    optOverwriteNo.Value = True", _
        "End Sub", _
        "", _
        "Private Sub btnGenerate_Click()", _
        "    Cancelled = False", _
        "    OverwriteScope = IIf(optOverwriteYes.Value, 1, 2)", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub btnCancel_Click()", _
        "    Cancelled = True", _
        "    Me.Hide", _
        "End Sub", _
        "", _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", _
        "    If CloseMode = vbFormControlMenu Then", _
        "        Cancelled = True", _
        "        Cancel = 1", _
        "        Me.Hide", _
        "    End If", _
        "End Sub" _
    ), vbCrLf)
End Sub
