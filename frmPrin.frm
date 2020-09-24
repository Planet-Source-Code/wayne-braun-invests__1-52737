VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrin 
   BackColor       =   &H00D2B7A6&
   Caption         =   " Post a Principal Transaction"
   ClientHeight    =   3600
   ClientLeft      =   4260
   ClientTop       =   3045
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3870
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1665
      MaxLength       =   14
      TabIndex        =   0
      Top             =   1215
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "   Cancel Changes and Exit"
      Height          =   735
      Left            =   1845
      Picture         =   "frmPrin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2835
      Width           =   1590
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Post the Principal Transaction and Exit"
      Height          =   735
      Left            =   90
      Picture         =   "frmPrin.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2835
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DTPickerPrin 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1665
      TabIndex        =   1
      Top             =   1800
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   13809574
      CalendarTitleBackColor=   12582912
      CalendarTitleForeColor=   16777215
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   24444931
      CurrentDate     =   38038
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D2B7A6&
      Caption         =   "*  (Required Field)"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1125
      TabIndex        =   6
      Top             =   495
      Width           =   1545
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00D2B7A6&
      Caption         =   "Date *"
      Height          =   240
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   510
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00D2B7A6&
      Caption         =   "       Amount *          (Digits 0 to 9          or a - only)"
      Height          =   690
      Left            =   360
      TabIndex        =   4
      Top             =   1035
      Width           =   1275
   End
End
Attribute VB_Name = "frmPrin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Form_Load()
   frmPrin.Top = 3915   '---position the form to hide the Income form
   frmPrin.Left = 3950
         DTPickerPrin.Value = Now
   frmPrin.Caption = "Post a Principal Transaction"
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdSave_Click()
   If frmMain.datInvestments.Recordset.RecordCount > 0 Then  '---skip if no investments
      InvestID = frmMain.datInvestments.Recordset!InvestmentID
      If txtAmount = "" Then     '---enforce required field of InvestmentName
         txtAmount.SetFocus
         MsgBox "You did not enter the Principal Amount.  This is a required field."
         Exit Sub
      End If
      If IsDate(DTPickerPrin.Value) = False Then
         DTPickerPrin.SetFocus
         MsgBox "You did not enter a Principal Date.  This is a required field."
         Exit Sub
      End If
      frmMain.datPrincipal.Recordset.AddNew  '---Not doing AddNew until committed to
                                 '---the save makes cmdCancel below a trivial matter
      If DTPickerPrin.Year Mod (2) = 0 Then   '---put Head & Tail markings on even years
         frmMain.datPrincipal.Recordset!Head = "*"   '---to visually separate them
         frmMain.datPrincipal.Recordset!Tail = "*"  '---from the odd years.
      End If
      frmMain.datPrincipal.Recordset!Amount = txtAmount
      frmMain.datPrincipal.Recordset!PrinDate = DTPickerPrin.Value
      frmMain.datPrincipal.Recordset!InvestmentID = frmMain.datSearch.Recordset!ID '---link transaction to selected investment
      frmMain.datPrincipal.Recordset.Update   '---must now update
      frmMain.datPrincipal.Recordset.Requery   '---the recordset
      If frmMain.datPrincipal.Recordset.EOF = True Then
         frmMain.datPrincipal.Recordset.MoveLast
      End If
      Call frmMain.ReBuildPrincipal
      Call frmMain.ResetSummary
      frmMain.datInvestments.Recordset.MoveFirst  '---prepare for the Find
      frmMain.datInvestments.Recordset.Find "InvestmentID = " & InvestID
      Call frmMain.RefreshGrids   '---refresh the principal, income and scheduled pays grids
      Unload Me
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdCancel_Click()
   Unload Me
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    '---allow taking a "-" only if 1st character or digits 0 to 9
    '---else beep and leave cursor where it is.
    Select Case KeyAscii
      Case 8   '---backspace key
      Case 45:   '---   - key
          If Len(txtAmount) >= 1 Then Beep: KeyAscii = 0
      Case 48 To 57  '---  0 to 9 keys
      Case Else:
          Beep
          KeyAscii = 0
   End Select
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtAmount_LostFocus()
   Dim Temp As Double   '---add the decimal of number automatically
   If txtAmount = "" Or txtAmount = "-" Then
      Temp = 0
   Else
      Temp = txtAmount.Text
   End If
   If InStr(1, Temp, ".") = 0 Then  '---don't divide again if already divided
      Temp = Temp / 100
   End If
   txtAmount = Format(Str(Temp), "currency")
End Sub
