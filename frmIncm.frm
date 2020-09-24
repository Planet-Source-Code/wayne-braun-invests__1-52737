VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIncm 
   BackColor       =   &H00ADC4AA&
   Caption         =   "Post an Income Transaction"
   ClientHeight    =   3600
   ClientLeft      =   105
   ClientTop       =   4260
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3870
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel Changes and Exit"
      Height          =   735
      Left            =   1845
      Picture         =   "frmIncm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2835
      Width           =   1590
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Post the Income Transaction and Exit"
      Height          =   735
      Left            =   135
      Picture         =   "frmIncm.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2835
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DTPickerIncm 
      Height          =   330
      Left            =   1710
      TabIndex        =   3
      Top             =   1575
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   11388074
      CalendarTitleBackColor=   5666903
      CalendarTitleForeColor=   16777215
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   24510467
      CurrentDate     =   38036
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1710
      MaxLength       =   14
      TabIndex        =   2
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label lblRequired 
      BackColor       =   &H00ADC4AA&
      Caption         =   "*  (Required Field)"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1125
      TabIndex        =   6
      Top             =   405
      Width           =   1455
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00ADC4AA&
      Caption         =   "Date *"
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1620
      Width           =   510
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00ADC4AA&
      Caption         =   "       Amount *          (Digits 0 to 9          or a - only)"
      Height          =   600
      Left            =   405
      TabIndex        =   0
      Top             =   855
      Width           =   1230
   End
End
Attribute VB_Name = "frmIncm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Form_Load()
   frmIncm.Top = 3915   '---position the form to hide the principal form
   frmIncm.Left = 50
         DTPickerIncm.Value = Now
   frmIncm.Caption = "Post an Income Transaction"
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdSave_Click()
   If frmMain.datInvestments.Recordset.RecordCount > 0 Then  '---skip if no investments
      InvestID = frmMain.datInvestments.Recordset!InvestmentID
      If txtAmount = "" Then     '---enforce required field of InvestmentName
         txtAmount.SetFocus
         MsgBox "You did not enter the Income Amount.  This is a required field."
         Exit Sub
      End If
      If IsDate(DTPickerIncm.Value) = False Then
         DTPickerIncm.SetFocus
         MsgBox "You did not enter an Income Date.  This is a required field."
         Exit Sub
      End If
      frmMain.datIncome.Recordset.AddNew  '---Not doing AddNew until committed to
                                 '---the save makes cmdCancel below a trivial matter
      If DTPickerIncm.Year Mod (2) = 0 Then   '---put Head & Tail markings on even years
         frmMain.datIncome.Recordset!Head = "*"   '---to visually separate them
         frmMain.datIncome.Recordset!Tail = "*"  '---from the odd years.
      End If
      frmMain.datIncome.Recordset!Amount = txtAmount
      frmMain.datIncome.Recordset!IncmDate = DTPickerIncm.Value
      frmMain.datIncome.Recordset!InvestmentID = frmMain.datSearch.Recordset!ID '---link transaction to selected investment
      frmMain.datIncome.Recordset.Update   '---must now update
      frmMain.datIncome.Recordset.Requery   '---the recordset
      If frmMain.datIncome.Recordset.EOF = True Then
         frmMain.datIncome.Recordset.MoveLast
      End If
      Call frmMain.ReBuildIncome
      'Call frmMain.ResetSummary
      'frmMain.datInvestments.Recordset.MoveFirst  '---prepare for the Find
      'frmMain.datInvestments.Recordset.Find "InvestmentID = " & InvestID
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
