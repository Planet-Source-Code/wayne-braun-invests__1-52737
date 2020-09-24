VERSION 5.00
Begin VB.Form frmScheduled 
   BackColor       =   &H008BE2F8&
   Caption         =   "Scheduled Payments"
   ClientHeight    =   4845
   ClientLeft      =   3570
   ClientTop       =   2205
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   3885
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save the Scheduled Payment and Exit"
      Height          =   735
      Left            =   180
      Picture         =   "frmScheduled.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4095
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel Changes and Exit"
      Height          =   735
      Left            =   2115
      Picture         =   "frmScheduled.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4050
      Width           =   1590
   End
   Begin VB.ListBox lstSchdDay 
      Columns         =   4
      Height          =   2055
      IntegralHeight  =   0   'False
      Left            =   2070
      TabIndex        =   2
      Top             =   1395
      Width           =   1590
   End
   Begin VB.ListBox lstSchdMonth 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   135
      TabIndex        =   1
      Top             =   1395
      Width           =   1635
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   2115
      MaxLength       =   14
      TabIndex        =   0
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H008BE2F8&
      Caption         =   "Payment  Amount *          (Digits 0 to 9              or a - only)"
      Height          =   600
      Left            =   540
      TabIndex        =   5
      Top             =   270
      Width           =   1455
   End
   Begin VB.Label lblRequired 
      BackColor       =   &H008BE2F8&
      Caption         =   "*  ( Required Field )"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   990
      TabIndex        =   4
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label lblMonthDay 
      BackColor       =   &H008BE2F8&
      Caption         =   "                   Payment scheduled to pay :                         ( *  Select Month )               ( * Select  Day )"
      Height          =   465
      Left            =   135
      TabIndex        =   3
      Top             =   945
      Width           =   3705
   End
End
Attribute VB_Name = "frmScheduled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Form_Load()
   Dim I As Integer
   frmScheduled.Left = 4000   '---give initial position to form
   frmScheduled.Top = 2900
   lstSchdMonth.AddItem " 1  January"  '---fill the month list
   lstSchdMonth.AddItem " 2  February"
   lstSchdMonth.AddItem " 3  March"
   lstSchdMonth.AddItem " 4  April"
   lstSchdMonth.AddItem " 5  May"
   lstSchdMonth.AddItem " 6  June"
   lstSchdMonth.AddItem " 7  July"
   lstSchdMonth.AddItem " 8  August"
   lstSchdMonth.AddItem " 9  September"
   lstSchdMonth.AddItem "10  October"
   lstSchdMonth.AddItem "11  November"
   lstSchdMonth.AddItem "12  December"
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub lstSchdMonth_Click()
   Dim Days As Integer
   Dim I As Integer
   lstSchdDay.Clear  '---clear any old results
   Select Case lstSchdMonth.ListIndex  '---fill the day list with
   Case 1                            '---appropriate number of days
      Days = 29                      '---for the chosen month
   Case 3, 5, 8, 10
      Days = 30
   Case Else
      Days = 31
   End Select
   For I = 1 To Days
      lstSchdDay.AddItem I
   Next I
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdSave_Click()
   If frmMain.datInvestments.Recordset.RecordCount > 0 Then  '---skip if no investments
      InvestID = frmMain.datInvestments.Recordset!InvestmentID
      If txtAmount = "" Then     '---enforce required field of InvestmentName
         txtAmount.SetFocus
         MsgBox "You did not enter the Payment Amount.  This is a required field."
         Exit Sub
      End If
      If lstSchdMonth.ListIndex = -1 Then
         lstSchdMonth.SetFocus
         MsgBox "You did not click a Payment Month.  This is a required field."
         Exit Sub
      End If
      If lstSchdDay.ListIndex = -1 Then
         lstSchdDay.SetFocus
         MsgBox "You did not click a Payment Day.  This is a required field."
         Exit Sub
      End If
      frmMain.datScheduledPays.Recordset.AddNew  '---Not doing AddNew until committed to
                                 '---the save makes cmdCancel below a trivial matter
      frmMain.datScheduledPays.Recordset!Amount = txtAmount
      frmMain.datScheduledPays.Recordset!SchdMonth = lstSchdMonth.ListIndex + 1
      frmMain.datScheduledPays.Recordset!SchdDay = lstSchdDay.List(lstSchdDay.ListIndex)
      frmMain.datScheduledPays.Recordset!InvestmentID = frmMain.datSearch.Recordset!ID
                                            '---link scheduled pay to selected investment
      frmMain.datScheduledPays.Recordset.Update   '---must now update
      frmMain.datScheduledPays.Recordset.Requery   '---the recordset
      If frmMain.datScheduledPays.Recordset.EOF = True Then
         frmMain.datScheduledPays.Recordset.MoveLast
      End If
      Call frmMain.ResetSummary
      frmMain.datInvestments.Recordset.MoveFirst  '---return to selected investment
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


