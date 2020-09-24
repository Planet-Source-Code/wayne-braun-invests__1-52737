VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddEdit 
   BackColor       =   &H00E6F2FF&
   Caption         =   "Investment Management"
   ClientHeight    =   5265
   ClientLeft      =   195
   ClientTop       =   2790
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   11580
   Begin VB.TextBox txtCostBasis 
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   5
      TabIndex        =   20
      Top             =   3825
      Width           =   780
   End
   Begin VB.TextBox txtID 
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   6795
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3735
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ListBox lstTaxFree 
      DataSource      =   "datInvestments"
      Height          =   450
      ItemData        =   "frmAddEditDelete.frx":0000
      Left            =   3015
      List            =   "frmAddEditDelete.frx":000A
      TabIndex        =   8
      Top             =   1710
      Width           =   465
   End
   Begin MSAdodcLib.Adodc datSearch 
      Height          =   375
      Left            =   8595
      Top             =   90
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc datInvestments 
      Height          =   375
      Left            =   5760
      Top             =   45
      Visible         =   0   'False
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel Changes and Exit"
      Height          =   600
      Left            =   8730
      Picture         =   "frmAddEditDelete.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2835
      Width           =   1950
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save and Exit"
      Height          =   600
      Left            =   6525
      Picture         =   "frmAddEditDelete.frx":0161
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2835
      Width           =   1950
   End
   Begin VB.TextBox txtMemo 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   960
      Left            =   990
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   4275
      Width           =   10500
   End
   Begin VB.TextBox txtPurchasePrice 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      TabIndex        =   19
      ToolTipText     =   "Enter Price without a $, comma or decimal point"
      Top             =   3510
      Width           =   1185
   End
   Begin VB.TextBox txtDocLoc 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3150
      Width           =   1950
   End
   Begin VB.TextBox txtPaysBy 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   12
      TabIndex        =   15
      Top             =   2835
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker DTPickerAcquire 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3015
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   15135487
      CalendarTitleBackColor=   5336445
      CalendarTitleForeColor=   16777215
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   24510467
      CurrentDate     =   38023
   End
   Begin VB.TextBox txtCusip 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   16
      TabIndex        =   11
      Top             =   2205
      Width           =   1680
   End
   Begin MSComCtl2.DTPicker DTPickerMature 
      Bindings        =   "frmAddEditDelete.frx":02AB
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3015
      TabIndex        =   7
      Top             =   1395
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   15135487
      CalendarTitleBackColor=   5336445
      CalendarTitleForeColor=   16777215
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   24510467
      CurrentDate     =   38023
   End
   Begin VB.TextBox txtRate 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1080
      Width           =   1050
   End
   Begin VB.TextBox txtBroker 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   40
      TabIndex        =   3
      Top             =   765
      Width           =   3930
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "datInvestments"
      Height          =   285
      Left            =   3015
      MaxLength       =   40
      TabIndex        =   1
      Top             =   450
      Width           =   3930
   End
   Begin VB.Label lblCostBasis 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Cost Basis (0.000 to 9.999)"
      Height          =   240
      Left            =   1035
      TabIndex        =   28
      Top             =   3870
      Width           =   1950
   End
   Begin VB.Label lblAdjustYr 
      BackColor       =   &H00E6F2FF&
      Caption         =   $"frmAddEditDelete.frx":02C4
      Height          =   870
      Left            =   4905
      TabIndex        =   26
      Top             =   1350
      Width           =   5010
   End
   Begin VB.Label lblRequired 
      BackColor       =   &H00E6F2FF&
      Caption         =   "*  (Required Field)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3015
      TabIndex        =   25
      Top             =   90
      Width           =   2400
   End
   Begin VB.Label lblMemo 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Call Dates and Notes"
      Height          =   555
      Left            =   90
      TabIndex        =   21
      Top             =   4320
      Width           =   825
   End
   Begin VB.Label lblPurcasePrice 
      BackColor       =   &H00E6F2FF&
      Caption         =   "      Purchase Price *                (Digits 0 to 9  or a - only)"
      Height          =   405
      Left            =   1035
      TabIndex        =   18
      Top             =   3420
      Width           =   1950
   End
   Begin VB.Label lblDocLoc 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Location of Documents"
      Height          =   180
      Left            =   1305
      TabIndex        =   16
      Top             =   3195
      Width           =   1680
   End
   Begin VB.Label lblPaysBy 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Method Investment Pays By"
      Height          =   180
      Left            =   900
      TabIndex        =   14
      Top             =   2880
      Width           =   2085
   End
   Begin VB.Label lblAcquire 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Acquire Date  *"
      Height          =   180
      Left            =   1755
      TabIndex        =   12
      Top             =   2565
      Width           =   1230
   End
   Begin VB.Label lblCusip 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Cusip Number"
      Height          =   180
      Left            =   1980
      TabIndex        =   10
      Top             =   2250
      Width           =   1005
   End
   Begin VB.Label lblTaxFree 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Tax Free  ( Click One )  *"
      Height          =   180
      Left            =   1170
      TabIndex        =   9
      Top             =   1800
      Width           =   1860
   End
   Begin VB.Label lblMature 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Mature Date  *"
      Height          =   180
      Left            =   1890
      TabIndex        =   6
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Rate"
      Height          =   180
      Left            =   2565
      TabIndex        =   4
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label lblBroker 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Broker Name && Phone"
      Height          =   180
      Left            =   1350
      TabIndex        =   2
      Top             =   810
      Width           =   1635
   End
   Begin VB.Label lblName 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Investment Name  *"
      Height          =   180
      Left            =   1530
      TabIndex        =   0
      Top             =   495
      Width           =   1500
   End
End
Attribute VB_Name = "frmAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ButtonChoice As String    '---to retain the add, or edit choice
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Form_Load()
   Dim Ptr As String   '---holds chosen record ID as text
   frmAddEdit.Top = 150
   frmAddEdit.Left = 150
   datInvestments.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datInvestments.RecordSource = "Select * from Investment Order By InvestmentName"
   datInvestments.Refresh
   datSearch.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datSearch.RecordSource = "Search"   '---Link datSearch to the Search table
   datSearch.Refresh
   If datInvestments.Recordset.RecordCount > 0 Then '---go to selected investment if there is one
      Ptr = Str(datSearch.Recordset!ID)
      datInvestments.Recordset.Find "InvestmentID = " & Ptr   '---move to the chosen record
   End If
   txtID.DataField = "InvestmentID"
   txtName.DataField = "InvestmentName"
   txtBroker.DataField = "BrokerNamePhone"
   txtRate.DataField = "Rate"
   lstTaxFree.DataField = "TaxFree"
   txtCusip.DataField = "CusipNumber"
   txtPaysBy.DataField = "PaysBy"
   txtDocLoc.DataField = "DocumentsLocation"
   txtPurchasePrice.DataField = "PurchasePrice"
   txtCostBasis.DataField = "CostBasis"
   txtMemo.DataField = "Notes"
   If ButtonChoice = "add" Then
      frmAddEdit.Caption = "       Add an Investment"
      datInvestments.Recordset.AddNew
      DTPickerAcquire.Value = Now
      DTPickerMature.Value = Now
   End If
   If ButtonChoice = "edit" Then
      frmAddEdit.Caption = "        Edit this Investment"
      DTPickerAcquire.Value = datInvestments.Recordset!AcquireDate
      DTPickerMature.Value = datInvestments.Recordset!MatureDate
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdSave_Click()
   Dim NewName As String   '---to move to new investment (or edited one) when done
   If txtName.Text = "" Then      '---enforce required field of InvestmentName
      txtName.SetFocus
      MsgBox "You did not enter an Investment Name.  This is a required field."
      Exit Sub
   End If
   If txtPurchasePrice = "" Then
      txtPurchasePrice.SetFocus
      MsgBox "You did not enter a Purchase Price.  This is a required field."
      Exit Sub
   End If
   If IsDate(DTPickerMature.Value) = False Then
      MsgBox ("You must select a Mature Date.  This is a required field.")
      Exit Sub
   End If
   If IsDate(DTPickerAcquire.Value) = False Then
      MsgBox ("You must select an Acquire Date.  This is a required field.")
      Exit Sub
   End If
   If lstTaxFree.ListIndex = -1 Then  '---if Tax Free not selected, don't proceed
      MsgBox ("You must select the Tax Free status as Yes/No.  This is a required field.")
      Exit Sub
   End If
   datInvestments.Recordset!TaxFree = lstTaxFree.Text
   datInvestments.Recordset!MatureDate = DTPickerMature.Value
   datInvestments.Recordset!AcquireDate = DTPickerAcquire.Value
   NewName = txtName.Text
   frmMain.datSearch.Recordset.Update
   frmMain.datSearch.Recordset.Requery   '---make the above assignment available
   datInvestments.Recordset.Update   '---must now update
   datInvestments.Recordset.Requery  '---the recordset
   frmMain.datInvestments.Refresh    '---and dbgInvestments on frmMain
   frmMain.datSearch.Recordset!ID = datInvestments.Recordset!InvestmentID
   
   If ButtonChoice = "add" Then
      MsgBox " Remember to post initial Principal transaction(s) ! "
   End If
   
   Call frmMain.ResetSummary '---rebuild the investments summary in frmMain.lstSummary
                     '---note that this procedure is Public to be visible from this form
   frmMain.datInvestments.Recordset.MoveFirst   '---prepare for the Find
   frmMain.datInvestments.Recordset.Find "InvestmentName = " & "'" & NewName & "'"
   '---move to the edited or added investment
   Call frmMain.RefreshGrids   '---refresh the principal, income and scheduled pays grids
   Unload Me
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdCancel_Click()
   datInvestments.Recordset.CancelUpdate
   Unload Me
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtCostBasis_KeyPress(KeyAscii As Integer)
   '---allow only backspace and digits 0-9 and decimal
   Select Case KeyAscii
      Case 8   '---backspace key
      Case 46
      Case 48 To 57  '---  0 to 9 keys
      Case Else:
          Beep
          KeyAscii = 0
   End Select
End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtPurchasePrice_KeyPress(KeyAscii As Integer)
    '---allow taking a "-" only if 1st character or digits 0 to 9
    '---else beep and leave cursor where it is.
    Select Case KeyAscii
      Case 8   '---backspace key
      Case 45:   '---   - key
          If Len(txtPurchasePrice) >= 1 Then Beep: KeyAscii = 0
      Case 48 To 57  '---  0 to 9 keys
      Case Else:
          Beep
          KeyAscii = 0
   End Select
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtPurchasePrice_LostFocus()
   Dim Temp As Double   '---add the decimal of number automatically
   If txtPurchasePrice = "" Or txtPurchasePrice = "-" Then
      Temp = 0
   Else
      Temp = txtPurchasePrice.Text
   End If
   If InStr(1, Temp, ".") = 0 Then  '---don't divide again if already divided
      Temp = Temp / 100
   End If
   txtPurchasePrice = Format(Str(Temp), "currency")
End Sub

