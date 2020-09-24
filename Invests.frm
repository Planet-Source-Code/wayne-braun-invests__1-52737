VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   -330
   ClientTop       =   855
   ClientWidth     =   11940
   Icon            =   "Invests.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11940
   Begin VB.ListBox lstSummary 
      BackColor       =   &H00CBB674&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7830
      TabIndex        =   21
      Top             =   7380
      Width           =   4065
   End
   Begin VB.CommandButton cmdPrintYrEndReports 
      BackColor       =   &H009C79B9&
      Caption         =   "Print Year-End Report"
      Height          =   500
      Left            =   5580
      Picture         =   "Invests.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7695
      Width           =   1680
   End
   Begin VB.CommandButton cmdPrintAllInvestments 
      BackColor       =   &H009C79B9&
      Caption         =   "Print All Investments"
      Height          =   500
      Left            =   3960
      Picture         =   "Invests.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7695
      Width           =   1545
   End
   Begin VB.CommandButton cmdPrintInvestment 
      BackColor       =   &H009C79B9&
      Caption         =   "Print Selected Investment"
      Height          =   500
      Left            =   1935
      Picture         =   "Invests.frx":1956
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7695
      Width           =   1950
   End
   Begin VB.CommandButton cmdPrintIncomeSchedule 
      BackColor       =   &H009C79B9&
      Caption         =   "Print Income Schedule"
      Height          =   500
      Left            =   45
      Picture         =   "Invests.frx":1EE0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7695
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeleteSelectedPayment 
      BackColor       =   &H008BE2F8&
      Caption         =   "Delete Selected Payment"
      Height          =   645
      Left            =   9270
      Picture         =   "Invests.frx":246A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6750
      Width           =   1290
   End
   Begin VB.CommandButton cmdAddSchdPay 
      BackColor       =   &H008BE2F8&
      Caption         =   "Add Scheduled Payment"
      Height          =   645
      Left            =   7965
      Picture         =   "Invests.frx":25B4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6750
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid dbgInvestments 
      CausesValidation=   0   'False
      Height          =   3120
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5503
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   15135487
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   12
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "InvestmentID"
         Caption         =   "#"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "InvestmentName"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "BrokerNamePhone"
         Caption         =   "Broker    "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CurrentPrincipal"
         Caption         =   "Current Principal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Rate"
         Caption         =   "Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "MatureDate"
         Caption         =   "Matures"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "TaxFree"
         Caption         =   "Tax Free?"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "AcquireDate"
         Caption         =   "Acquired"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "CostBasis"
         Caption         =   "Cost Basis"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "PurchasePrice"
         Caption         =   "Purchase Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "PaysBy"
         Caption         =   "Pays By"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "DocumentsLocation"
         Caption         =   "Documents Location"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "CusipNumber"
         Caption         =   "Cusip #"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   44.787
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   3360.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDeleteIncomeTransaction 
      BackColor       =   &H00ADC4AA&
      Caption         =   "Delete Selected Income Transaction"
      Height          =   650
      Left            =   5985
      MouseIcon       =   "Invests.frx":2B3E
      Picture         =   "Invests.frx":30C8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6975
      Width           =   1770
   End
   Begin VB.CommandButton cmdPostIncome 
      BackColor       =   &H00ADC4AA&
      Caption         =   "Post Income To Selected Investment"
      Height          =   650
      Left            =   4095
      Picture         =   "Invests.frx":3212
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6975
      Width           =   1770
   End
   Begin VB.CommandButton cmdDeletePrincipalTransaction 
      BackColor       =   &H00D2B7A6&
      Caption         =   "Delete Selected Principal Transaction"
      Height          =   650
      Left            =   1980
      Picture         =   "Invests.frx":379C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6975
      Width           =   1770
   End
   Begin VB.CommandButton cmdPostPrincipal 
      BackColor       =   &H00D2B7A6&
      Caption         =   "Post Principal To Selected Investment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   90
      Picture         =   "Invests.frx":38E6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6975
      Width           =   1770
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00243DE8&
      Caption         =   "Exit"
      Height          =   525
      Left            =   10620
      Picture         =   "Invests.frx":3E70
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6795
      Width           =   1275
   End
   Begin VB.CommandButton cmdDeleteSelectedInvestment 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Delete Selected Investment"
      Height          =   640
      Left            =   10620
      Picture         =   "Invests.frx":3FBA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6075
      Width           =   1290
   End
   Begin VB.CommandButton cmdEditInvestment 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Edit Selected Investment"
      Height          =   640
      Left            =   10620
      Picture         =   "Invests.frx":4104
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5355
      Width           =   1290
   End
   Begin VB.CommandButton cmdAddInvestment 
      BackColor       =   &H00E6F2FF&
      Caption         =   "Add a New Investment"
      Height          =   640
      Left            =   10620
      Picture         =   "Invests.frx":468E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4635
      Width           =   1290
   End
   Begin VB.CommandButton cmdShowSummary 
      BackColor       =   &H00CBB674&
      Caption         =   "Show Income by Month"
      Height          =   640
      Left            =   10620
      Picture         =   "Invests.frx":4C18
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3915
      Width           =   1290
   End
   Begin VB.TextBox txtCallsAndNotes 
      BackColor       =   &H00E6F2FF&
      Height          =   735
      Left            =   945
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3150
      Width           =   10500
   End
   Begin MSDataGridLib.DataGrid dbgScheduledPays 
      Height          =   2805
      Left            =   8010
      TabIndex        =   0
      Top             =   3915
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4948
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   9167608
      ForeColor       =   -2147483625
      HeadLines       =   1
      RowHeight       =   12
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Scheduled Payments"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "SchdMonth"
         Caption         =   "Month"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SchdDay"
         Caption         =   "Day"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Amount"
         Caption         =   "Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datScheduledPays 
      Height          =   330
      Left            =   6525
      Top             =   0
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "datScheduledPays"
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
   Begin MSAdodcLib.Adodc datPrincipal 
      Height          =   330
      Left            =   4365
      Top             =   -45
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "datPrincipal"
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
   Begin MSAdodcLib.Adodc datIncome 
      Height          =   330
      Left            =   2295
      Top             =   -45
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "datIncome"
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
   Begin MSDataGridLib.DataGrid dbgPrincipal 
      Height          =   3030
      Left            =   0
      TabIndex        =   1
      Top             =   3915
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   5345
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   13678502
      HeadLines       =   1
      RowHeight       =   12
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Principal Transactions"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Head"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PrinDate"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "M/dd/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Amount"
         Caption         =   "Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Balance"
         Caption         =   "Balance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tail"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   134.929
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   134.929
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dbgIncome 
      Height          =   3030
      Left            =   4005
      TabIndex        =   14
      Top             =   3915
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   5345
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   11388074
      HeadLines       =   1
      RowHeight       =   12
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Income Transactions"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Head"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "IncmDate"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Amount"
         Caption         =   "Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Balance"
         Caption         =   "Balance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tail"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   134.929
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   134.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datSearch 
      Height          =   330
      Left            =   9135
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
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
      Caption         =   "datSearch"
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
      Height          =   330
      Left            =   135
      Top             =   0
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "datInvestments"
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
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Calls and Other Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   3
      Top             =   3150
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---Invests is used to keep books on investments like municipal bonds (mainly tax free).
'---It may give help to learning to use events, the DataGrid control, ADO, data-bound
'---controls, ADO record add-edit-delete, the DTPicker control for calendar work,
'---text box, list box, multiple forms, a module to give project-wide scope to variables,
'---log-in with encrypted password, the Keypress event to control user input, SQL,
'---database tables, EOF and BOF, variable scope, modality with forms, basic sorted
'---financial transactions, form load event, automatic processing of user dollar amount,
'---allowing inputs of just digits such as 238794 to yield $2387.94 so that decimal points
'---need not be typed, presenting a list of days for a given month (30 for June, 31 for May)
'---This is not a tutorial program, but instead hopefully gives helpful examples and
'---enough commenting that the code can be understood.
'---Invests.mdb database tables are used here.

'---Color is used on the various forms to help avoid user errors.  The user will soon
'---associate working with income transactions with the color green and thus avoid
'---doing them as principal transactions which are blue.

'---Note that on frmMain the data controls are behind dbgInvestments, and you may need
'---to look at the preset properties of controls in order to understand them.

'---Because the DTPicker control cannot bind to a database where the field is null,
'---(especially a problem with a database with no records yet), whenever a DTPicker is
'---used in this program it's value is stored in the database by means of code assignment
'---instead of attempting to use DTPicker as a data-bound control.

'---Some controls need to be added to the project toolbox by selecting project from the menu,
'---then clicking components to get the selection window.  Then check select the following:
'---   Microsoft ADO Data Control 6.0 (OLEDB)
'---   Microsoft Data Grid Control 6.0 (OLEDB)
'---   Microsoft Windows Common Controls-2 6.0(SP3)

Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Form_Load()
   BLANKS = Space$(40) '---40 blanks for Formats
   frmMain.Top = 0   '---position the
   frmMain.Left = 0   '---form
   dbgInvestments.Columns(1).Caption = "Investments For " & OwnerName  '---caption the grid
   '---Also now put Owner into the form caption and instruct for grid sorting
   frmMain.Caption = "Investments For " & OwnerName & "          " & _
      "( Click any column header in the beige investment grid to sort by that column. )"
   datInvestments.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datInvestments.RecordSource = "Select * from Investment Order By InvestmentName"
   datInvestments.Refresh
   Set dbgInvestments.DataSource = datInvestments  '---fill the investments grid
   dbgInvestments.MarqueeStyle = 3   '---causes whole row to highlight when a cell clicked
   If datInvestments.Recordset.RecordCount > 0 Then
      InvestID = datInvestments.Recordset!InvestmentID '---point to record to return to
   End If
   
   '---Only one record of table Search is used and it has only the fields of
   '---ID,Month & Year which will hold the record number of the selected investment or
   '---Month used to supply the SQL for extracting principal and income transactions
   '---belonging to the selected investment record number or month, or Year to filter to.
   '---(Because SQL must use a database variable and not an ordinary variable.)
   datSearch.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datSearch.RecordSource = "Search"   '---Link datSearch to the Search table
   datSearch.Refresh
   If datSearch.Recordset.RecordCount = 0 Then
      datSearch.Recordset.AddNew
   End If
   If datInvestments.Recordset.RecordCount = 0 Then
      datSearch.Recordset!ID = 0
   Else
      datSearch.Recordset!ID = datInvestments.Recordset!InvestmentID
   End If
   
   datPrincipal.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datPrincipal.CommandType = adCmdText
   dbgPrincipal.MarqueeStyle = 3   '---causes whole row to highlight when a cell clicked
   
   datIncome.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datIncome.CommandType = adCmdText
   dbgIncome.MarqueeStyle = 3   '---causes whole row to highlight when a cell clicked
   
   datScheduledPays.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   datScheduledPays.CommandType = adCmdText
   dbgScheduledPays.MarqueeStyle = 3   '---causes whole row to highlight when a cell clicked
   
   Set txtCallsAndNotes.DataSource = datInvestments   '---link in the Calls and
   txtCallsAndNotes.DataField = "Notes"         '---Notes list
   Call ResetSummary '---rebuild the investments summary lstSummary
   If datInvestments.Recordset.RecordCount > 0 Then   '---do if investment(s) exist
      datInvestments.Recordset.MoveFirst
      datSearch.Recordset!ID = InvestID
      datSearch.Recordset.Update
      datSearch.Recordset.Requery   '---make the above assignment available
      Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
   Else   '---do if no investment exists
      datPrincipal.RecordSource = "SELECT * from Principal"
      datIncome.RecordSource = "SELECT * from Income"
      datScheduledPays.RecordSource = "SELECT * from ScheduledPays"
   End If
   
   Set dbgPrincipal.DataSource = datPrincipal   '---Fill the grid with Principal transactions
                                    '---that belong to the selected investment
   Set dbgIncome.DataSource = datIncome   '---fill the grid with income transactions
                                '---that belong to the selected investment
   Set dbgScheduledPays.DataSource = datScheduledPays   '---fill grid with scheduled pays
                                           '---that belong to selected investment
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdShowSummary_Click()   '---show a summary of the investment portfolio
   If datInvestments.Recordset.RecordCount > 0 Then
      InvestID = datInvestments.Recordset!InvestmentID
      frmSummary.Show vbModal
      Call ResetSummary
      datInvestments.Recordset.MoveFirst  '---return to selected investment
      datInvestments.Recordset.Find "InvestmentID = " & InvestID
      Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
   Else
      MsgBox "There are no investments to summarize. "
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdAddInvestment_Click()
   frmAddEdit.ButtonChoice = "add"
   frmAddEdit.Show vbModal   '---deny user access to other forms while on frmAddEdit
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdEditInvestment_Click()
   If datInvestments.Recordset.RecordCount = 0 Then  '---skip if no investments
      Exit Sub
   End If
   frmAddEdit.ButtonChoice = "edit"
   frmAddEdit.Show vbModal   '---deny user access to other forms while on frmAddEdit
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdDeleteSelectedInvestment_Click()
   Dim I As Integer
   Dim Response As String '---longer response is used here to give more accident protection
   If datInvestments.Recordset.RecordCount > 0 Then '---don't try to delete if no investments
      Response = InputBox("WARNING ! ! ! !    Are you sure you want to PERMANENTLY DELETE" _
      & " the above selected investment and all of its transactions?" _
      & "  To delete it, DOUBLE CHECK the SELECTED INVESTMENT, " _
      & "type the word    delete    (all lower case) in the response box, and click OK, " _
      & "otherwise, click cancel.", "Delete Warning", , 2000, 4000)
      If Response = "delete" Then
      '---delete principal and income transactions and scheduled pays first
         If datPrincipal.Recordset.RecordCount > 0 Then
            datPrincipal.Recordset.MoveFirst
            For I = 1 To datPrincipal.Recordset.RecordCount
               datPrincipal.Recordset.Delete  '---delete the principal transaction
               datPrincipal.Recordset.MoveNext  '---move to record after deleted record
            Next I
            datPrincipal.Recordset.Requery
         End If
      
         If datIncome.Recordset.RecordCount > 0 Then
            datIncome.Recordset.MoveFirst
            For I = 1 To datIncome.Recordset.RecordCount
               datIncome.Recordset.Delete
               datIncome.Recordset.MoveNext
            Next I
            datIncome.Recordset.Requery
         End If
      
         If datScheduledPays.Recordset.RecordCount > 0 Then
            datScheduledPays.Recordset.MoveFirst
            For I = 1 To datScheduledPays.Recordset.RecordCount
               datScheduledPays.Recordset.Delete
               datScheduledPays.Recordset.MoveNext
            Next I
            datScheduledPays.Recordset.Requery
         End If
      
         datSearch.Recordset!ID = 0   '---to indicate that there is no investment
         datSearch.Recordset.Update
         datSearch.Recordset.Requery   '---make the above assignment available
         If datInvestments.Recordset.RecordCount > 0 Then  '---delete the investment
            datInvestments.Recordset.Delete
            datInvestments.Recordset.MoveNext
            If datInvestments.Recordset.EOF = True And _
            datInvestments.Recordset.RecordCount > 0 Then
               datInvestments.Recordset.MoveLast
            End If
            datInvestments.Recordset.Requery
            Call ResetSummary '---rebuild the investments summary in lstSummary
            If datInvestments.Recordset.RecordCount > 0 Then
               InvestID = datInvestments.Recordset!InvestmentID '---point to record to return to
               datInvestments.Recordset.MoveFirst
               datInvestments.Recordset.Find "InvestmentID = " & InvestID '---next investment
               Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
            End If
            MsgBox ("The selected investment was deleted") '---if one exists
         End If
      End If
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPostPrincipal_Click()
   If datInvestments.Recordset.RecordCount > 0 Then
      frmPrin.Show vbModal
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdDeletePrincipalTransaction_Click()
   Dim PrinPtr As String  '---to point to next tran. at exit, or last one if no next one
   Dim Response As String
   '---skip if no investments or principal transactions
   If datPrincipal.Recordset.RecordCount > 0 And datInvestments.Recordset.RecordCount > 0 Then
      InvestID = datInvestments.Recordset!InvestmentID
      Response = InputBox("Are you sure you want to PERMANENTLY DELETE" & Chr$(13) _
      & "the selected Principal transaction below?" & Chr$(13) & Chr$(13) _
      & "To delete it, DOUBLE CHECK the SELECTED TRANSACTION," & Chr$(13) _
      & "type a    *    in the response box, and click OK." & Chr$(13) & Chr$(13) _
      & "Otherwise, click Cancel.                                     " _
      , "Delete Warning", , 50, 1650)
      If Response = "*" Then
         datPrincipal.Recordset.Delete   '---delete the current record
         datPrincipal.Recordset.MoveNext '---must now update the recordset
         If datPrincipal.Recordset.EOF = True And datPrincipal.Recordset.RecordCount > 0 Then
            datPrincipal.Recordset.MoveLast
         End If
         If datPrincipal.Recordset.RecordCount > 0 Then
            PrinPtr = datPrincipal.Recordset!PrinDate  '---save transaction date for below
         End If
         datPrincipal.Recordset.Requery
         Call ReBuildPrincipal
         Call ResetSummary
         datInvestments.Recordset.MoveFirst  '---return to selected investment
         datInvestments.Recordset.Find "InvestmentID = " & InvestID
         Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
         If datPrincipal.Recordset.RecordCount > 0 Then
            datPrincipal.Recordset.MoveFirst  '---prepare for the Find
            datPrincipal.Recordset.Find "PrinDate = " & "'" & PrinPtr _
               & "'" '---point to transaction as above
         End If
      End If
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPostIncome_Click()
   If datInvestments.Recordset.RecordCount > 0 Then
      frmIncm.Show vbModal
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdDeleteIncomeTransaction_Click()
   Dim IncmPtr As String  '---to point to next tran. at exit, or last one if no next one
   Dim Response As String
   '---skip if no investments or income transactions
   If datIncome.Recordset.RecordCount > 0 And datInvestments.Recordset.RecordCount > 0 Then
      InvestID = datInvestments.Recordset!InvestmentID
      Response = InputBox("Are you sure you want to PERMANENTLY" & Chr$(13) _
      & "DELETE the selected Income transaction below?" & Chr$(13) & Chr$(13) _
      & "To delete it, DOUBLE CHECK the SELECTED TRANSACTION," & Chr$(13) _
      & "type a    *    in the response box, and click OK." & Chr$(13) & Chr$(13) _
      & "Otherwise, click Cancel.                                     " _
      , "Delete Warning", , 3150, 1650)
      If Response = "*" Then
         datIncome.Recordset.Delete   '---delete the current record
         datIncome.Recordset.MoveNext '---must now update the recordset
         If datIncome.Recordset.EOF = True And datIncome.Recordset.RecordCount > 0 Then
            datIncome.Recordset.MoveLast
         End If
         If datIncome.Recordset.RecordCount > 0 Then
            IncmPtr = datIncome.Recordset!IncmDate  '---save transaction date for below
         End If
         datIncome.Recordset.Requery
         Call ReBuildIncome
         datInvestments.Recordset.MoveFirst  '---return to selected investment
         datInvestments.Recordset.Find "InvestmentID = " & InvestID
         Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
         If datIncome.Recordset.RecordCount > 0 Then
            datIncome.Recordset.MoveFirst   '---prepare for the Find
            datIncome.Recordset.Find "IncmDate = " & "'" & IncmPtr _
               & "'" '---point to transaction as above
         End If
      End If
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdAddSchdPay_Click()
   If datInvestments.Recordset.RecordCount > 0 Then
      frmScheduled.Show vbModal
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdDeleteSelectedPayment_Click()
   Dim Response As String
   If datScheduledPays.Recordset.RecordCount > 0 And _
   datInvestments.Recordset.RecordCount > 0 Then '---skip if no investments or payments
      InvestID = datInvestments.Recordset!InvestmentID
      Response = InputBox("To delete the selected Scheduled Payment" & Chr$(13) _
      & "below, type    *    in the response box, then click OK," & Chr$(13) & Chr$(13) _
      & "Otherwise, click Cancel.                                " _
      , "Delete Warning", , 6500, 2100)
      If Response = "*" Then
         datScheduledPays.Recordset.Delete   '---delete the current record
         datScheduledPays.Recordset.MoveNext '---must now update the recordset
         If datScheduledPays.Recordset.EOF = True And _
         datScheduledPays.Recordset.RecordCount > 0 Then
            datScheduledPays.Recordset.MoveLast
         End If
         datScheduledPays.Recordset.Requery
         Call ResetSummary
         datInvestments.Recordset.MoveFirst  '---return to selected investment
         datInvestments.Recordset.Find "InvestmentID = " & InvestID
         Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
      End If
   End If
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrintIncomeSchedule_Click()
   Dim J As Integer   '---to traverse the scheduled payments
   Dim Response As String
   If datInvestments.Recordset.RecordCount = 0 Then
      MsgBox ("There are no investments to print for.")
      Exit Sub
   End If
   InvestID = datInvestments.Recordset!InvestmentID  '---to return to selected investment
   datScheduledPays.RecordSource = "SELECT SchdMonth, SchdDay, Amount " & _
        "FROM ScheduledPays " & _
        "ORDER BY ScheduledPays.SchdMonth"
   datScheduledPays.Refresh
   Response = InputBox("To print the Income Schedule, check that the printer " _
   & " is turned on and paper is ready, then:" & Chr$(13) & Chr$(13) _
   & " To print small ( for billfold size ), type S in the " & Chr$(13) _
   & " response box and click OK. " & Chr$(13) & Chr$(13) _
   & " To print full size, type F in the response box and click OK. " & Chr$(13) & Chr$(13) _
   & "otherwise, click cancel.", "Print Income Schedule", , 4000, 3000)
   If Response = "" Then  '---exit for cancel clicked
      Exit Sub
   End If
   Select Case Response
      Case "s", "S"
         Printer.FontSize = 6
      Case "f", "F"
         Printer.FontSize = 12
      Case Else
         MsgBox ("Invalid choice, printing terminated.")
         Exit Sub
   End Select
   Response = InputBox("Type the year this report is for in the response box.")
   If Response = "" Then  '---exit for cancel clicked
      Exit Sub
   End If
   Printer.Font = "courier new"
   Printer.FontBold = True
   Printer.Print "       " & OwnerName & " for the year " & Response  '---print owner & year report is for
   Printer.Print "       Printed on " & Format(Now, "M/dd/yyyy")
   Printer.Print
   For J = 1 To datScheduledPays.Recordset.RecordCount  '---print all scheduled payments
      Printer.Print Right$("   " & datScheduledPays.Recordset!SchdMonth & "/" & _
        datScheduledPays.Recordset!SchdDay, 5);
      Printer.Print Left$(" " & datInvestments.Recordset!InvestmentName & " ", 41);
      Printer.Print Right$("           " & Format(datScheduledPays.Recordset!Amount, _
         "$##,###,###.00"), 12)
      If J < datScheduledPays.Recordset.RecordCount Then
         datScheduledPays.Recordset.MoveNext  '---to next scheduled pay for investment
      End If
   Next J
   Printer.Print lstSummary.List(0)  '---print the totals list
   Printer.Print lstSummary.List(1)
   Printer.Print lstSummary.List(2)
   Printer.FontBold = False    '---reset the printer
   Printer.FontSize = 10
   Printer.EndDoc
   datInvestments.Recordset.MoveFirst  '---return to selected investment
   datInvestments.Recordset.Find "InvestmentID = " & InvestID
   Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrintInvestment_Click()
   Dim Yr As Integer  '---the year to get the total income and principal for
   Mark = "|"   '---this is used to divide printed columns
   If datInvestments.Recordset.RecordCount = 0 Then
      MsgBox ("There are no investments to print for.")
      Exit Sub
   Else
      InvestID = datInvestments.Recordset!InvestmentID  '---to return to selected investment
      PrintYear = InputBox("To print the Investment, check that the printer " _
   & "is turned on and paper is ready, " & Chr$(13) & Chr$(13) _
   & "Then enter the year you want Principal and Income" & Chr$(13) _
   & "totals for (must be 2000-2200)" & Chr$(13) & Chr$(13) _
   & "otherwise, click cancel.", "Printer Setup and Input Year", , 4000, 3000)
   End If
   If PrintYear = "" Then   '---handle the cancel
      Exit Sub
   End If
   Yr = Val(PrintYear)
   If Yr < 2000 Or Yr > 2200 Then '---exit if cancel or bad year input
      MsgBox ("Improper year was entered and printing was canceled.")
      Exit Sub
   End If
   Call FillTransGetTotals(Yr)  '---fill print buffers and get totals for Yr
   Call PrintCurrentInvestment
   datInvestments.Recordset.MoveFirst  '---return to selected investment
   datInvestments.Recordset.Find "InvestmentID = " & InvestID
   Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrintAllInvestments_Click()
   Dim I As Integer  '---to traverse investments
   Dim Yr As Integer  '---the year to get the total income and principal for
   Mark = "|"   '---this is used to divide printed columns
   If datInvestments.Recordset.RecordCount = 0 Then
      MsgBox ("There are no investments to print for.")
      Exit Sub
   Else
      InvestID = datInvestments.Recordset!InvestmentID  '---to return to selected investment
      PrintYear = InputBox("To print the Investments, check that the printer " _
   & "is turned on and paper is ready, " & Chr$(13) & Chr$(13) _
   & "Then enter the year you want Principal and Income" & Chr$(13) _
   & "totals for (must be 2000-2200)" & Chr$(13) & Chr$(13) _
   & "otherwise, click cancel.", "Printer Setup and Input Year", , 4000, 3000)
   End If
   If PrintYear = "" Then   '---handle the cancel
      Exit Sub
   End If
   Yr = Val(PrintYear)
   If Yr < 2000 Or Yr > 2200 Then '---exit if cancel or bad year input
      MsgBox ("Improper year was entered and printing was canceled.")
      Exit Sub
   End If
   datInvestments.Recordset.MoveFirst
   For I = 1 To datInvestments.Recordset.RecordCount
      Call FillTransGetTotals(Yr)  '---fill print buffers and get totals for Yr
      Call PrintCurrentInvestment
      datInvestments.Recordset.MoveNext
   Next I
   datInvestments.Recordset.MoveFirst  '---return to selected investment
   datInvestments.Recordset.Find "InvestmentID = " & InvestID
   Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrintYrEndReports_Click()
   Dim I As Integer  '---to traverse investments
   Dim J As Integer  '---to traverse principal transactions
   Dim Yr As Integer  '---the year to get the total income and principal for
   Dim PrinNow As Currency
   Dim TaxFreeYr As Currency
   Mark = "|"   '---this is used to divide printed columns
   Dim Indent As String
   Indent = Space$(5)
   If datInvestments.Recordset.RecordCount = 0 Then
      MsgBox ("There are no investments to print for.")
      Exit Sub
   Else
      InvestID = datInvestments.Recordset!InvestmentID  '---to return to selected investment
      PrintYear = InputBox("To print the Investment Year-End Report, " _
   & "check that the printer is turned on and paper is ready, " & Chr$(13) & Chr$(13) _
   & "Then enter the year you want the report to apply to( must be 2000-2200)" & Chr$(13) _
   & Chr$(13) _
   & "otherwise, click cancel.", "Printer Setup and Input Year", , 4000, 3000)
   End If
   If PrintYear = "" Then   '---handle the cancel
      Exit Sub
   End If
   Yr = Val(PrintYear)
   If Yr < 2000 Or Yr > 2200 Then '---exit if cancel or bad year input
      MsgBox ("Improper year was entered and printing was canceled.")
      Exit Sub
   End If
   datInvestments.Recordset.MoveFirst
   Printer.Orientation = 2 '---print in landscape mode
   Printer.Font = "courier new"
   Printer.FontSize = 8.5
   Printer.FontBold = True
   Printer.Print   '---provide a top margin
   Printer.Print
   Printer.Print
   '---now print a report title
   Printer.Print Space$(30); Format(Yr, "####"); " Year End Report for "; _
      OwnerName; "    Printed on: "; Format(Now, "mm/dd/yyyy")
   Printer.Print
   Printer.Print
   '---then a column header
   Printer.Print Space$(28); "Investment"; Space$(12); "Broker"; Space$(6); _
        "Principal Now"; Space$(2); Format(Yr, "####"); " Tax Free"; Space$(3); _
        "Prin.Returned w/Date  Date Acquired  Cost Basis"
   Printer.Print
   For I = 1 To datInvestments.Recordset.RecordCount   '---now traverse the investments
      Call FillTransGetTotals(Yr)  '---get totals for Yr, also buffered transactions
      Printer.Print Indent; Right$(BLANKS & datInvestments.Recordset!InvestmentName & _
         Mark, 40); " ";
      Printer.Print Left$(datInvestments.Recordset!BrokerNamePhone & BLANKS, 13); " "; Mark;
      Printer.Print Right$(BLANKS & Format(datInvestments.Recordset!CurrentPrincipal, _
                        "currency"), 14); Mark;
      PrinNow = PrinNow + datInvestments.Recordset!CurrentPrincipal '---build prin. total
      If datInvestments.Recordset!TaxFree = "Yes" And IncmTotal > 0 Then
         Printer.Print Right$(BLANKS & Format(IncmTotal, "currency"), 14); Mark;
         TaxFreeYr = TaxFreeYr + IncmTotal  '---build tax free for year total
      Else
         Printer.Print Space$(14); Mark;
      End If
      '---now get principal transactions for the year Yr
      datSearch.Recordset.MoveFirst
      datSearch.Recordset!Year = Yr
      datSearch.Recordset.Update
      datSearch.Recordset.Requery   '---make the above assignment available
      datPrincipal.RecordSource = "SELECT PrincipalID, PrinDate, Amount, " & _
       "InvestmentID, Search.Year, Search.ID " & _
       "FROM Principal, Search " & _
       "WHERE Year(PrinDate) = Search.Year " & _
       "AND InvestmentID = Search.ID " & _
       "ORDER BY PrincipalID"
      datPrincipal.Refresh
      If datPrincipal.Recordset.RecordCount > 0 Then
         datPrincipal.Recordset.MoveFirst
         For J = 1 To datPrincipal.Recordset.RecordCount
            If datPrincipal.Recordset!Amount < 0 Then
               Printer.Print Tab(92); _
                    Right$(BLANKS & Format(Abs(datPrincipal.Recordset!Amount), _
                   "currency"), 14);
               Printer.Print Format(datPrincipal.Recordset!PrinDate, " mm/dd/yy");
            End If
            datPrincipal.Recordset.MoveNext
         Next J
      End If
      Printer.Print Tab(117); Mark; Format(datInvestments.Recordset!AcquireDate, _
         " mm/dd/yy "); Mark;
      Printer.Print Format(datInvestments.Recordset!CostBasis, "   0.000")
      Printer.Print  '---double space lines for investments
      datInvestments.Recordset.MoveNext
   Next I
   Printer.Print Indent; String$(131, "_")
   Printer.Print Tab(55); "Totals ";
   Printer.Print Right$(BLANKS & Format(PrinNow, "currency"), 14);
   Printer.Print Right$(BLANKS & Format(TaxFreeYr, "currency"), 15);
   Printer.FontSize = 10   '---return printer settings
   Printer.FontBold = False
   Printer.EndDoc
   Printer.Orientation = 1  '---return to portrait mode
   datInvestments.Recordset.MoveFirst  '---return to selected investment
   datInvestments.Recordset.Find "InvestmentID = " & InvestID
   Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Sub ResetSummary()  '---rebuild the investments summary in lstSummary
   '---whenever investments are deleted or principal transactions are posted or deleted.
   Dim MonthIncome As Currency   '---Builds income from investments for each month
   Dim Total As Currency '---Totals scheduled payments
   Dim TotalPrincipal As Currency
   Dim Month As Integer  '---To traverse the 12 months
   Dim I As Integer      '---To traverse the amounts for selected month
   Dim Months As String
   Dim TotPortfolio As Currency   '---to build total portfolio principal
   Dim TotTaxFree As Currency     '---to build total tax free principal
   Months = "  January February    March    April      May     June" & _
            "     July   AugustSeptember  October November December"
   frmSummary.Top = 330  '---position the form
   frmSummary.Left = 1860
   If datInvestments.Recordset.RecordCount = 0 Then
      lstSummary.Clear
      Exit Sub
   End If
   frmSummary.lstResults.Clear   '---clear any old results
   frmSummary.lstResults.AddItem "             Scheduled Income"
   frmSummary.lstResults.AddItem " "
   For Month = 1 To 12
      datSearch.Recordset!Month = Month
      datSearch.Recordset.Update
      datSearch.Recordset.Requery   '---make the above assignment available
      MonthIncome = 0
      datScheduledPays.RecordSource = "SELECT ScheduledPays.SchdMonth, Amount, Search.Month " & _
         "FROM ScheduledPays, Search " & _
         "WHERE ScheduledPays.SchdMonth = Search.Month"
      datScheduledPays.Refresh
      If datScheduledPays.Recordset.RecordCount > 0 Then  '--if amounts exist, build them
         datScheduledPays.Recordset.MoveFirst
         For I = 1 To datScheduledPays.Recordset.RecordCount
            MonthIncome = MonthIncome + datScheduledPays.Recordset!Amount
            Total = Total + datScheduledPays.Recordset!Amount
            If I < datScheduledPays.Recordset.RecordCount Then
               datScheduledPays.Recordset.MoveNext
            End If
         Next I
      End If
      frmSummary.lstResults.AddItem "               " _
      & Mid$(Months, (Month - 1) * 9 + 1, 9) & _
      Right$(BLANKS & Format(MonthIncome, "currency"), 14)
   Next Month
   frmSummary.lstResults.AddItem " _____________________________________"
   frmSummary.lstResults.AddItem "     Total Annual Income" & _
          Right$(BLANKS & Format(Total, "currency"), 14)
   frmSummary.lstResults.AddItem ""
   frmSummary.lstResults.AddItem " " & lstSummary.List(0)
   frmSummary.lstResults.AddItem " " & lstSummary.List(1)
   '---now restore the scheduled payments
   datScheduledPays.RecordSource = "SELECT SchdMonth, SchdDay, Amount, " & _
        "InvestmentID " & _
        "FROM ScheduledPays, Search " & _
        "WHERE Search.ID = ScheduledPays.InvestmentID " & _
        "ORDER BY ScheduledPays.SchdMonth"
   datScheduledPays.Refresh
   TotPortfolio = 0
   TotTaxFree = 0
   datInvestments.Recordset.MoveFirst
   For I = 1 To datInvestments.Recordset.RecordCount
      TotPortfolio = TotPortfolio + datInvestments.Recordset!CurrentPrincipal
      If datInvestments.Recordset!TaxFree = "Yes" Then
         TotTaxFree = TotTaxFree + datInvestments.Recordset!CurrentPrincipal
      End If
      If I < datInvestments.Recordset.RecordCount Then  '---don't move past end record
         datInvestments.Recordset.MoveNext
      End If
   Next I
   lstSummary.Clear  '---erase old list
   lstSummary.AddItem datInvestments.Recordset.RecordCount _
   & " Investments Totaling " & Right$(BLANKS & Format(TotPortfolio, "currency"), 14)
   lstSummary.AddItem "   Tax Free portion is " & _
   Right$(BLANKS & Format(TotTaxFree, "currency"), 14) '---show portfolio total invests
   lstSummary.AddItem "    Total Annual Income" & _
          Right$(BLANKS & Format(Total, "currency"), 14)
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Sub ReBuildPrincipal()
   '---recalculates the running balance of the principal transactions when they change,
   '---then puts that total in the selected investments CurrentPrincipal,
   '---then refreshes dbgInvestments and dbgPrincipal
   Dim I As Integer
   Dim Total As Currency
   If datPrincipal.Recordset.RecordCount > 0 Then
      datPrincipal.Recordset.MoveFirst
      For I = 1 To datPrincipal.Recordset.RecordCount
         Total = Total + datPrincipal.Recordset!Amount
         datPrincipal.Recordset!Balance = Total
         datPrincipal.Recordset.MoveNext
      Next I
      datPrincipal.Recordset.MoveLast
      datPrincipal.Recordset.Requery
   End If
   datInvestments.Recordset!CurrentPrincipal = Total
   datInvestments.Recordset.Update
   datInvestments.Recordset.Requery
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Sub ReBuildIncome()
   '---recalculates the running balance of the income transactions when they change,
   '---then refreshes dbgIncome
   Dim I As Integer
   Dim Total As Currency
   If datIncome.Recordset.RecordCount > 0 Then
      datIncome.Recordset.MoveFirst
      For I = 1 To datIncome.Recordset.RecordCount
         Total = Total + datIncome.Recordset!Amount
         datIncome.Recordset!Balance = Total
         datIncome.Recordset.MoveNext
      Next I
      datIncome.Recordset.MoveLast
      datIncome.Recordset.Requery
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub dbgInvestments_HeadClick(ByVal ColIndex As Integer)
   '---Sort the grid by the column header clicked
   datInvestments.RecordSource = "Select * FROM Investment ORDER By " & _
      dbgInvestments.Columns(ColIndex).DataField
   datInvestments.Refresh
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub dbgInvestments_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   '---allows user to click a cell or row button to select an investment
   '---but do nothing if at BOF or EOF
   '---selecting an investment also refreshes principal, income and scheduled payments
   If datInvestments.Recordset.EOF = False And datInvestments.Recordset.BOF = False Then
      Call RefreshGrids   '---refresh the principal, income and scheduled pays grids
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Sub RefreshGrids()   '---refreshes the principal, income and scheduled pays grids
   datSearch.Recordset.MoveFirst
   datSearch.Recordset!ID = datInvestments.Recordset!InvestmentID  '---save investment ID
   datSearch.Recordset.Update
   datSearch.Recordset.Requery  '---make the above assignment available
   
   '---The Head and Tail simply give visual markings to even and odd years
   datPrincipal.RecordSource = "SELECT  Head, PrinDate, Amount, Balance, Tail, " & _
      "PrincipalID, InvestmentID " & _
      "FROM Principal, Search " & _
      "WHERE Principal.InvestmentID = Search.ID " & _
      "ORDER BY PrincipalID"
   datPrincipal.Refresh    '---refresh the list of principal transactions
   If datPrincipal.Recordset.RecordCount > 0 Then
      datPrincipal.Recordset.MoveLast  '---move to the most recent transaction if exists
   End If
      
   datIncome.RecordSource = "SELECT Head, IncmDate, Amount, Balance, Tail, " & _
      "IncomeID, InvestmentID " & _
      "FROM Income, Search " & _
      "WHERE Search.ID = Income.InvestmentID " & _
      "ORDER BY IncomeID"
   datIncome.Refresh       '---refresh the list of income transactions
   If datIncome.Recordset.RecordCount > 0 Then
      datIncome.Recordset.MoveLast  '---move to the most recent transaction if exists
   End If
      
   datScheduledPays.RecordSource = "SELECT SchdMonth, SchdDay, Amount, " & _
      "InvestmentID " & _
      "FROM ScheduledPays, Search " & _
      "WHERE Search.ID = ScheduledPays.InvestmentID " & _
      "ORDER BY ScheduledPays.SchdMonth"
   datScheduledPays.Refresh  '---refresh the list of scheduled payments
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub PrintCurrentInvestment()
   Dim a As String   '---divider character
   Dim Header As String
   Dim I As Integer   '---For Loop counter
   Indent = Space$(5)
   a = " "
   Header = "  Principal Entries W/Running Total  " & Mark _
          & "    Income Entries W/Running Total   " & Mark _
          & "    Income Entries W/Running Total   " & Mark & " Income Schedule"
   Printer.Font = "courier new"
   Printer.FontSize = 8.5
   Printer.FontBold = True
   Printer.FontUnderline = False
   Printer.Print   '---print a 3 line upper margin
   Printer.Print
   Printer.Print
   Printer.Print Indent; OwnerName; a; "    ";  '---start 1st line
   Printer.FontUnderline = True
   Printer.Print "Acquired:";
   Printer.FontUnderline = False
   Printer.Print Format(datInvestments.Recordset!AcquireDate, "MM/dd/yy"); " "; a;
   Printer.FontUnderline = True
   Printer.Print "Matures:";
   Printer.FontUnderline = False
   Printer.Print Format(datInvestments.Recordset!MatureDate, "MM/dd/yyyy"); " "; a;
   Printer.FontUnderline = True
   Printer.Print "Printed On:";
   Printer.FontUnderline = False
   Printer.Print Format(Now, "MM/dd/yy")
   
   Printer.Print Indent;  '---start 2nd line
   Printer.FontUnderline = True
   Printer.Print "Rate:";
   Printer.FontUnderline = False
   Printer.Print Format(datInvestments.Recordset!Rate, " ##.00"); a; a; a; a;
   Printer.Print datInvestments.Recordset!InvestmentName; a; "     ";
   Printer.FontUnderline = True
   Printer.Print "Broker: ";
   Printer.FontUnderline = False
   Printer.Print datInvestments.Recordset!BrokerNamePhone

   Printer.Print Indent;    '---start 3rd line
   Printer.FontUnderline = True
   Printer.Print "Cusip:";
   Printer.FontUnderline = False
   Printer.Print datInvestments.Recordset!CusipNumber; a; "  ";
   Printer.FontUnderline = True
   Printer.Print "Documents Location:";
   Printer.FontUnderline = False
   Printer.Print datInvestments.Recordset!DocumentsLocation; a; "  ";
   Printer.FontUnderline = True
   Printer.Print "Pays By:";
   Printer.FontUnderline = False
   Printer.Print datInvestments.Recordset!PaysBy; a; "  ";
   Printer.FontUnderline = True
   Printer.Print "Tax Free: ";
   Printer.FontUnderline = False
   Printer.Print datInvestments.Recordset!TaxFree
   
   For I = 1 To 4   '---print the first 4 lines of Notes
      If datInvestments.Recordset!Notes <> "" Then
         Printer.Print Indent; Mid$(datInvestments.Recordset!Notes, I * 106 - 105, 106)
      Else
         Printer.Print
      End If
   Next I
   Printer.FontSize = 7
   Printer.Print Indent; Header
   LineCount = 7
      
   PPtr = 1   '---point to 1st principal transaction location
   IPtr = 1   '---point to 1st income transaction location
   Do
      Call PrintTrans    '---start 4th column printing
      Call PrintISchedule  '---print income schedule & set ProgressFlag to 1 when done
   Loop While ProgressFlag = 0
   Call PrintTrans   '---4th column blank to separate next section
   Printer.Print
   Call PrintTrans   '---print lines that have income & principal totals (4th column)
   Printer.Print " Total Income"
   Call PrintTrans
   Printer.Print " During " & PrintYear
   Call PrintTrans
   Printer.Print Right$(BLANKS & Format(IncmTotal, "$##,###,###.00"), 14)
   Call PrintTrans
   Printer.Print " Total Principal"
   Call PrintTrans
   Printer.Print " During " & PrintYear
   Call PrintTrans
   Printer.Print Right$(BLANKS & Format(PrinTotal, "$##,###,###.00"), 14)
   Do   '---print the rest of the page
      If PPtr <= PCount Or IPtr <= ICount Then
         Call PrintTrans
         Printer.Print
      Else
         Printer.FontSize = 10   '---reset the printer
         Printer.EndDoc
         Exit Sub
      End If
   Loop While LineCount < 93
   
   Do   '---continue printing pages
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print Header
      Do  '---until out of transactions (not an infinite loop as it might appear)
         If PPtr <= PCount Or IPtr <= ICount Then
            Call PrintTrans
            Printer.Print
         Else
            Printer.FontSize = 10   '---reset the printer
            Printer.EndDoc
            Exit Sub
         End If
      Loop While LineCount < 93
   Loop
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub FillTransGetTotals(Yr As Integer)  '---fill buffers, get total principal
                                             '---and total income
   Dim Mark2 As String
   Dim I As Integer
   PCount = datPrincipal.Recordset.RecordCount
   ICount = datIncome.Recordset.RecordCount
   If PCount > 0 Then
      ReDim PBuf(1 To PCount)   '---size the principal buffer  (remember max 96 lines per
                 '---page with print sizes below <85 transactions on 1st page> )
                 '---and 2 columns of income transactions
   End If
   If ICount > 0 Then
      ReDim IBuf(1 To ICount)  '---size the income buffer
   End If
   '---arrays are filled with principal and income transactions to more easily place
   '---the transactions in the printed columns
   If datPrincipal.Recordset.RecordCount > 0 Then   '---prepare to traverse Prin. tran.
      datPrincipal.Recordset.MoveFirst
   End If
   For I = 1 To PCount  '---fill the principal transaction array
      Mark2 = " " '---set to odd numbered year with "no" mark2
      If datPrincipal.Recordset!Head = "*" Then  '---& reset to "*" if even year
         Mark2 = Chr$(149)   '--- this is a dot in courier new
      End If
      PBuf(I) = Mark2 _
         & Right$("  " & Format(Month(datPrincipal.Recordset!PrinDate), "##"), 2) & "/" _
         & Right$("  " & Format(Day(datPrincipal.Recordset!PrinDate), "##"), 2) & "/" _
         & Right$(Format(Year(datPrincipal.Recordset!PrinDate), "0#"), 2) _
         & Right$(BLANKS & Format(datPrincipal.Recordset!Amount, "$##,###,###.00 "), 14) _
         & Right$(BLANKS & Format(datPrincipal.Recordset!Balance, "$##,###,###.00"), 13) _
         & Mark2 & Mark
      datPrincipal.Recordset.MoveNext
   Next I
   If datIncome.Recordset.RecordCount > 0 Then   '---prepare to traverse Income tran.
      datIncome.Recordset.MoveFirst
   End If
   For I = 1 To ICount  '---fill the income transaction array
      Mark2 = " " '---set to odd numbered year with "no" mark2
      If datIncome.Recordset!Head = "*" Then  '---& reset to "*" if even year
         Mark2 = Chr$(149)  '--- this is a dot in courier new
      End If
      IBuf(I) = Mark2 _
         & Right$("  " & Format(Month(datIncome.Recordset!IncmDate), "##"), 2) & "/" _
         & Right$("  " & Format(Day(datIncome.Recordset!IncmDate), "##"), 2) & "/" _
         & Right$(Format(Year(datIncome.Recordset!IncmDate), "0#"), 2) _
         & Right$(BLANKS & Format(datIncome.Recordset!Amount, "$##,###,###.00 "), 14) _
         & Right$(BLANKS & Format(datIncome.Recordset!Balance, "$##,###,###.00"), 13) _
         & Mark2 & Mark
      datIncome.Recordset.MoveNext
   Next I
   datSearch.Recordset.MoveFirst
   datSearch.Recordset!Year = Yr
   datSearch.Recordset.Update
   datSearch.Recordset.Requery   '---make the above assignment available
   datPrincipal.RecordSource = "SELECT PrincipalID, PrinDate, Amount, " & _
       "InvestmentID, Search.Year, Search.ID " & _
       "FROM Principal, Search " & _
       "WHERE Year(PrinDate) = Search.Year " & _
       "AND InvestmentID = Search.ID " & _
       "ORDER BY PrincipalID"
   datPrincipal.Refresh
   PrinTotal = 0
   If datPrincipal.Recordset.RecordCount > 0 Then
      datPrincipal.Recordset.MoveFirst
      For I = 1 To datPrincipal.Recordset.RecordCount  '---build principal for specified year
         PrinTotal = PrinTotal + datPrincipal.Recordset!Amount
         datPrincipal.Recordset.MoveNext
      Next I
   End If
   
   datIncome.RecordSource = "SELECT IncomeID, IncmDate, Amount, " & _
       "InvestmentID, Search.Year,  Search.ID " & _
       "FROM Income, Search " & _
       "WHERE Year(IncmDate) = Search.Year " & _
       "AND InvestmentID = Search.ID " & _
       "ORDER BY IncomeID"
   datIncome.Refresh
   IncmTotal = 0
   If datIncome.Recordset.RecordCount > 0 Then
      datIncome.Recordset.MoveFirst
      For I = 1 To datIncome.Recordset.RecordCount   '---build income for specified year
         IncmTotal = IncmTotal + datIncome.Recordset!Amount
         datIncome.Recordset.MoveNext
      Next I
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub PrintTrans()  '---print the transactions for one line
   If PPtr <= PCount Then
      Printer.Print Indent; PBuf(PPtr);
      PPtr = PPtr + 1
   Else
      Printer.Print Indent; Space$(38);
   End If
   If IPtr <= ICount Then
      Printer.Print IBuf(IPtr);
      IPtr = IPtr + 1
   Else
      Printer.Print Space$(38);
   End If
   If IPtr + 84 <= ICount Then
      Printer.Print IBuf(IPtr + 84);   '---I was incremented so do 85 income tran later
   Else
      Printer.Print Space$(38);
   End If
   LineCount = LineCount + 1
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub PrintISchedule()
   If datScheduledPays.Recordset.EOF = False Then  '---print a scheduled payment and
             '---move to next one or print the first part of the line if no more schd.pays
      Printer.Print Format(datScheduledPays.Recordset!SchdMonth, "MM") _
           & Format(datScheduledPays.Recordset!SchdDay, "/dd") & " " _
           & Right$(BLANKS & Format(datScheduledPays.Recordset!Amount, "$###,###.00"), 11)
      datScheduledPays.Recordset.MoveNext
   Else
      Printer.Print
      ProgressFlag = 1
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdExit_Click()
   End
End Sub


