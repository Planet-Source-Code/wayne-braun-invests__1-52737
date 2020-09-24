VERSION 5.00
Begin VB.Form frmPrinting 
   BackColor       =   &H00EDD8E8&
   Caption         =   "    Printing Choices"
   ClientHeight    =   5670
   ClientLeft      =   1710
   ClientTop       =   1860
   ClientWidth     =   6390
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPrinting.frx":000C
   ScaleHeight     =   5670
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdYrEndReports 
      BackColor       =   &H00CF87BA&
      Caption         =   "Print Year End Reports"
      Height          =   1065
      Left            =   -45
      Picture         =   "frmPrinting.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4095
      Width           =   2040
   End
   Begin VB.CommandButton cmdPrintAllInvestments 
      BackColor       =   &H00CF87BA&
      Caption         =   "Print All Investments"
      Height          =   1065
      Left            =   180
      Picture         =   "frmPrinting.frx":0E60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2040
   End
   Begin VB.CommandButton cmdPrintInvestment 
      BackColor       =   &H00CF87BA&
      Caption         =   "Print Selected Investment"
      Height          =   1065
      Left            =   180
      Picture         =   "frmPrinting.frx":172A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1575
      Width           =   2040
   End
   Begin VB.CommandButton cmdPrintIncomeSchedule 
      BackColor       =   &H00CF87BA&
      Caption         =   "Print Portfolio Income Schedule"
      Height          =   1065
      Left            =   3060
      Picture         =   "frmPrinting.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1620
      Width           =   2040
   End
End
Attribute VB_Name = "frmPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()

End Sub
