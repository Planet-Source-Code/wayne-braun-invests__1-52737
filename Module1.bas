Attribute VB_Name = "Module1"
Option Explicit
'---if this project had been done by a team, there would not be
'---so many global variables.
Public OwnerName As String   '---allow owner name of portfolio
Public BLANKS As String
Public PBuf() As String  '---array dimemsioned later to hold the principal trans. to print
Public IBuf() As String  '---array dimemsioned later to hold the income trans. to print
Public PCount As Integer '---the number of principal transactions
Public ICount As Integer '---the number of income transactions
Public PPtr As Integer  '---to traverse principal transactions while printing
Public IPtr As Integer  '---to traverse income transactions while printing
Public PrinTotal As Currency  '---for total principal for a year for printing
Public IncmTotal As Currency   '---for total income for a year for printing
Public PrintYear As String  '---the year for above totals
Public LineCount As Integer  '---count the printed lines
Public Mark As String   '---column marker
Public ProgressFlag As Integer  '---to monitor printing of scheduled pays & totals figures
Public Indent As String  '---printing indent
Public InvestID As Integer  '---to remember an innvestment to return to
