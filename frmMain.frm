VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Divi Datenanalyse"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdDataByState2 
      Caption         =   "Process State Data"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdReadData 
      Caption         =   "Read Raw Data"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDataByState 
      Caption         =   "Raw Data by State"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Peter Baier (http://coronadiktatur.wordpress.com)"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label lblProcData 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblDataByState 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblCountLines 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdDataByState_Click()
 
  Dim T0, T1 As Long
  Dim maxStates As Integer
  Dim strState As String
    
  maxStates = 16
  
  For T0 = 1 To maxStates
    strState = LeadingZero(Format(T0), 2)
    Open App.Path & "\RawDataByState\" & strState & ".csv" For Output As #1
    For T1 = 1 To MAXDATA
      If strState = GetInfoFromString(strData(T1), 2) Then
        Print #1, strData(T1)
      End If
      lblDataByState.Caption = "State: " & strState & " Lines: " & Format(T1)
      lblDataByState.Refresh
    Next
    Close #1
  Next
  
    
    
End Sub

Private Sub cmdDataByState2_Click()
  
  Dim T0, T1 As Long
  Dim strS As String
  Dim maxStates As Integer
  Dim strDate As String
  Dim strState As String
  Dim strOut As String
  Dim intBedsOccupied As Integer
  Dim intData(7) As Integer
  Dim intDataT(7) As Integer
  maxStates = 16
  
  For T0 = 1 To maxStates
    strState = LeadingZero(Format(T0), 2)
    Open App.Path & "\RawDataByState\" & strState & ".csv" For Input As #1
    Open App.Path & "\ProcessedDataByState\" & strState & ".csv" For Output As #2
    Print #2, "datum, faelle_covid_aktuell,faelle_covid_aktuell_invasiv_beatmet,betten_frei,betten_belegt,betten_belegt_nur_erwachsen,betten_frei_nur_erwachsen,betten_gesamt"

    lblProcData.Caption = "State: " & strState
    lblProcData.Refresh
    
    While Not EOF(1)
      Line Input #1, strS
      
      strDate = GetInfoFromString(strS, 1)
      
      For T1 = 6 To 11
          intDataT(T1 - 5) = 0
      Next
      
      While GetInfoFromString(strS, 1) = strDate
        For T1 = 6 To 11
          intData(T1 - 5) = GetInfoFromString(strS, Val(T1))
          intDataT(T1 - 5) = intDataT(T1 - 5) + intData(T1 - 5)
        Next
        
        If (Not EOF(1)) Then
          Line Input #1, strS
        Else
          strS = ""
        End If
          
      Wend
      
      strOut = ""
      
      For T1 = 1 To 6
        strOut = strOut & Format(intDataT(T1)) & ","
      Next
      intBedsOccupied = intDataT(3) + intDataT(4)
      strOut = strOut & Format(intBedsOccupied)
      Print #2, strDate & "," & strOut
      
    Wend
    Close #1
    Close #2
    
    lblProcData.Caption = "State: " & strState
    lblProcData.Refresh
  
  Next
  
End Sub

Private Sub cmdReadData_Click()
    
    Dim intCnt As Long
    Dim strLine As String
    intCnt = 1
    
    Open App.Path & "\daten.csv" For Input As #1
    
    While Not EOF(1)
      Line Input #1, strLine
      strData(intCnt) = strLine
      intCnt = intCnt + 1
      lblCountLines.Caption = Format(intCnt) & " lines processed."
      lblCountLines.Refresh
    Wend
    
    Close #1
    Debug.Print intCnt
    

End Sub
