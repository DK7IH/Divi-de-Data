Attribute VB_Name = "Module1"
Option Explicit

Global strData(230000) As String
Global Const MAXDATA = 228000


Function GetInfoFromString(strS0 As String, intPos As Integer) As String
  
  Dim T1 As Integer
  Dim intCnt As Integer
  Dim intEnd As Integer
  Dim strOut As String
  Dim strS As String
  Dim intP1 As Integer
  
  strOut = ""
  intCnt = 0
  strS = strS0
  
  intP1 = 1
  For T1 = 1 To Len(strS)
    If Mid$(strS, T1, 1) = "," Or T1 = Len(strS) Then
      intCnt = intCnt + 1
      
      If intCnt = intPos Then
          If T1 = Len(strS) Then
              strOut = Mid$(strS, intP1, T1 - intP1 + 1)
          Else
              strOut = Mid$(strS, intP1, T1 - intP1)
          End If
          T1 = Len(strS)
      End If
      intP1 = T1 + 1
    End If
      
  Next
  GetInfoFromString = strOut
  
End Function

Function LeadingZero(strS As String, intN As Integer) As String
  
  LeadingZero = String(intN - Len(strS), "0") & strS
  
End Function

