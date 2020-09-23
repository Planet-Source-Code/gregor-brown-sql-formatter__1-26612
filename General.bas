Attribute VB_Name = "modGeneral"
Option Explicit
Public gastrClause()                              As String

Public Function GetWord(S, Indx As Integer, SPos As Integer)
'
' Extracts a word in text where words are separated by 1 or more spaces.
' SPos - Returns the start position of the word.
Dim i As Integer, WC As Integer, Count As Integer, EPos As Integer, OnASpace As Integer
  WC = CountWords(S)
  If Indx < 1 Or Indx > WC Then
    GetWord = Null
    Exit Function
  End If
  Count = 0
  OnASpace = True
  For i = 1 To Len(S)
    If Mid(S, i, 1) = " " Then
      OnASpace = True
    Else
      If OnASpace Then
        OnASpace = False
        Count = Count + 1
        If Count = Indx Then
          SPos = i
          Exit For
        End If
      End If
    End If
  Next i
  EPos = InStr(SPos, S, " ") - 1
  If EPos <= 0 Then EPos = Len(S)
  GetWord = Mid(S, SPos, EPos - SPos + 1)
End Function


Public Function CountWords(S) As Integer
'
' Counts words in a string separated by 1 or more spaces
'
Dim WC As Integer, i As Integer, OnASpace As Integer
  If VarType(S) <> 8 Or Len(Trim(S)) = 0 Then
    CountWords = 0
    Exit Function
  End If
  WC = 0
  OnASpace = True
  For i = 1 To Len(S)
    If Mid(S, i, 1) = " " Then
      OnASpace = True
    Else
      If OnASpace Then
        OnASpace = False
        WC = WC + 1
      End If
    End If
  Next i
  CountWords = WC
End Function

Public Function ReadFileIntoArray(strFileName As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------
'Purpose    :Takes a file and reads into an array.
'Parameters :
' [strFileName] File name to be read into array.
' [lngNumberOfRecords] Returns the number of records in the file.
'Returns    :
'Created By :Gregor L. Brown
'Created On :02.01.2001 16:06.31
'Comments   :
'---------------------------------------------------------------------------------------------------------------------------

Dim lngN                      As Long
Dim lngFileNum                As Long

ReadFileIntoArray = False

'Set up the variables.
lngN = 0
lngFileNum = FreeFile()

'Open text file.
Open strFileName For Input As #lngFileNum

'Loop thru each line.
Do Until EOF(lngFileNum)
    
  ReDim Preserve gastrClause(0 To lngN)
  Input #lngFileNum, gastrClause(lngN)
  
  lngN = lngN + 1

Loop

'Close out file
Close #lngFileNum

ReadFileIntoArray = True

End Function



