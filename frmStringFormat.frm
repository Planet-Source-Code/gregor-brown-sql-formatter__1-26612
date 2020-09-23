VERSION 5.00
Begin VB.Form frmStringFormat 
   BackColor       =   &H008080FF&
   Caption         =   "Format SQL v3.0"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmStringFormat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClearAll 
      Caption         =   "Clear &All"
      Height          =   390
      Left            =   2550
      TabIndex        =   6
      Top             =   7350
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   7425
      TabIndex        =   7
      Top             =   7350
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   390
      Left            =   6360
      TabIndex        =   5
      Top             =   7350
      Width           =   990
   End
   Begin VB.Frame fOutput 
      BackColor       =   &H008080FF&
      Caption         =   "O&utput"
      Height          =   615
      Left            =   150
      TabIndex        =   14
      Top             =   7200
      Width           =   2265
      Begin VB.OptionButton optFile 
         BackColor       =   &H008080FF&
         Caption         =   "Notepad"
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   225
         Width           =   915
      End
      Begin VB.OptionButton optClipBoard 
         BackColor       =   &H008080FF&
         Caption         =   "Clipboard"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   225
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "&Format String"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4035
      Width           =   3735
   End
   Begin VB.Frame fOptions 
      BackColor       =   &H008080FF&
      Caption         =   "O&ptions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2040
      TabIndex        =   8
      Top             =   150
      Width           =   6375
      Begin VB.CheckBox ckContinue 
         BackColor       =   &H008080FF&
         Caption         =   "Line Continuation"
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   525
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.TextBox txtLineLen 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4350
         TabIndex        =   13
         Text            =   "75"
         Top             =   450
         Width           =   390
      End
      Begin VB.CheckBox ckQuotes 
         BackColor       =   &H008080FF&
         Caption         =   "Double Quotes to Single Quotes"
         Height          =   255
         Left            =   3300
         TabIndex        =   11
         Top             =   225
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox ckVarible 
         BackColor       =   &H008080FF&
         Caption         =   "&Make Variable"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLineLen 
         BackColor       =   &H008080FF&
         Caption         =   "&Line Length:"
         Height          =   240
         Left            =   3300
         TabIndex        =   12
         Top             =   525
         Width           =   990
      End
   End
   Begin VB.TextBox txtNewString 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4440
      Width           =   8295
   End
   Begin VB.TextBox txtOldString 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1320
      Width           =   8295
   End
   Begin VB.TextBox txtVar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Text            =   "gstrSQL"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Modified by: Gregor Brown"
      Height          =   240
      Left            =   3600
      TabIndex        =   21
      Top             =   7440
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Written By: Chris Shell"
      Height          =   240
      Left            =   3600
      TabIndex        =   20
      Top             =   7200
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   3300
      TabIndex        =   19
      Top             =   375
      Width           =   1665
   End
   Begin VB.Label lblNewString 
      BackStyle       =   0  'Transparent
      Caption         =   "&New String:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lOldText 
      BackStyle       =   0  'Transparent
      Caption         =   "&String to be Formatted:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1950
   End
   Begin VB.Label lVar 
      BackStyle       =   0  'Transparent
      Caption         =   "&Varible Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   195
      Width           =   1215
   End
End
Attribute VB_Name = "frmStringFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *************************************************************
'  Format String
'  Chris Shell
'  http://www.cshellvb.com
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************

Const DIM_STR1                                    As String = "Dim "
Const DIM_STR2                                    As String = " as String"
Const CONT_STR                                    As String = " & _"
Const CONNECT_STR                                 As String = " & "

Dim maSQLVar()                                    As Integer

'**************************************
'Windows API/Global Declarations for :
'Create links from labels!
'**************************************

Public Enum OpType
    Startup = 1
    Click = 2
    FormMove = 3
    LinkMove = 4
End Enum

Dim Clicked As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Function CleanString(szOriginal)
    If szOriginal = "" Then
        CleanString = "NULL"
    Else
        CleanString = Substitute(szOriginal, Chr(34), "'")
        'CleanString = Substitute(szOriginal, "'", "''")
        'CleanString = Substitute(CleanString, "’", "’’")
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub CmdClearAll_Click()

    If MsgBox("Clear all text boxes?", vbYesNo + vbQuestion, "Clear") = vbYes Then
        txtNewString.Text = ""
        txtOldString.Text = ""
        
    End If
    
    Me.Refresh
    
End Sub

Private Sub cmdFormat_Click()
Dim bContinue As Boolean, strVariableName As String
Dim bQuote As Boolean, intMaxLineLength As Integer, bSQLSmart As Boolean

    bContinue = False
    strVariableName = ""
    bQuote = False
    bSQLSmart = False
    
    
    If Len(txtOldString.Text) = 0 Then
        MsgBox "No String entered!", vbExclamation
        Exit Sub
    End If
    
'    If ckSQLSmart.Value = vbChecked Then
'       bSQLSmart = True
'    End If
    
    If ckContinue.Value = vbChecked Then
       bContinue = True
    End If
    
    If ckQuotes.Value = vbChecked Then
        bQuote = True
    End If
    
    'set up the variable.
    strVariableName = txtVar.Text
    
    If IsNumeric(txtLineLen.Text) Then
      intMaxLineLength = CInt(txtLineLen.Text)
    Else
      'Default.
      intMaxLineLength = 50
    End If
        
    txtNewString.Text = FormatSQL(txtOldString.Text, bContinue, intMaxLineLength, strVariableName, bQuote)

    'Auto-tab
    SendKeys "{TAB}"

End Sub
Function Substitute(szBuff, szOldString, szNewString)
    Dim iStart
    Dim iEnd
    
    
    ''' Find first substring
    iStart = InStr(1, szBuff, szOldString)
    
    ''' Loop through finding substrings
    Do While iStart <> 0
        ''' Find end of string
        iEnd = iStart + Len(szOldString)
        ''' Concatenate new string
        szBuff = Left(szBuff, iStart - 1) & szNewString & Right(szBuff, Len(szBuff) - iEnd + 1)
        ''' Advance past new string
        iStart = iStart + Len(szNewString)
        ''' Find next occurrence
        iStart = InStr(iStart, szBuff, szOldString)
    Loop
    
    Substitute = szBuff
End Function

Function RemoveChar(sText As String, sChar As String) As String
    Dim iPos As Integer, iStart As Integer
    Dim sTemp As String
    iStart = 1


    Do
        iPos = InStr(iStart, sText, sChar)


        If iPos <> 0 Then
            sTemp = sTemp & Mid(sText, iStart, (iPos - iStart))
            iStart = iPos + 1
        End If
    Loop Until iPos = 0
    sTemp = sTemp & Mid(sText, iStart)
    RemoveChar = sTemp
End Function

Sub SQLVarPos(ByVal sSQL As String)

Dim intPosition                                   As Integer
Dim intLength                                     As Integer
Dim intClauseI                                    As Integer
Dim intI                                          As Integer
Dim intStartPos                                   As Integer
Dim intArrayArgument                              As Integer
Dim intTemp                                       As Integer
Dim intWordCount                                  As Integer
Dim strClause                                     As String
Dim strWord                                       As String

On Error GoTo ehHandle


intWordCount = CountWords(sSQL)

'initialize
intArrayArgument = 0
'Redim.
ReDim maSQLVar(0 To intWordCount - 1, 0 To 1)

'Add files to arrary.
ReadFileIntoArray App.Path & "\SqlClauses.txt"


'Go thru all words
For intI = 1 To intWordCount
  intPosition = 0
  'Get word.
  strWord = GetWord(sSQL, intI, intStartPos)
  'Now test to see if word is a key word.
  
  'Try reading file.
  For intClauseI = 0 To UBound(gastrClause)                'Add acronyms to list.
    strClause = gastrClause(intClauseI)
    Debug.Print strClause
    'Find out the number of words in the clause.
    If CountWords(strClause) = 1 Then
      'Check for a match.
      If StrComp(strWord, strClause, vbTextCompare) = 0 Then
        intPosition = intStartPos
        intLength = Len(strClause)
      End If
    Else
      'Compare first word.
      If StrComp(strWord, GetWord(strClause, 1, intTemp), vbTextCompare) = 0 Then
        intPosition = intStartPos
        'Get next word
        intI = intI + 1
        strWord = strWord & " " & GetWord(sSQL, intI, intStartPos)
        If StrComp(strWord, strClause, vbTextCompare) = 0 Then
          intLength = Len(strClause)
        Else
          'Set position to zero...No match.
          intPosition = 0
        End If
      End If
    End If
    'Check for match to exit for and write off values.
    If intPosition > 0 Then
      'Write values to array
      maSQLVar(intArrayArgument, 0) = intPosition
      maSQLVar(intArrayArgument, 1) = intLength
      intArrayArgument = intArrayArgument + 1
      Exit For
    End If
  Next intClauseI
 
Next intI
    
SUB_EXIT:
    Exit Sub
    
ehHandle:
    MsgBox "SQLVarPos: " & Err.Number & " - " & Err.Description
    Resume Next
    
    
    
End Sub

Private Sub cmdOK_Click()
Dim hFile As Integer
Dim sFilename As String
Dim iFileName As Integer

    

    If optClipBoard.Value = True Then
        ClipboardCopy txtNewString.Text
        MsgBox "Your code on the Clipboard, Enjoy!", vbExclamation
        
    Else
        
        iFileName = Int((10000 - 100 + 1) * Rnd + 100)
        'obtain the next free file handle from the system
        hFile = FreeFile
        sFilename = App.Path & "\tmp" & iFileName & ".txt"
         
        'open and save the textbox to a file
        Open sFilename For Output As #hFile
            Print #hFile, (txtNewString.Text)
        Close #hFile

        If Err.Number <> 0 Then
            MsgBox "Problem creating temporary file! The disk may be full or read only.", vbExclamation
            Err.Clear
            Exit Sub
        End If
        
        Call Shell("Notepad " & sFilename, vbNormalFocus)
                
        Kill sFilename
        
        MsgBox "Your code is in Notepad, Enjoy!", vbExclamation
        
    End If
    

    Unload Me
    

End Sub
Public Sub ClipboardCopy(Text As String)
'Copies text to the clipboard
On Error GoTo error
    Clipboard.Clear
    Clipboard.SetText Text$
    
Exit Sub

error:  MsgBox Err.Description, vbExclamation, "Error"

End Sub






Private Sub Form_Resize()
    ResizeForm Me
    
End Sub

Public Function FormatSQL(strSql As String, booLineContinue As Boolean, intMaxLineLength As Integer, strVariableName As String, booFixQuotes As Boolean) As String

Dim booEnd                                        As Boolean
Dim booKeyWordBreak                               As Boolean
Dim booKeyWordNoTab                               As Boolean
Dim intArrayArgument                              As Integer
Dim intI                                          As Integer
Dim intLineCount                                  As Integer
Dim intLineLength                                 As Integer
Dim intStartPos                                   As Integer
Dim intTotalLineLength                            As Integer
Dim intWordCount                                  As Integer
Dim intWordI                                      As Integer
Dim strSqlPart                                    As String
Dim strWord                                       As String

'********************************************
'This class was written by:
'   Karl E. Peterson
'   http://www.mvps.org/vb/
'See the class for more detail. Thank you to
'him for this code, I got it via VBPJ Article
'on string building in ASP.
'********************************************
Dim cSBld As New CStringBuilder

On Error GoTo ehHandle
    
        
'********************************************
'Clean up String before we begin...
'********************************************
  'Remove Tab Characters
  strSql = RemoveChar(strSql, CStr(vbTab))
  
  'Remove Carriage Returns
  strSql = Replace(strSql, CStr(vbCr), CStr(Chr(32)))
  
  'Remove Line Feeds
  strSql = RemoveChar(strSql, CStr(vbLf))
  
  'Remove extra Spaces
  strSql = Trim(strSql)
        
    
'********************************************
'Ready to Rock...
'********************************************
  'Replace any quotes with single quotes if desired
  If booFixQuotes = True Then
      strSql = CleanString(strSql)
  End If
 
  'If a variable is given te use it...
  If Len(strVariableName) > 0 Then
    If ckVarible.Value = vbChecked Then
      cSBld.Append DIM_STR1 & strVariableName & DIM_STR2 & vbCrLf & vbCrLf
    End If
  End If
  
  'Set Key SQL Positions in Array
  Call SQLVarPos(strSql)
  
  'Initialize.
  booEnd = False
  booKeyWordBreak = False
  booKeyWordNoTab = False
  intArrayArgument = 1
  
  'Start with append of variable name.
  cSBld.Append strVariableName & " = "
      
  'Store original length
  intTotalLineLength = Len(strSql)
  
  'Check for a short SQL line.
  If intTotalLineLength <= intMaxLineLength Then
    cSBld.Append ContinueString(strSql, True)
  Else
    
    'Get word count.
    intWordCount = CountWords(strSql)
    
    'Go thru all words.
    For intWordI = 1 To intWordCount
      'Get word.
      strWord = GetWord(strSql, intWordI, intStartPos)
      
      'Build.
      strSqlPart = strSqlPart & strWord & " "
          
      'Look for key SQL Clause at the start position of the next word.
      If maSQLVar(intArrayArgument, 0) = intStartPos + Len(strWord) + 1 Then
        'Set the array argument.
        intArrayArgument = intArrayArgument + 1
        booKeyWordBreak = True
      End If
      
      'Check length.
      intLineLength = Len(strSqlPart)
      
      'Check for end of SQL.
      If intWordI = intWordCount Then
        booEnd = True
        strSqlPart = Left(strSqlPart, Len(strSqlPart) - 1)
      End If
        
      'Append if needed.
      If intLineLength >= intMaxLineLength Or booEnd Or booKeyWordBreak Then
        intLineCount = intLineCount + 1
        
        'Check if the user wants line continuation.
        If booLineContinue Then
          
          'Add in general tab for every line but the first.
          If intLineCount > 1 Then cSBld.Append "  "
            
          'Check for no tab.
          If booKeyWordNoTab Then
            'Set the boolean back to its origional value.
            booKeyWordNoTab = False
          Else
            'Add in extra tab for none keywords (if not the first line).
            If intLineCount > 1 Then cSBld.Append "  "
          End If
          'Set the boolean for next loop.
          If booKeyWordBreak Then booKeyWordNoTab = True
          
          'Add the string.
          cSBld.Append ContinueString(strSqlPart, booEnd)
          Debug.Print strSqlPart
        Else
          
          cSBld.Append AppendString(strVariableName, strSqlPart, booEnd)
        End If
  
        'Append line feed.
        cSBld.Append vbCrLf
      
        'Reset temp string.
        strSqlPart = ""
         
        'Reset value.
        booKeyWordBreak = False
      End If
      
    
    Next intWordI
  End If
    
  'Pass the String Back...
  FormatSQL = cSBld.ToString
  
  Set cSBld = Nothing
    
    
ExitFunc:
    Exit Function


ehHandle:
    MsgBox "ERROR: " & Err.Number & " - " & Err.Description
    Resume Next


End Function

Private Function ContinueString(sLine As String, bEnd As Boolean) As String
    
    If bEnd Then
        ContinueString = Chr(34) & sLine & Chr(34)
    Else
        ContinueString = Chr(34) & sLine & Chr(34) & CONT_STR
    End If
    

End Function

Private Function AppendString(sVar As String, sLine As String, bEnd As Boolean) As String

    If bEnd Then
         AppendString = sVar & " = " & sVar & CONNECT_STR & _
                                Chr(34) & sLine & Chr(34)
    Else
         AppendString = sVar & " = " & sVar & CONNECT_STR & _
                                Chr(34) & sLine & Chr(34)
    End If

End Function


Private Sub txtNewString_GotFocus()

'To highlight text on enter.
txtNewString.SelStart = 0
'A very high value always works.
txtNewString.SelLength = 19999

End Sub




Private Sub txtOldString_GotFocus()
'To highlight text on enter.
txtOldString.SelStart = 0
'A very high value always works.
txtOldString.SelLength = 19999


End Sub


