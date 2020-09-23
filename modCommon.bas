Attribute VB_Name = "modCommon"
'File Open Dialog

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10



'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS

Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
 End Type

Public OFN As OPENFILENAME

Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" _
   (pOpenfilename As OPENFILENAME) As Long

Public Declare Function CommDlgExtendedError _
    Lib "comdlg32.dll" () As Long

Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
   (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long


Public Function FileSave_Dialog(bShortName As Boolean, strTitle As String, strFileDesc As String, strFilter As String, bAll As Boolean, _
    strFileName As String, OwnerHwnd As Long, Optional strInitDir As String) As String

  Dim r As Long
  Dim sp As Long
  Dim LongName As String
  Dim shortName As String
  Dim ShortSize As Long

 'to keep lines short(er), I've abbreviated a
 'Null$ to n and n2, and the filter$ to f.
  Dim n As String
  Dim n2 As String
  Dim f As String
  n = Chr$(0)
  n2 = n & n

 '------------------------------------------------
 'INITIALIZATION
 '------------------------------------------------
 'fill in the size of the OFN structure
  OFN.nStructSize = Len(OFN)

 'assign the owner of the dialog; this can
 'be null if no owner.
  OFN.hwndOwner = OwnerHwnd

 'Set filter i.e. *.txt
  f = strFileDesc & n & strFilter
  If bAll Then
    f = f & n & "All Files" & n & "*.*" & n2
  
  End If
    
  OFN.sFilter = f

 'Set filter index or which filter is used at startup
 
    OFN.nFilterIndex = 1   '  "User Input"

'Set default file name and space for path

    OFN.sFile = strFileName & Space$(1024) & n
    OFN.nFileSize = Len(OFN.sFile)

 'default extension applied to a selected file if
 'it has no extension.
  OFN.sDefFileExt = Right$(sFilter, 3) & n
  

 'File Title
  OFN.sFileTitle = Space$(512) & n
  OFN.nTitleSize = Len(OFN.sFileTitle)

 'Startup directory
 
    If IsNull(strInitDir) Or strInitDir = "" Then
    
        OFN.sInitDir = n
    Else
        OFN.sInitDir = strInitDir & n
    End If
'Dialog Title

  OFN.sDlgTitle = strTitle & n

 'flags are the actions and options for the dialog.
  OFN.flags = OFS_FILE_SAVE_FLAGS

 'Finally, show the File Open Dialog
  r = GetOpenFileName(OFN)
  
  If r Then
   
    'Path & File Returned (OFN.sFile):
     If bShortName Then
        FileOpen_Dialog = Mid$(OFN.sFile, OFN.nFileOffset + 1, _
                  Len(OFN.sFile) - OFN.nFileOffset - 1)
     Else
        FileOpen_Dialog = OFN.sFile
     End If

  End If


End Function
