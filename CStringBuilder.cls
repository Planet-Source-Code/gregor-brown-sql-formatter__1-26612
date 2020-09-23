VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
'  Copyright ©2000-2001, Karl E. Peterson
'  http://www.mvps.org/vb/
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit

' Win32 APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
' *** Use proper function, depending on VB Version!
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Var() As Any) As Long
' Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

' Member variables
Private m_Bump As Long
Private m_Len As Long
Private m_BufferLen As Long
Private m_Buffer() As Byte
Private m_Scratch() As Byte

' Internal class constants
Private Const defBufferBump As Long = 2048
Private Const BaseModule As String = "CStringBuilder"
Private Const errArgumentOutOfRange As Long = vbObjectError + 512

' ************************************************
'  Initialization and Termination
' ************************************************
Private Sub Class_Initialize()
   ' dotNET uses a 16-char default size, but this
   ' class is really written for heavier use where
   ' longer strings will be more the norm.
   m_Bump = defBufferBump
   m_Len = 0
   ReDim m_Buffer(0 To m_Bump - 1) As Byte
   ReDim m_Scratch(0 To m_Bump - 1) As Byte
   m_BufferLen = UBound(m_Buffer) + 1
End Sub

' ************************************************
'  Properties
' ************************************************
Public Property Let BufferBump(ByVal NewVal As Long)
   ' Accept new value (bytes) by which to resize buffer
   ' when adjustment is needed.  Good to keep this
   ' fairly large in most cases.
   If NewVal > 0 Then
      ' Round up, if user passed odd number.
      If NewVal Mod 2 Then
         NewVal = NewVal + 1
      End If
      ' Store number of bytes to bump buffer by.
      m_Bump = NewVal
   End If
End Property

Public Property Get BufferBump() As Long
   ' Return number of bytes buffer gets bumped by.
   BufferBump = m_Bump
End Property

Public Property Let Capacity(ByVal NewVal As Long)
   Dim NewSize As Long
   ' This is an *expensive* routine, to be avoided
   ' at most reasonable costs. Adjusts the number of
   ' *bytes* in string buffer.
   
   If NewVal > 0 Then
      ' Adjust buffer to hold at least NewVal bytes,
      ' adjusting to next larger bufferbump size.
      NewSize = (NewVal \ m_Bump) * m_Bump
      If (NewVal Mod m_Bump) Then
         NewSize = NewSize + m_Bump
      End If
            
      ' Adjust end-of-buffer pointer to match
      ' truncated buffer, if needed.
      If NewSize < m_Len Then
         m_Len = NewSize
      End If
      
      ' Resize buffer.
      ReDim Preserve m_Buffer(0 To NewSize - 1) As Byte
      ' Ensure scratch space matches.
      ReDim m_Scratch(0 To NewSize - 1) As Byte
      
      ' Cache length of new buffer.
      m_BufferLen = UBound(m_Buffer) + 1
   End If
End Property

Public Property Get Capacity() As Long
   ' Return number of *bytes* in buffer.
   Capacity = m_BufferLen
End Property

Public Property Let Chars(Optional ByVal Position As Long = 0, ByRef NewVal As String)
   ' Replaces the character at Position index within
   ' the current StringBuilder. Note that the index
   ' of the first character is 0.
   If Position >= 0 And Position < m_Len Then
      ' Assign passed character.
      Call CopyMemory(m_Buffer(Position * 2), AscW(NewVal), 2)
   Else
      ' Tried to assign past end of current buffer.
      Err.Raise Number:=errArgumentOutOfRange, _
         Source:=BaseModule & ".Chars", _
         Description:="Tried assigning character past end of buffer."
   End If
End Property

Public Property Get Chars(Optional ByVal Position As Long = 0) As String
   Static nResult As Long
   ' Retrieves the character at Position index within
   ' the current StringBuilder. Note that the index
   ' of the first character is 0.
   If Position >= 0 And Position < m_Len Then
      Call CopyMemory(nResult, m_Buffer(Position * 2), 2)
      Chars = ChrW(nResult)
   Else
      ' Tried to read past end of current buffer
      Err.Raise Number:=errArgumentOutOfRange, _
         Source:=BaseModule & ".Chars", _
         Description:="Tried reading character past end of buffer."
   End If
End Property

Public Property Let Length(ByVal NewVal As Long)
   ' Sets length, in *chars*, of stringbuilder.
   ' Value must not be less than zero.
   If NewVal >= 0 Then
      ' Resize buffer if not big enough.
      If (NewVal * 2) > m_BufferLen Then
         Me.Capacity = NewVal * 2
      End If
      ' Adjust end-of-buffer pointer.
      m_Len = NewVal * 2
   End If
End Property

Public Property Get Length() As Long
   ' Returns length, in *chars*, of stringbuilder.
   Length = m_Len \ 2
End Property

' ************************************************
'  Public Methods
' ************************************************
Public Function Append(ByRef NewText As String) As CStringBuilder
   Static LenNewText As Long
   
   ' Cache length of passed text.
   LenNewText = LenB(NewText)
   
   ' Don't bother unless new text was passed.
   If LenNewText Then
   
      ' Resize buffer, if needed.
      Me.EnsureCapacity m_Len + LenNewText
      
      ' Copy text into buffer.
      Call CopyMemory(m_Buffer(m_Len), ByVal StrPtr(NewText), LenNewText)
      
      ' Extend end-of-buffer pointer.
      m_Len = m_Len + LenNewText
   End If
   
   ' Return reference to Me
   Set Append = Me
End Function

Public Function EnsureCapacity(ByVal NewVal As Long) As Long
   ' Resize buffer, if needed.
   If NewVal > m_BufferLen Then
      Me.Capacity = NewVal
   End If
   
   ' Return size of buffer in bytes.
   EnsureCapacity = m_BufferLen
End Function

Public Function Insert(ByVal StartIndex As Long, ByRef NewText As String) As CStringBuilder
   Static LenNewText As Long
   
   ' Cache length of passed text.
   LenNewText = LenB(NewText)
   
   ' Precalculate pointer to start index, accounting
   ' for double-byte count.
   StartIndex = StartIndex * 2
   
   ' Reality check arguments, and bail if bad.
   If (StartIndex < 0) Or (StartIndex > m_BufferLen) Then
      ' Toss out-of-bounds error.
      Err.Raise Number:=errArgumentOutOfRange, _
         Source:=BaseModule & ".Insert", _
         Description:="Argument(s) out of range."
      Exit Function
   End If
   
   ' Don't bother unless new text was passed.
   If LenNewText Then
   
      ' Resize buffer, if needed.
      Me.EnsureCapacity m_Len + LenNewText
      
' **************************************************************
'  Bill McCarthy's double-buffer optimization method...
'  A *lot* more API calls, but definitely faster,
'  because burden is even across all inserts.
      ' Copy beginning of main buffer into scratch buffer.
      Call CopyMemory(m_Scratch(0), _
         m_Buffer(0), StartIndex)
      ' Copy passed text into scratch buffer.
      Call CopyMemory(m_Scratch(StartIndex), _
         ByVal StrPtr(NewText), LenNewText)
      ' Copy main buffer remainder into scratch buffer.
      Call CopyMemory(m_Scratch(StartIndex + LenNewText), _
         m_Buffer(StartIndex), m_Len - StartIndex)
      ' Swap main and scratch buffers.
      Call SwapArrays
' **************************************************************
      
' **************************************************************
'  Double-buffer method...
'  Slower than above, especially on prepends, because
'  overhead increases as insert approaches position 0.
'      ' Sling existing characters to scratch buffer.
'      Call CopyMemory(m_Scratch(0), _
'         m_Buffer(StartIndex), m_Len - StartIndex)
'
'      ' Insert passed text into buffer.
'      Call CopyMemory(m_Buffer(StartIndex), _
'         ByVal StrPtr(NewText), LenNewText)
'
'      ' Return scratch string to main buffer.
'      Call CopyMemory(m_Buffer(StartIndex + LenNewText), _
'         m_Scratch(0), m_Len - StartIndex)
' **************************************************************

' **************************************************************
'  Overlapped copy method...
'  Slowest, because RtlMoveMemory operates one
'  byte at a time on overlapped requests?
'      ' Slide existing characters up in buffer.
'      Call CopyMemory(m_Buffer(StartIndex + LenNewText), _
'         m_Buffer(StartIndex), m_Len - StartIndex)
'      ' Insert passed text into buffer.
'      Call CopyMemory(m_Buffer(StartIndex), _
'         ByVal StrPtr(NewText), LenNewText)
' **************************************************************

      ' Extend end-of-buffer pointer.
      m_Len = m_Len + LenNewText
   End If
   
   ' Return reference to Me
   Set Insert = Me
End Function

Public Function Remove(ByVal StartIndex As Long, Optional ByVal Length As Long = 1) As CStringBuilder
   ' Adjust arguments for byte count.
   Length = Length * 2
   StartIndex = StartIndex * 2
   
   ' Reality check arguments, and bail if bad.
   ' Plus2 on 3rd test accounts for 0-base.
   If (StartIndex < 0) Or (Length < 0) Or _
      ((StartIndex + Length) > (m_Len + 2)) Then
      ' Toss out-of-bounds error.
      Err.Raise Number:=errArgumentOutOfRange, _
         Source:=BaseModule & ".Remove", _
         Description:="Argument(s) out of range."
      Exit Function
   End If
   
' **************************************************************
'  Double-buffer method...
'  Unlike Insert, this isn't as fast as overlapped?
'      ' Copy beginning of main buffer into scratch buffer.
'      Call CopyMemory(m_Scratch(0), _
'         m_Buffer(0), StartIndex)
'      ' Copy main buffer remainder into scratch buffer.
'      Call CopyMemory(m_Scratch(StartIndex), _
'         m_Buffer(StartIndex + Length), m_Len - StartIndex)
'      ' Swap main and scratch buffers.
'      Call SwapArrays
' **************************************************************

' **************************************************************
'  Overlapped copy method...
      ' Slide text backwards by Length chars. Note that
      ' the index of the first character is 0.
      Call CopyMemory(m_Buffer(StartIndex), _
         m_Buffer(StartIndex + Length), _
         m_Len - StartIndex - Length)
' **************************************************************
   
   ' Retract end-of-buffer pointer.
   m_Len = m_Len - Length
   
   ' Return reference to Me
   Set Remove = Me
End Function

Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
   ' Return truncated buffer contents.
   If m_Len > 0 Then
      ToString = Space$(m_Len \ 2)
      Call CopyMemory(ByVal StrPtr(ToString), m_Buffer(0), m_Len)
   End If
End Function

' ************************************************
'  Private Methods
' ************************************************
Private Sub SwapPtrs(s1 As String, s2 As String)
   Static lpTmp As Long
   lpTmp = StrPtr(s1)
   Call CopyMemory(ByVal VarPtr(s1), ByVal VarPtr(s2), 4&)
   Call CopyMemory(ByVal VarPtr(s2), lpTmp, 4&)
End Sub

Private Sub SwapArrays()
   Static lpSA As Long
   CopyMemory lpSA, ByVal VarPtrArray(m_Buffer), 4&
   CopyMemory ByVal VarPtrArray(m_Buffer), ByVal VarPtrArray(m_Scratch), 4&
   CopyMemory ByVal VarPtrArray(m_Scratch), lpSA, 4&
End Sub


