VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event BlockRead(ByVal lngCurrentBlock As Long, ByVal lngTotalBlocks As Long, ByVal lBlockMode As enBlockMode, ByRef bCancel As Boolean)
Public Event FileDecodeComplete(strOriginalFile As String, strDecodedFile As String)
Public Event FileEncodeComplete(strOriginalFile As String, strEncodedFile As String)
Public Event ErrorOccured(ByVal lngCode As Long, ByVal strDescription As String)
Public Event BeforeFileOpenIn(ByVal strFileName As String, ByRef bCancel As Boolean)
Public Event BeforeFileOpenOut(ByVal strFileName As String, ByRef bCancel As Boolean)
Public Event AfterFileCloseIn(ByVal strFileName As String)
Public Event AfterFileCloseOut(ByVal strFileName As String)

Private m_oXMLDoc As Object
Private m_oXMLElement As Object
Private m_bXMLObjectOK As Boolean

Private m_lngLastErrorCode As Long
Private m_strLastErrorDesc As String
Private m_lngEncBlockSize As Long
Private m_lngDecBlockSize As Long
Private m_bCancel As Boolean

Public Enum enBlockMode
    b64Encode
    b64Decode
End Enum

Private Const cMSXMLDom = "MSXML.DomDocument"
Private Const cDefBlockSize = 1048576 'Default block size: 1Mb

Private Sub Class_Initialize()
    Set m_oXMLDoc = fCreateObject(cMSXMLDom)
    ClearError
    m_bXMLObjectOK = Not m_oXMLDoc Is Nothing
    If Not m_bXMLObjectOK Then
        'This error will never be raised, because we are in the _Initialize event...
        'However, it WILL store the last error...
        RaiseError "90001", "Error creating required '" & cMSXMLDom & "' object"
    Else
        Set m_oXMLElement = m_oXMLDoc.createElement("TmpBase64")
        m_oXMLElement.dataType = "bin.base64"
    End If
    'Initialize blocksizes to approx. 1Mb
    Me.BlockSize = cDefBlockSize
End Sub

Private Sub Class_Terminate()
    'Free objects
    Set m_oXMLElement = Nothing
    Set m_oXMLDoc = Nothing
End Sub

Public Function Encode(ByVal strIn As String) As String
    'Encode a string to base64
    Dim oElement As Object
    Dim sData As String
    Dim aData() As Byte
    
    On Error GoTo errEncode:
    If m_bXMLObjectOK And Len(strIn) > 0 Then
        aData = StrConv(strIn, vbFromUnicode)
        m_oXMLElement.nodetypedvalue = aData
        Encode = m_oXMLElement.Text
        m_oXMLElement.Text = "" 'Free memory
    End If
    Exit Function
errEncode:
    RaiseError Err.Number, Err.Description
End Function

Public Function Decode(ByVal strIn As String) As String
    'Decode a base64 string
    Dim oElement As Object
    
    On Error GoTo errDecode:
    If m_bXMLObjectOK And Len(strIn) > 0 Then
        m_oXMLElement.Text = strIn
        Decode = StrConv(m_oXMLElement.nodetypedvalue, vbUnicode)
        m_oXMLElement.Text = "" 'Free memory
    End If
    Exit Function
errDecode:
    RaiseError Err.Number, Err.Description
End Function

Public Function EncodeFile(ByRef strInFile As String, ByRef strOutFile As String) As Boolean
    'Read blocks of the file in a multiple of 54 bits. This ensures, that because
    'of base64 is 3:4 the output will always be 72bits (+ 8 bits for a LineFeed)
    'Return True if succesful, else false
    EncodeFile = EncDecFile(strInFile, strOutFile, True)
End Function

Public Function DecodeFile(ByRef strInFile As String, ByRef strOutFile As String) As Boolean
    'Read the file in a multiple of 584bits (73 bytes). When filesize is NOT a
    'multiple of 73bytes it will be invalid anyway (Base64 encoded data is always
    'a multiple of 73 bytes). These chunks can be decoded.
    'Return True if succesful, else false
    DecodeFile = EncDecFile(strInFile, strOutFile, False)
End Function

Public Property Get LastErrorCode() As Long
    'Returns the last error that occured. 0 = no error
    LastErrorCode = m_lngLastErrorCode
End Property

Public Property Get LastErrorDescription() As String
    'Returns the last error description that occured. "" = no error
    LastErrorDescription = m_strLastErrorDesc
End Property

Public Property Let BlockSize(ByVal lngBytes As Long)
    'Round to next 54bit boundary
    m_lngEncBlockSize = ((lngBytes \ 54)) * 54 + IIf(IsOnBoundary(lngBytes, 54), 0, 54)
    'Round to next 584bit boundary
    m_lngDecBlockSize = ((lngBytes \ 584)) * 584 + IIf(IsOnBoundary(lngBytes, 584), 0, 584)
End Property

Public Sub ClearError()
    'Clear last error
    m_lngLastErrorCode = 0
    m_strLastErrorDesc = ""
End Sub

Private Function IsOnBoundary(lLen As Long, lBoundary As Long) As Boolean
    'Returns if a bytelength is exactly on a boundary of a chunk
    IsOnBoundary = lLen Mod lBoundary = 0
End Function

Private Sub RaiseError(ByVal lngCode As Long, ByVal strDescription As String)
    'Store the error and raise the event
    m_lngLastErrorCode = lngCode
    m_strLastErrorDesc = strDescription
    RaiseEvent ErrorOccured(lngCode, strDescription)
End Sub

Private Function fCreateObject(sObject As String) As Object
    'Tries to create an object (using createobject) but returns nothing when createobject fails
    On Error GoTo errCreateObject
    Set fCreateObject = CreateObject(sObject)
    On Error GoTo 0
    Exit Function

errCreateObject:
    Set fCreateObject = Nothing
End Function

Private Function EncDecFile(strInFile As String, strOutFile As String, bEncode As Boolean)
    'Encodes/decodes a file to/from base64
    Dim lBlockCount As Long
    Dim lSize As Long
    Dim lMaxBlocks As Long
    Dim lRest As Long
    Dim strTMP As String
    Dim FFIn As Integer
    Dim FFOut As Integer
    Dim lBlockSize As Long
    Dim strOut As String
    
    On Error GoTo errEncDecFile
    If bEncode Then lBlockSize = m_lngEncBlockSize Else lBlockSize = m_lngDecBlockSize
    
    m_bCancel = False
    lBlockCount = 0                 'Number of read blocks so far
    If Len(Dir(strInFile)) > 0 Then 'Does the file exist?
        RaiseEvent BeforeFileOpenIn(strInFile, m_bCancel)
        If m_bCancel Then Exit Function
        
        FFIn = FreeFile             'Get a free file handle
        Open strInFile For Binary As #FFIn
        
        RaiseEvent BeforeFileOpenOut(strOutFile, m_bCancel)
        If m_bCancel Then
            Close #FFIn 'Close already opened file handle
            Exit Function
        End If
        
        FFOut = FreeFile
        Open strOutFile For Binary As #FFOut
        
        lSize = LOF(FFIn)                'Get size of file
        lMaxBlocks = lSize \ lBlockSize  'Get number of blocks to read
        lRest = lSize - (lMaxBlocks * lBlockSize)        'Number of bytes left to read when all complete chunks have been read
        strTMP = String(lBlockSize, Chr(0))             'Initialize buffer
        While Not EOF(FFIn) And (lBlockCount < lMaxBlocks) And (Not m_bCancel) 'Read entire file
            Get #FFIn, , strTMP                               'Read a chunk
            If bEncode Then strOut = Me.Encode(strTMP) & IIf(lRest > 0, vbLf, "") Else strOut = Me.Decode(strTMP)
            Put #FFOut, , strOut
            lBlockCount = lBlockCount + 1                       'Count chunks
            RaiseEvent BlockRead(lBlockCount, lMaxBlocks + IIf(lRest > 0, 1, 0), IIf(bEncode, b64Encode, b64Decode), m_bCancel)
        Wend
        If (lRest > 0) And (Not EOF(FFIn)) And (Not m_bCancel) Then     'Read rest if required
            strTMP = String(lRest, Chr(0))                      'Initialize buffer
            Get #FFIn, , strTMP                                 'Read a chunk
            If bEncode Then strOut = Me.Encode(strTMP) Else strOut = Me.Decode(strTMP)
            Put #FFOut, , strOut
            RaiseEvent BlockRead(lBlockCount + 1, lMaxBlocks + 1, IIf(bEncode, b64Encode, b64Decode), m_bCancel)
        End If
        Close #FFOut
        RaiseEvent AfterFileCloseOut(strOutFile)
        
        Close #FFIn
        RaiseEvent AfterFileCloseIn(strInFile)
        If bEncode Then
            RaiseEvent FileEncodeComplete(strInFile, strOutFile)
        Else
            RaiseEvent FileDecodeComplete(strInFile, strOutFile)
        End If
        EncDecFile = True
    Else
        RaiseError 90002, "File does not exist / error opening file:" & strInFile
    End If
    Exit Function

errEncDecFile:
    RaiseError Err.Number, Err.Description
End Function
