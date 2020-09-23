VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Base64 Test application"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Encode / Decode files"
      Height          =   4095
      Left            =   0
      TabIndex        =   14
      Top             =   3600
      Width           =   5535
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "c:\test.txt"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton cmdDecodeFile 
         Caption         =   "Decode"
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   735
         Width           =   1095
      End
      Begin VB.TextBox txtDecodeFile 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton cmdEncodeFile 
         Caption         =   "Encode"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   375
         Width           =   1095
      End
      Begin VB.TextBox txtEncodeFile 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
      Begin VB.ListBox lstLog 
         Height          =   2205
         ItemData        =   "frmTest.frx":0000
         Left            =   120
         List            =   "frmTest.frx":0002
         TabIndex        =   11
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Output to:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1125
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Result:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encode / Decode strings"
      Height          =   3495
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtResult 
         Height          =   1935
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox txtEncode 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdEncode 
         Caption         =   "Encode"
         Height          =   255
         Left            =   4320
         TabIndex        =   1
         Top             =   375
         Width           =   1095
      End
      Begin VB.TextBox txtDecode 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton cmdDecode 
         Caption         =   "Decode"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Result:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myBase64 As CBTBase64.Base64
Attribute myBase64.VB_VarHelpID = -1

'Encoding and decoding of strings
'================================
Private Sub cmdDecode_Click()
    txtResult.Text = myBase64.Decode(txtDecode.Text)
End Sub

Private Sub cmdEncode_Click()
    txtResult.Text = myBase64.Encode(txtEncode.Text)
End Sub

'Encoding and decoding of files
'==============================
Private Sub cmdEncodeFile_Click()
    Dim bRes As Boolean
    
    'Disable form to avoid starting something twice while busy ;-)
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    lstLog.Clear
    
    'This is where all the work is done...
    bRes = myBase64.EncodeFile(txtEncodeFile.Text, txtOutput.Text)
    
    'Present results and re-enable form
    lstLog.AddItem "Encoding of file '" & txtEncodeFile.Text & "' " & IIf(bRes, "succeeded", "failed")
    Me.MousePointer = vbNormal
    Me.Enabled = True
End Sub

Private Sub cmdDecodeFile_Click()
    Dim bRes As Boolean
    
    'Disable form to avoid starting something twice while busy ;-)
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    lstLog.Clear
    
    'This is where all the work is done...
    bRes = myBase64.DecodeFile(txtDecodeFile.Text, txtOutput.Text)
    
    'Present results and re-enable form
    lstLog.AddItem "Decoding of file '" & txtDecodeFile.Text & "' " & IIf(bRes, "succeeded", "failed")
    Me.MousePointer = vbNormal
    Me.Enabled = True
End Sub

'Initialize the base64 object and release it when the form terminates...
'=======================================================================
Private Sub Form_Initialize()
    Set myBase64 = New CBTBase64.Base64
    myBase64.BlockSize = 4 * 1048576 'Set blocksize to approx. 4Mb blocks
End Sub

Private Sub Form_Terminate()
    Set myBase64 = Nothing
End Sub

'Catch all events and display what's going on...
'===============================================
Private Sub myBase64_AfterFileCloseIn(ByVal strFileName As String)
    lstLog.AddItem "Closing file:" & strFileName
End Sub

Private Sub myBase64_AfterFileCloseOut(ByVal strFileName As String)
    lstLog.AddItem "Closing file:" & strFileName
End Sub

Private Sub myBase64_BeforeFileOpenIn(ByVal strFileName As String, bCancel As Boolean)
    lstLog.AddItem "Opening file:" & strFileName
End Sub

Private Sub myBase64_BeforeFileOpenOut(ByVal strFileName As String, bCancel As Boolean)
    lstLog.AddItem "Opening file:" & strFileName
End Sub

Private Sub myBase64_BlockRead(ByVal lngCurrentBlock As Long, ByVal lngTotalBlocks As Long, ByVal lBlockMode As CBTBase64.enBlockMode, bCancel As Boolean)
    Dim sPerc As Single
    sPerc = (lngCurrentBlock / lngTotalBlocks) * 100
    lstLog.AddItem "Block " & lngCurrentBlock & " of " & lngTotalBlocks & " (" & Round(sPerc, 2) & "%) " & IIf(lBlockMode = b64Encode, "encoding", "decoding")
    lstLog.ListIndex = lstLog.ListCount - 1
    lstLog.Refresh
End Sub

Private Sub myBase64_ErrorOccured(ByVal lngCode As Long, ByVal strDescription As String)
    lstLog.AddItem "Error!!! Code: " & lngCode & ", description: " & strDescription
End Sub

Private Sub myBase64_FileEncodeComplete(strOriginalFile As String, strEncodedFile As String)
    lstLog.AddItem "Completed encoding " & strOriginalFile & " to " & strEncodedFile
End Sub

Private Sub myBase64_FileDecodeComplete(strOriginalFile As String, strDecodedFile As String)
    lstLog.AddItem "Completed decoding " & strOriginalFile & " to " & strDecodedFile
End Sub
