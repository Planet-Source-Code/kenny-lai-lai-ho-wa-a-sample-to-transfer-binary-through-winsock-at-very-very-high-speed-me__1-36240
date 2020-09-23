VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Receiver"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.TextBox txtLocal 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Text            =   "Local IP: "
      Top             =   120
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock wsReceiver 
      Index           =   0
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrDownloadSpeed 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   240
   End
   Begin MSWinsockLib.Winsock wsInfo 
      Left            =   120
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save to File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      Height          =   270
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Remote IP: "
      Top             =   480
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPackageCount 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Package Count: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblSpeed 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Download Speed: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblTotalSize 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Total File Size: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblFileSize 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Byte Received: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TotalByte() As Byte

Dim CurrentFileSize As Long
Dim FileSize As Long

Dim DownloadSecond As Long
Dim DownloadSpeed As Long

Dim ByteNow As Long
Dim TotalByteNow As Long

Dim PackageCount As Long

Dim Package() As New Collection
Dim EntireFile As New Collection

Dim PackageByteNow() As Long
Dim PackageSize() As Long

Dim ReceiveComplete As Boolean

Private Sub cmdSave_Click()
With cd1
    .CancelError = True
    .Filter = "*.*|*.*"
    .Flags = cdlOFNOverwritePrompt
    On Error GoTo OpenError
    .ShowSave
    
    SaveBinaryArray .Filename, TotalByte
    
End With
Exit Sub
OpenError:

End Sub

Private Sub Form_Load()

txtLocal.Text = "Local IP: " & wsInfo.LocalIP

With wsInfo
    .LocalPort = 1700
    .Listen
End With

With wsReceiver(0)
    .LocalPort = 2000
    .Listen
End With

End Sub

Sub SetCaption(txt As String)
Me.lblStatus.Caption = txt
End Sub

Private Sub tmrDownloadSpeed_Timer()

DownloadSpeed = TotalByteNow - DownloadSecond

lblSpeed.Caption = "Download Speed: " & Round((DownloadSpeed / 1024) * 2, 3) & "KB/sec"

DownloadSecond = TotalByteNow

End Sub

Private Sub wsInfo_ConnectionRequest(ByVal requestID As Long)
wsInfo.Close
wsInfo.Accept requestID
End Sub

Private Sub wsInfo_DataArrival(ByVal bytesTotal As Long)
Dim i As Long
Dim a As String
wsInfo.GetData a

Select Case Left(a, 3)
    
    Case "FSC"
        ReceiveComplete = True
        'cmdSave.Enabled = True
        'tmrDownloadSpeed.Enabled = False
        'lblSpeed.Caption = "Completed"
        'FileSize = 0
        'pb.Value = 0
        
        'RebuildFile
        
        'With wsBinary
        '    .Close
        '    .LocalPort = 1800
        '    .Listen
        '    SetCaption "Listening at port " & .LocalPort
        'End With
        
        MsgBox "File Receive Complete!", vbInformation
    
    'Case "PKF"
    '    wsReceiver(CInt(Mid(a, 4))).Tag = "LAST"
    
    Case "FIS"
        CurrentFileSize = CLng(Mid(a, 4, InStr(1, a, "|") - 4))
        PackageCount = CLng(Right(a, Len(a) - InStr(1, a, "|")))
        lblTotalSize.Caption = "Total File Size: " & CurrentFileSize
        lblPackageCount.Caption = "Package Count: " & PackageCount
        
        ReDim PackageByteNow(1 To PackageCount) As Long
        ReDim PackageSize(1 To PackageCount) As Long
        ReDim Package(1 To PackageCount) As New Collection
        
        'MsgBox PackageCount, , "PackageCount"
        For i = 1 To PackageCount
            Load wsReceiver(i)
            wsReceiver(i).LocalPort = 2000 + i
            wsReceiver(i).Listen
            'MsgBox i & " is Listening at port " & wsReceiver(i).LocalPort
        Next i
        'MsgBox wsReceiver.UBound, , "wsReceiver.UBound"

End Select

End Sub

Private Sub wsReceiver_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'MsgBox Index, , "Connection Request"
wsReceiver(Index).Close
wsReceiver(Index).Accept requestID

If Index = 0 Then
    txtIP.Text = "Remote IP: " & wsReceiver(0).RemoteHostIP
End If

End Sub

Private Sub wsReceiver_DataArrival(Index As Integer, ByVal bytesTotal As Long)

If Index = 0 Then Exit Sub

Dim DataIn() As Byte

ReDim DataIn(1 To bytesTotal) As Byte
wsReceiver(Index).GetData DataIn()

Dim i As Long

tmrDownloadSpeed.Enabled = True

DoEvents
PackageByteNow(Index) = PackageSize(Index)
PackageSize(Index) = PackageSize(Index) + bytesTotal

'Set Progress Bar Value
On Error GoTo IgnorePB
pb.Max = CurrentFileSize
pb.Min = 0

ByteNow = 0

For i = LBound(PackageSize) To UBound(PackageSize)
    ByteNow = ByteNow + PackageSize(i)
Next
TotalByteNow = ByteNow

pb.Value = TotalByteNow

IgnorePB:

SetCaption "Receiving file from " & wsReceiver(Index).RemoteHost
    
    Package(Index).Add DataIn()
    
    On Error GoTo RemoveError
    EntireFile.Remove (Trim(Str(Index)))
RemoveError:
    EntireFile.Add Package(Index), Trim(Str(Index))

lblFileSize.Caption = "Byte Received: " & TotalByteNow

If TotalByteNow >= CurrentFileSize Or ReceiveComplete = True Then
        cmdSave.Enabled = True
        tmrDownloadSpeed.Enabled = False
        lblSpeed.Caption = "Completed"
        FileSize = 0
        pb.Value = 0
        
        RebuildFile
        
        DoEvents
        wsInfo.SendData "RFC"
        
        DoEvents
        MsgBox "Receive file complete.", vbInformation
End If

End Sub

Sub RebuildFile()

Dim d As Long

Dim i As Long, j As Long, k As Long

Dim WS As Long
Dim InChunk() As Byte
Dim InPackage As New Collection

ReDim TotalByte(1 To CurrentFileSize) As Byte

Dim WriteByteNow As Long
WriteByteNow = 0

For i = 1 To EntireFile.Count
    
    Set InPackage = EntireFile(Trim(Str(i)))
    
    For j = 1 To InPackage.Count
    
        InChunk = InPackage(j)
            For k = 1 To (UBound(InChunk) - LBound(InChunk) + 1)
                TotalByte(WriteByteNow + k) = InChunk(k - 1)
            Next k
            
            WriteByteNow = WriteByteNow + (UBound(InChunk) - LBound(InChunk) + 1)
    Next j
    
Next i

End Sub
