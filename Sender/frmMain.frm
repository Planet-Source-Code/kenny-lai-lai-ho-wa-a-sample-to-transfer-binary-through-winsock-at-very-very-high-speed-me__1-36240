VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Sender"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
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
   ScaleHeight     =   4875
   ScaleWidth      =   4785
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command1 
      Caption         =   "Send Finished Commend"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSendNow 
      Caption         =   "&Send the File Now"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtPackage 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "1024"
      Top             =   1680
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock wsSender 
      Index           =   0
      Left            =   1440
      Tag             =   "0"
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsInfo 
      Left            =   960
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChunk 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "4096"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2160
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  '¥­­±
      Height          =   270
      Left            =   600
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Open a file to send"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1920
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Chunk/Package: "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Chunk Size: "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblFileSize 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "File Size: "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentFileSize As Long

Dim ChunkByte() As Byte
Dim ChunkSize As Long

Dim Chunks As New Collection
Dim Package As Collection
Dim EntireFile As New Collection

Dim PackageSize As Long

Dim ThePackageSize() As Long
Dim ChunkSent() As Long

Dim ByteNow As Long
Dim ChunkNow As Long
Dim PackageNow As Long

Dim SendComplete As Boolean
Dim PackageComplete() As Boolean

Dim RemoteHost As String

Dim TotalByteSent As Long

Private Sub cmdSend_Click()
PrepareFile

RemoteHost = txtIP.Text

With wsSender(0)
    .Close
    .RemoteHost = RemoteHost
    .RemotePort = 2000
    .Connect
End With
End Sub

Private Sub cmdSendNow_Click()

'With wsBinary
'    .Close
'    .RemoteHost = txtIP.Text
'    .RemotePort = 1800
'    .Connect
'End With
Dim i As Long

    For i = 1 To PackageNow
        Load wsSender(i)
        With wsSender(i)
            .Close
            .RemoteHost = RemoteHost
            .RemotePort = 2000 + i
            DoEvents
            .Connect
        End With
    Next i
        
End Sub

Sub SetCaption(txt As String)
Me.lblStatus.Caption = txt
End Sub

Private Sub Command1_Click()
wsInfo.SendData "FSC"
End Sub

Private Sub wsInfo_Connect()

pb.Max = CurrentFileSize
pb.Min = 0
pb.Value = 0

wsInfo.SendData "FIS" & CurrentFileSize & "|" & PackageNow
End Sub

Private Sub wsInfo_DataArrival(ByVal bytesTotal As Long)

'MsgBox "wsInfo_DataArrival"

Dim a As String
wsInfo.GetData a

Select Case Left(a, 3)

    Case "RFC"
        MsgBox "Whole file sending complete.", vbInformation
        pb.Value = 0
    
End Select

End Sub

Private Sub wsSender_Connect(Index As Integer)

Dim i As Long

If Index = 0 Then
        
        With wsInfo
            .Close
            .RemoteHost = RemoteHost
            .RemotePort = 1700
            DoEvents
            .Connect
        End With
        
        'MsgBox "Ready to connect"
        
Else
    'The senders are connected
    'MsgBox Index & " is connected to port " & wsSender(Index).RemotePort
    
    Dim ThisPackage As Collection
    Set ThisPackage = EntireFile(Index)
    
    Dim ThisByte() As Byte
    
    ThePackageSize(Index) = ThisPackage.Count

    For i = 1 To ThisPackage.Count
        ThisByte = ThisPackage(i)
        'MsgBox "This Byte is " & UBound(ThisByte)
        DoEvents
        wsSender(Index).SendData ThisByte
        pb.Value = pb.Value + UBound(ThisByte)
    Next i
    'MsgBox "wsSender(" & Index & ") Send Data Command Executed"
    
    'SendFinishedMessage
    
End If

Exit Sub
OpenError:

End Sub

Sub PrepareFile()

Dim i As Long
    'The main gate is connected
        With cd1
            
            Dim Source As String
            
            .Filter = "*.*|*.*"
            .CancelError = True
            .Flags = cdlOFNFileMustExist
            
            On Error GoTo OpenError
            .ShowOpen
            
            Source = .Filename
        End With
        
        CurrentFileSize = FileLen(Source)
        ChunkSize = CLng(txtChunk.Text)
        
        PackageSize = CLng(txtPackage.Text)

'****************************************************
'Check file correct
'cd1.ShowOpen
'Dim ExactBytes() As Byte
'ExactBytes = ReadBinaryArray(cd1.Filename)
'*****************************************************
        
        '--------------------------------------------------
        'Create groups of chunks of byteArray
        '--------------------------------------------------
        
        'Read the entire file into an array
        Dim DataOut() As Byte
        DataOut = ReadBinaryArray(Source)
        
        ByteNow = 0
        ChunkNow = 0
        PackageNow = 0
        
                '-----------------------------------------------------------------------------------------------
                'Load a part of bytes into all chunks
                Do Until (CurrentFileSize - ByteNow) <= ChunkSize 'Wait for the Last chunk
                    ChunkNow = ChunkNow + 1
                    
                    ReDim ChunkByte(1 To ChunkSize)
                        For i = 1 To ChunkSize
                            ChunkByte(i) = DataOut(i + ByteNow)
                        Next
                    
                    'Test the file
                    'If ChunkNow = 1 Then
                    '    For i = 1 To 128
                    '        Debug.Print ExactBytes(i), ChunkByte(i)
                    '    Next i
                    'End If
                        
                    Chunks.Add ChunkByte(), Trim(Str(ChunkNow))
                    
                    ByteNow = ByteNow + ChunkSize
                Loop
                
                Dim LastChunkSize As Long 'Deal with the last chunk in a package
                LastChunkSize = CurrentFileSize - ByteNow
                    ChunkNow = ChunkNow + 1
                    'This is the last chunk of the entire file
                    ReDim ChunkByte(1 To LastChunkSize)
                    For i = 1 To LastChunkSize
                        ChunkByte(i) = DataOut(i + ByteNow)
                    Next i
                    
                    Chunks.Add ChunkByte(), Trim(Str(ChunkNow))
                    
                    ByteNow = ByteNow + LastChunkSize
                    '--------------------------------------------------------------------------------------------
                    
            'Load the chunks into packages
            ChunkNow = 0
            
            Do Until (Chunks.Count - ChunkNow) <= PackageSize 'Wait for the last package
            
            PackageNow = PackageNow + 1
            
                Set Package = New Collection
                
                For i = 1 To PackageSize
                    Package.Add Chunks(Trim(Str(i + ChunkNow)))
                Next i
                
                EntireFile.Add Package, Trim(Str(PackageNow))
                
                ChunkNow = ChunkNow + PackageSize
            Loop
            
            Dim LastPackageSize As Long
            LastPackageSize = Chunks.Count - ChunkNow
            
                PackageNow = PackageNow + 1
                
                Set Package = New Collection
                
                For i = 1 To LastPackageSize
                    Package.Add Chunks(Trim(Str(i + ChunkNow)))
                Next i
                
                EntireFile.Add Package, Trim(Str(PackageNow))
                
                ChunkNow = ChunkNow + LastPackageSize
                
        ReDim PackageComplete(1 To EntireFile.Count) As Boolean
        ReDim ThePackageSize(1 To EntireFile.Count) As Long
        ReDim ChunkSent(1 To EntireFile.Count) As Long
        
        lblFileSize.Caption = "File Size: " & CurrentFileSize
        'MsgBox "CurrentFileSize: " & CurrentFileSize
        'MsgBox "Chunks Collection Count: " & Chunks.Count
        'MsgBox "ChunkNow: " & ChunkNow
        'MsgBox "ByteNow: " & ByteNow
        'MsgBox "PackageNow: " & PackageNow
        'MsgBox "EntireFile Collection Count: " & EntireFile.Count
Exit Sub
OpenError:
End Sub

Private Sub wsSender_SendComplete(Index As Integer)

'Dim i As Long

'ChunkSent(Index) = ChunkSent(Index) + 1
'If ChunkSent(Index) = ThePackageSize(Index) Then
'    PackageComplete(Index) = True
'ElseIf ChunkSent(Index) = ThePackageSize(Index) - 1 Then
'    wsInfo.SendData "PKF" & Index
'End If

'For i = 1 To wsSender.UBound
'    If PackageComplete(i) = False Then Exit Sub
'Next i

End Sub

Private Sub SendFinishedMessage()
    
    DoEvents
    wsInfo.SendData "FSC"
    pb.Value = 0

End Sub
