Attribute VB_Name = "modBinaryFile"
Sub SaveBinaryArray(ByVal Filename As String, WriteData() As Byte)

    Dim t As Integer
    t = FreeFile
    Open Filename For Binary Access Write As #t
        
            Put #t, , WriteData()
        
    Close #t
    
End Sub

Function ReadBinaryArray(ByVal Source As String)

    Dim bytBuf() As Byte
    Dim intN As Long
    
    Dim t As Integer
    t = FreeFile
    
    Open Source For Binary Access Read As #t
    
    Dim n As Long
    
    ReDim bytBuf(1 To LOF(t)) As Byte
    Get #t, , bytBuf()
    
    ReadBinaryArray = bytBuf()
    
    Close #t
    
End Function
