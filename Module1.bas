Attribute VB_Name = "Module1"
Declare Function GetTickCount Lib "kernel32" () As Long
Public LngChunkPackSize As Long
Public LngPosition() As Long
Public StrDataChunk() As String
Public Authorizationstring As String
Public TaskIDToCheck As String
Public TimeToCheck As String
Public TaskTimeToCheck As String
Public StrKey As String
Public ControlString1 As String
Public BrokenFileName As String
Public ResumeFileName As String
Public BrokenFileLen As Long
Public ResumeFileLen As Long
Public ResumeFilePacket As String * 512
Public ResumeFilePacketPos As Long
Public BrokenFileNum As Long
Public ResRetval As Long
Public strfilename As String
Public OptionsIP As String
Public OptionsTop As Long
Public OptionsLeft As Long
Public LngNextPack As Long
Public LngTotalPacks As Long
Public StrDistributedFileName As String
Public LngDistributedFileLength As Long
Public LngDistributedFileNum As Long
Public LngDistributedFileChunkPos As Long
Public StrDistributedFileDataChunk As String
Public StrDistributedFileBackUpName As String
Public StrDistributedFileSavingLocation As String
Public StrDistributedFileTargetLocation As String
Public StrDistributedFileSavingName As String
Public ScreenFileNum As Long
Public SendScreen As Boolean
Public ResumeTargetLocation As String
Public ResumeTargetFolder As String
Public ResumeFileName2 As String
Public KioskAppHwnd As Long
Public SckStatusRemoteIP As String
Public SckStatusRemotePort As String
Public GelenDosya As String
Public MAX_CHUNK As Long
Public bReplied As Boolean
Public lTIme As Long
Public LastRemoteIP As String

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub
Public Function DividePacks(ByVal FileToDivide As String, ByVal PacketSize As Long)

On Error GoTo Errorhandler

strfilename = FileToDivide
    Form1.StatusBar1.Panels(2).Text = strfilename
    LngSendFileNum = FreeFile()

'===Clean Up
'=========
        Form1.ListView1.ListItems.Clear
        Erase StrDataChunk
        Erase LngPosition

'===Lets get the last chunk size.....
'===Dosya Boyuna göre Toplam Packet adedini ve Son Paket uzunluðunu

        LngFileSize = (FileLen(strfilename))
        Form1.StatusBar1.Panels(4).Text = LngFileSize
        
        LngPacketSize = PacketSize
        
        If LngFileSize < LngPacketSize Then
        
                    LngChunkPacks = 1
                    ReDim StrDataChunk(LngChunkPacks)
                    ReDim LngPosition(LngChunkPacks)
        Else
        
        LngLastChunkSize = LngFileSize Mod LngPacketSize
        Form1.StatusBar1.Panels(8).Text = LngLastChunkSize
        
'===Bu Bölüm kontrol açýsýndan öneli...Son Packet size ile
'===Toplam Packet sayýsý açýsýndan....
'===This section is to check the total len of the file and the total len of the chunks
        
        
        If LngLastChunkSize = 0 Then
                    LngChunkPacks = LngFileSize / LngPacketSize
                    ReDim StrDataChunk(LngChunkPacks)
                    ReDim LngPosition(LngChunkPacks)
                    Form1.StatusBar1.Panels(6).Text = LngChunkPacks
                Else
                    StrChunkPacks = LngFileSize / LngPacketSize
                    StrChunkPacks = Left(StrChunkPacks, (InStr(1, StrChunkPacks, ",")) - 1)
                    LngChunkPacks = Val(StrChunkPacks) + 1
                    ReDim StrDataChunk(LngChunkPacks)
                    ReDim LngPosition(LngChunkPacks)
                    Form1.StatusBar1.Panels(6).Text = LngChunkPacks
        End If
        
        End If
        
        
        '===Þimdi Dosyayý Packetler halinde okuyalým ve Array içersine alalým
        '===Now lets get the data packets into a dynamic array
        '__Binary Mode'da açýyoruz...
        '==Opening in BINARY MODE
        
        Open strfilename For Binary As LngSendFileNum
    
        '__Input ve Get methodlarýnýn her ikisi de kullanýlabilir...fakat Input Methodu ile
        '__Chunk Size set edilmesi daha kolay...Get Methodu ise Position By Position read/write
        '__Açýsýndan faydalý....Önce Input Methodu Deneyelim...
        
        Position = 1
        For lngIndex = 1 To LngChunkPacks
        
'==BU Packetleri Dinamik StrDataCHunk Array'ýna yazar
'==Writing Data Packets into Dynamic StrDataChunk ARRAY
        
            '==Get Data Chunk
            StrDataChunk(lngIndex) = Input(LngPacketSize, LngSendFileNum)
                '==Get Position and store
                LngPosition(lngIndex) = Position
                '==Increase Position counter automatically
                    Position = Position + LngPacketSize
            
            '__Burada Loc() ve Seek() Methodlarý ile son okunan Byte Position ve Next Byte To read
            '__Pozisyonlarýný alabiliriz...Bu positionlar Socket ile transfer esnasýnda remote tarafta
            '__Dosya yazýmýnda kullanýþlý olacaktýr....
            '__Lets get the Loc to tell the remote side where to write the Data Chunk in the file position
            
            LngLocPos = Loc(LngSendFileNum)
            LngSeekPos = Seek(LngSendFileNum)
            
            Set ItmX = Form1.ListView1.ListItems.Add(Form1.ListView1.ListItems.Count + 1, , ("Packet [" & lngIndex & "]"), 2, 2)
            ItmX.SubItems(1) = Len(StrDataChunk(lngIndex))
            ItmX.SubItems(2) = LngPosition(lngIndex)
    
            
        '==BU Packetleri Temporary folder altýnda textfile olarak yazar....
        '==Att.önce sil sonra yaz...!!!
        '==You can store the data chunks into a file for future use....
        
                    'strReadPacket = Input(4196, SendFileNum)
                    'Set Obj1 = CreateObject("Scripting.FileSystemObject")
                    'StrPacketName = "Packet[" & LngIndex & "]"
                    'Set a = Obj1.CreateTextFile("c:\windows\desktop\Temporary\" & StrPacketName, True)
                    'a.WriteLine strReadPacket
                    'a.Close
    
            

Next lngIndex
            
Close

If LngFileSize < LngPacketSize Then
    TestRetval = Len(StrDataChunk(1))
    Form1.StatusBar1.Panels(10).Text = "Total Packets [" & LngChunkPacks & "] File Size Check [" & TestRetval & "]"
        If TestRetval = LngFileSize Then
        Pause (200)
        Form1.StatusBar1.Panels(10).Text = "...Success"
        Else
        Pause (200)
        Form1.StatusBar1.Panels(10).Text = "...Error"
        Exit Function
        End If
Else
    TestRetval = ((LngChunkPacks - 1) * LngPacketSize) + LngLastChunkSize
    Form1.StatusBar1.Panels(10).Text = "Total Packets [" & LngChunkPacks & "] File Size Check [" & TestRetval & "]"
        If TestRetval = LngFileSize Then
        Pause (200)
        Form1.StatusBar1.Panels(10).Text = "...Success"
        Else
        Pause (200)
        Form1.StatusBar1.Panels(10).Text = "...Error"
        Exit Function
        End If
End If

Pause (200)
Form1.StatusBar1.Panels(10).Text = "...Ready"


Exit Function


Errorhandler:
Form1.StatusBar1.Panels(10).Text = " [1] Error :[" & Err.Number & "] Desc:[" & Err.Description & "]"
Err.Clear
Exit Function
End Function
