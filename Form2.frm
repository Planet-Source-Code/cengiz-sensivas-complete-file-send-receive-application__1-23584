VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Receiver"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   2  'Align Bottom
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   1429
      BandCount       =   6
      _CBWidth        =   9315
      _CBHeight       =   810
      _Version        =   "6.0.8450"
      Child1          =   "Text1"
      MinHeight1      =   285
      Width1          =   2850
      NewRow1         =   0   'False
      Child2          =   "Text2"
      MinWidth2       =   1005
      MinHeight2      =   285
      Width2          =   1095
      NewRow2         =   0   'False
      Child3          =   "ProgressBar1"
      MinHeight3      =   255
      Width3          =   2775
      NewRow3         =   0   'False
      MinHeight4      =   360
      NewRow4         =   0   'False
      Child5          =   "StatusBar1"
      MinHeight5      =   345
      Width5          =   4500
      NewRow5         =   -1  'True
      MinHeight6      =   360
      NewRow6         =   0   'False
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   4275
         TabIndex        =   4
         Top             =   75
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   345
         Left            =   165
         TabIndex        =   3
         Top             =   420
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   4
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   8378
               MinWidth        =   8378
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3087
               MinWidth        =   3087
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               Object.Width           =   2205
               MinWidth        =   2205
               TextSave        =   "14:50"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               Object.Width           =   2205
               MinWidth        =   2205
               TextSave        =   "30.05.2001"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         Height          =   285
         Left            =   165
         TabIndex        =   2
         Top             =   60
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000001&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3045
         TabIndex        =   1
         Top             =   60
         Width           =   1005
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   8760
      Top             =   2610
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalIncomingPacks As Long
Dim obj1 As FileSystemObject

Private Sub Form_Load()
Form2.Caption = "Enforma - File Receiver Waiting..."
    Socket(0).LocalPort = 1258
                        Socket(0).Listen
End Sub

Private Sub Socket_Close(Index As Integer)

Close LngDistributedFileNum
LngNextPack = 0
    If Socket(0).State <> sckConnected Then Socket(0).Close
        Pause (50)
            Socket(0).Close
                Pause (50)
                    Socket(0).LocalPort = 1257
                        Socket(0).Listen
                        Form1.Caption = "Enforma File Transfer -- [Ready]"

End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    If Socket(Index).State <> sckConnected Then Socket(Index).Close
        Socket(Index).Accept requestID
            Pause (50)
                If Socket(Index).State = sckConnected Then
                Socket(Index).SendData ("Autho")
                Form1.Caption = "Enforma File Transfer Connected To [" & Socket(Index).RemoteHostIP & "]"
                End If
                    

End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim StrIncomingData As String

'On Error GoTo Errorhandler
    
    Socket(Index).GetData StrIncomingData, vbString
    ControlString = Mid(StrIncomingData, 1, 9)

StatusBar1.Panels(2).Text = "Rcvd [" & ControlString & "]"
Select Case ControlString
    

'===Authentication Processing........======================================
'================================================================
    
    Case Is = "Authorize"
           
            StrKey = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
                If StrKey = "WelcomeToEnformaKioskServer" Then
                    Authorized = True
                If Socket(Index).State = sckConnected Then Socket(Index).SendData "WellC"
                    Pause (200)
                    Else
                        Authorized = False
                        If Socket(Index).State = sckConnected Then Socket(Index).SendData "RejeC"
                        Pause (200)
                        Socket(Index).Close
                End If
                Exit Sub
    


'===Distribute edilen Dosya Bilgisi....Dosya Adý & Lokasyonu ve Paketsayýsý...=======
'================================================================

    Case Is = "FileNamee"
            StrDistributedFileName = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
            Text1 = StrDistributedFileName
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "FileN"
    
    Case Is = "FileLengt"
            LngDistributedFileLength = Val(Mid(StrIncomingData, 11, Len(StrIncomingData) - 10))
            Text2 = LngDistributedFileLength
            ProgressBar1.Max = LngDistributedFileLength
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "FileL"
            
    Case Is = "TotalPack"
            LngTotalPacks = Val(Mid(StrIncomingData, 11, Len(StrIncomingData) - 10))
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "Total"
                
    Case Is = "BackUpNam"
            StrDistributedFileBackUpName = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "BackU"

    Case Is = "SaveFileT"
            StrDistributedFileSavingLocation = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "SaveF"

    Case Is = "TargetNam"
            StrDistributedFileTargetLocation = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "Targe"
            
    
    Case Is = "StartTran"
            LngDistributedFileNum = FreeFile()
            'StrDistributedFileSavingName = StrDistributedFileSavingLocation & StrDistributedFileName
            StrDistributedFileSavingName = "C:\Enforma System\Temporary Files\" & StrDistributedFileName
            StatusBar1.Panels(1).Text = "Saving As [" & StrDistributedFileSavingName & "]"
            Open StrDistributedFileSavingName For Binary As LngDistributedFileNum
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "FileO"
     
     Case Is = "FirsPacke"
            
            LngNextPack = 2
            StrIncomingData = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
            retval = InStr(1, StrIncomingData, "|")
            LngDistributedFileChunkPos = Val(Mid(StrIncomingData, 1, retval - 1))
            StrDistributedFileDataChunk = Mid(StrIncomingData, retval + 1, Len(StrIncomingData) - retval)
            retval = Len(StrDistributedFileDataChunk)
            TotalIncomingPacks = retval
            ProgressBar1.Value = TotalIncomingPacks
            Put LngDistributedFileNum, LngDistributedFileChunkPos, StrDistributedFileDataChunk
            If Socket(Index).State = sckConnected Then Socket(Index).SendData ("SndNx" & "|" & LngNextPack)
            
    Case Is = "DataPacke"
            LngNextPack = LngNextPack + 1
            StrIncomingData = Mid(StrIncomingData, 11, Len(StrIncomingData) - 10)
            retval = InStr(1, StrIncomingData, "|")
            LngDistributedFileChunkPos = Val(Mid(StrIncomingData, 1, retval - 1))
            StrDistributedFileDataChunk = Mid(StrIncomingData, retval + 1, Len(StrIncomingData) - retval)
            retval = Len(StrDistributedFileDataChunk)
            TotalIncomingPacks = TotalIncomingPacks + retval
            ProgressBar1.Value = TotalIncomingPacks
            Put LngDistributedFileNum, LngDistributedFileChunkPos, StrDistributedFileDataChunk
            StatusBar1.Panels(2).Text = "Rcvd [" & ControlString & "]" & "[" & LngNextPack - 1 & "]"
            If Socket(Index).State = sckConnected Then Socket(Index).SendData ("SndNx" & "|" & LngNextPack)
            
    
    Case Is = "CloseFile"
     
            Close LngDistributedFileNum
            LngNextPack = 0
            If Socket(Index).State = sckConnected Then Socket(Index).SendData "Close" & "|" & FileLen(StrDistributedFileSavingName)
            ProgressBar1.Value = 0
            StatusBar1.Panels(1).Text = "Transfer Completed"
            Text1 = ""
            Text2 = ""
            StatusBar1.Panels(2).Text = "Listening"
            StatusBar1.Panels(1).Text = "Ready"
            Form2.Caption = "Enforma - File Receiver Waiting..."
            Socket(Index).Close
                Socket(0).LocalPort = 1258
                        Socket(0).Listen
            
            
Case Else



End Select


Exit Sub
    
Errorhandler:
If sckServer.State = sckConnected Then sckServer.SendData ("ErrNo" & Err.Number & " [" & Err.Description & " ]")
Err.Clear
Exit Sub

End Sub

Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Close LngDistributedFileNum
LngNextPack = 0
If Socket(0).State <> sckConnected Then Socket(0).Close
    Pause (50)
    Socket(0).Close
    Pause (50)
    Socket(0).LocalPort = 1257
    Socket(0).Listen
    Form1.Caption = "Enforma File Transfer -- [Ready]"

End Sub

