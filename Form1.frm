VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   Caption         =   "Enforma - Intelligent Solutions"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl3.CoolBar CoolBar4 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4800
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   661
      BandCount       =   2
      _CBWidth        =   12210
      _CBHeight       =   375
      _Version        =   "6.0.8450"
      Child1          =   "StatusBar1"
      MinWidth1       =   2295
      MinHeight1      =   315
      Width1          =   2295
      NewRow1         =   0   'False
      Caption2        =   "Prgs"
      Child2          =   "ProgressBar1"
      MinWidth2       =   750
      MinHeight2      =   255
      Width2          =   735
      NewRow2         =   0   'False
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   11370
         TabIndex        =   15
         Top             =   60
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   315
         Left            =   165
         TabIndex        =   14
         Top             =   30
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   11
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "File Name"
               TextSave        =   "File Name"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "File Size"
               TextSave        =   "File Size"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1764
               MinWidth        =   1764
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "Packets"
               TextSave        =   "Packets"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   882
               MinWidth        =   882
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1676
               MinWidth        =   1676
               Text            =   "Last Chunk"
               TextSave        =   "Last Chunk"
            EndProperty
            BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1764
               MinWidth        =   1764
            EndProperty
            BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "Status"
               TextSave        =   "Status"
            EndProperty
            BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   4410
               MinWidth        =   4410
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   1920
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   3315
      ScaleWidth      =   8940
      TabIndex        =   11
      Top             =   1080
      Width           =   9000
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here To Start"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7080
         TabIndex        =   12
         Top             =   3000
         Width           =   1815
      End
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   741
      BandCount       =   2
      _CBWidth        =   12210
      _CBHeight       =   420
      _Version        =   "6.0.8450"
      Child1          =   "Text1"
      MinHeight1      =   285
      Width1          =   1995
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1500
      NewRow2         =   0   'False
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   165
         TabIndex        =   10
         Text            =   "Remote Host IP"
         Top             =   60
         Width           =   1800
      End
      Begin ComCtl3.CoolBar CoolBar3 
         Height          =   420
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   741
         BandCount       =   2
         _CBWidth        =   1695
         _CBHeight       =   420
         _Version        =   "6.0.8450"
         MinHeight1      =   360
         Width1          =   2880
         NewRow1         =   0   'False
         MinHeight2      =   360
         NewRow2         =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BCDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C12E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C582
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C9D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CE2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   3  'Align Left
      Height          =   4380
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   7726
      BandCount       =   2
      Orientation     =   1
      _CBWidth        =   765
      _CBHeight       =   4380
      _Version        =   "6.0.8450"
      Child1          =   "Toolbar1"
      MinHeight1      =   705
      Width1          =   300
      NewRow1         =   0   'False
      MinHeight2      =   360
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   3900
         Left            =   30
         TabIndex        =   7
         Top             =   165
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   6879
         ButtonWidth     =   1111
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList2"
         DisabledImageList=   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Browse"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Send"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Adjust"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Clean"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Rcvr"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start Receiver"
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transmit"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   3720
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7646
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   12648447
      BackColor       =   8421376
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D28A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DB32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DF86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EB4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EF9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F3F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F846
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FC9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":103EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11496
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":118EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "File Name"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Full File Path"
      Height          =   255
      Left            =   9600
      TabIndex        =   2
      Top             =   7440
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================
'=====================FILE TRANSFER & RECEIVE=================
'Programmed By  :   Cengiz SENSIVAS / Istanbul -TURKEY
'Level                  :   Intermediate
'Date                   :   30 May 2001 [Completed]
'O/S                    :   Win9x / WinNt / Win2000
'Version               :    v.1.4 [Next version will be with RESUME TRANSFER Option]
'Contact               :    cengiz@enforma.com.tr
'Copyright            :     Feel Free to use in YOUR OWN applications
'                               Distributing whole or part of this code is NOT allowed
'=============================================================
'=============================================================


Dim obj1 As FileSystemObject
Private Sub Command1_Click()
On Error Resume Next
If BrokenFileName = "" Or ResumeFileName = "" Then
retval = MsgBox("Please Select A File First", vbCritical, "Enforma")
Exit Sub
End If
    
    Socket(Index).RemoteHost = Text1.Text
    Socket(Index).RemotePort = 1258
    Socket(Index).Connect
    StatusBar1.Panels(11) = "Attempting to Connect to [" & Text1.Text & "] ..."
End Sub
Private Sub Command2_Click()
CommonDialog1.ShowOpen
Label1 = "Full File Path [" & CommonDialog1.FileName & "]"
Label2 = "Full File Path [" & CommonDialog1.FileTitle & "]"
BrokenFileName = CommonDialog1.FileName
ResumeFileName = CommonDialog1.FileTitle
strfilename = CommonDialog1.FileName
DividePacks strfilename, 4096
End Sub

Private Sub Command3_Click()
Form2.Show , Form1
End Sub

Private Sub Form_Load()
ListView1.ColumnHeaders.Add , , "Packets", (ListView1.Width / 12) * 4, lvwColumnLeft
ListView1.ColumnHeaders.Add , , "Length", (ListView1.Width / 12) * 4, lvwColumnLeft
ListView1.ColumnHeaders.Add , , "Position", (ListView1.Width / 12) * 4, lvwColumnLeft
ListView1.View = lvwReport
LngChunkPackSize = 4096
Text1 = Socket(Index).LocalIP
BrokenFileName = ""
ResumeFileName = ""
Form2.Show , Form1
End Sub

Private Sub Label3_Click()
Picture1.Visible = False
Label3.Visible = False
End Sub

Private Sub Picture1_Click()
Pause (3000)
Picture1.Visible = False
End Sub

Private Sub Socket_Close(Index As Integer)
Socket(Index).Close
                Form1.Caption = "Disconnected"
                Form1.ListView1.ListItems.Clear
                Erase StrDataChunk
                Erase LngPosition
                StatusBar1.Panels(2).Text = ""
                StatusBar1.Panels(4).Text = ""
                StatusBar1.Panels(6).Text = ""
                StatusBar1.Panels(8).Text = ""
                StatusBar1.Panels(10).Text = ""
                StatusBar1.Panels(10).Text = "[Transfer Completed]"
                ProgressBar1.Value = 0
                Pause (2000)
                StatusBar1.Panels(10).Text = "[Ready]"
                Form1.Caption = "Ready"
                
End Sub
Private Sub Socket_Connect(Index As Integer)
    Form1.Caption = "Connected"
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Errorhandler
Dim VtData As String
If Socket(Index).State = sckConnected Then Socket(Index).GetData VtData, vbString

Pause (50)
SelectionPart = Mid(VtData, 1, 5)

StatusBar1.Panels(2).Text = SelectionPart
Select Case SelectionPart


'================Distribution  Errorr=======================
'====================================================

Case Is = "ErrNo"
        ErrorNumara = Mid(VtData, 6, Len(VtData) - 5)
'================Autho  Authentication isteniyor==============
'====================================================

Case "Autho"
    Pause (200)
    Socket(Index).SendData ("Authorize" & "|" & "WelcomeToEnformaKioskServer")
    VtData = ""
    
    Exit Sub

'================RejeC Connection rejected================
'====================================================
Case "RejeC"
    Pause (200)
    VtData = ""
    Exit Sub
    

'================WellC Connection accepted===============
'====================================================
Case "WellC"
 
    Pause (200)
    VtData = ""
    Set obj1 = New FileSystemObject
    '====Þimdi Distribute edilecek dosya ile ilgili bilgileri yollayalým
    '====Lets send some file information regarding the file we will transfer
        '===Önce Dosya Adý
        '===First File Name
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("FileNamee" & "|" & obj1.GetFileName(strfilename))
                Pause (200)

Case Is = "FileN"
                Pause (200)
                '===Dosya Boyutu
                '===File Length
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("FileLengt" & "|" & FileLen(strfilename))
                Pause (200)
                
Case Is = "FileL"
                Pause (200)
                '===Sonra Kaç Paket Olduðu
                '===Then How Many Packs The File is Divided Into
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("TotalPack" & "|" & LngChunkPacks)
                Pause (200)
                        
                '===These Commands will be used to Resume Broken Transfers and Move the Transferred Files to Original Locations....
                '===These options are not applicable for the version prior to v 1.7
                
Case Is = "Total"
                Pause (200)
                '===Þimdi Back Up edilecek ismi....
                '===Now Back Up File Name
                With obj1
                    retFNM = .GetFileName(strfilename)
                    retEXT = .GetExtensionName(strfilename)
                    retINS = InStr(1, retFNM, retEXT)
                    retMID = Mid(retFNM, 1, retINS - 2)
                    strBackUpName = retMID & "_old." & retEXT
                End With
        
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("BackUpNam" & "|" & strBackUpName)
                Pause (200)
        
Case Is = "BackU"
                Pause (200)
                '===Þimdi Nereye Yazýlacaðý..........
                '===Now into where to store the File
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("SaveFileT" & "|" & StrSaveLocation)
                Pause (200)
        
Case Is = "SaveF"
                Pause (200)
                '===Þimdi En Son olarak Hangi Dosya ile Replace edileceði....
                '===Finally With which file to be replaced.........
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("TargetNam" & "|" & strfilename)
                Pause (200)
    
Case Is = "Targe"
                Pause (200)
                '===Þimdi START TRANSFER KOMUTU================
                '===OK Now we can start the transfer
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("StartTran")
                Pause (200)
        
Case Is = "FileO"
                '===File Opened on remote side
                Pause (200)
                '===Þimdi ÝLK PAKETÝ YOLLAYALIM==================
                '===Now lets send the first data packet and wait for the Send Next = SndNx command from the remote side
            
                If Socket(Index).State = sckConnected Then Socket(Index).SendData ("FirsPacke" & "|" & LngPosition(1) & "|" & StrDataChunk(1))
                ProgressBar1.Value = ProgressBar1.Value + Len(StrDataChunk(1))
                Pause (200)
                
Case Is = "SndNx"
                '===Remote side wants us to send the next packet
                '=====================================
                ReceivedPacketNo = Mid(VtData, 7, Len(VtData) - 6)
                StatusBar1.Panels(10).Text = "Sending Packet No : [" & ReceivedPacketNo & "]"
                Pause (200)
        
                '===Þimdi PAKETLERÝ ALDIKÇA GÖNDERELÝM===================
                '===Lets Keep Transferring===================================
              
                If ReceivedPacketNo <= UBound(StrDataChunk) Then
                    If Socket(Index).State = sckConnected Then Socket(Index).SendData ("DataPacke" & "|" & LngPosition(ReceivedPacketNo) & "|" & StrDataChunk(ReceivedPacketNo))
                    ProgressBar1.Value = ProgressBar1.Value + Len(StrDataChunk(ReceivedPacketNo))
                    Pause (200)
                Else
                    Pause (400)
                        '====At last we can tell the Remote Side to Close the file we transferred....
                        If Socket(Index).State = sckConnected Then Socket(Index).SendData ("CloseFile")
                End If

Case Is = "Close"
                '===Remote Side confirms that the file is transferred and closed
                
                RcvdFileLen = Mid(VtData, 7, Len(VtData) - 6)
                Pause (200)
                Socket(Index).Close
                '===OK ..... Dosya Gitti ve "Enforma System \ Temporary Files\" altýna yazýldý=============================
                '===OK....File is transferred succesfully and saved.

End Select

Exit Sub

Errorhandler:

StatusBar1.Panels(10).Text = "[" & Err.Number & "]" & "[" & Err.Description & "]" & "Socket[" & Index & "]"
Err.Clear
Socket(Index).Close
Exit Sub

End Sub

Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Socket(Index).Close
    StatusBar1.Panels(11) = "Err: [" & Text1.Text & "] " & "[ " & Number & "] [" & Description & "]"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1

    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub
    Label1 = "Full File Path [" & CommonDialog1.FileName & "]"
    Label2 = "Full File Path [" & CommonDialog1.FileTitle & "]"
    BrokenFileName = CommonDialog1.FileName
    ResumeFileName = CommonDialog1.FileTitle
    strfilename = CommonDialog1.FileName
    '===Lets Divide The File into Data Packets
    '===See Public Function(DividePacks) in Module1
    DividePacks strfilename, LngChunkPackSize
    ProgressBar1.Max = FileLen(BrokenFileName)
    
Case 2
    On Error Resume Next
    If BrokenFileName = "" Or ResumeFileName = "" Then
    retval = MsgBox("Please Select A File First", vbCritical, "Enforma")
    Exit Sub
    End If
    If Text1 = "Remote Host IP" Then
        retval = MsgBox("Enter Remote Host IP", vbCritical, "Enforma")
        Exit Sub
    Else
        Socket(Index).RemoteHost = Text1.Text
        Socket(Index).RemotePort = 1258
        Socket(Index).Connect
        StatusBar1.Panels(10) = "Attempting to Connect to [" & Text1.Text & "] ..."
    End If
Case 3
Form3.Show , Form1
Case 4
        Socket(Index).Close
        Form1.ListView1.ListItems.Clear
        Erase StrDataChunk
        Erase LngPosition
        StatusBar1.Panels(2).Text = ""
        StatusBar1.Panels(4).Text = ""
        StatusBar1.Panels(6).Text = ""
        StatusBar1.Panels(8).Text = ""
        StatusBar1.Panels(10).Text = ""
        
Case 5
Form2.Show , Form1
End Select
End Sub
