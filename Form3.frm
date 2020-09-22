VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust  Data Packet Size"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   1191
      BandCount       =   2
      _CBWidth        =   7170
      _CBHeight       =   675
      _Version        =   "6.0.8450"
      Caption1        =   "Automatic"
      Child1          =   "Slider1"
      MinWidth1       =   495
      MinHeight1      =   615
      Width1          =   750
      NewRow1         =   0   'False
      Caption2        =   "Manual"
      Child2          =   "Text1"
      MinWidth2       =   600
      MinHeight2      =   285
      Width2          =   750
      NewRow2         =   0   'False
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6480
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   195
         Width           =   600
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   615
         Left            =   930
         TabIndex        =   1
         Top             =   30
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   1085
         _Version        =   393216
         LargeChange     =   1920
         SmallChange     =   128
         Min             =   128
         Max             =   12288
         SelStart        =   4096
         TickFrequency   =   128
         Value           =   4096
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Slider1.Value = LngChunkPackSize
Form3.Caption = Form3.Caption & "[ " & LngChunkPackSize & " ]"
Text1 = LngChunkPackSize
End Sub
Private Sub Slider1_Change()
LngChunkPackSize = Slider1.Value
Text1 = Slider1.Value
Form3.Caption = "Data Packet Size Set To :[ " & LngChunkPackSize & " ] Bytes"
End Sub

Private Sub Slider1_Click()
LngChunkPackSize = Slider1.Value
Text1 = Slider1.Value
Form3.Caption = "Data Packet Size Set To :[ " & LngChunkPackSize & " ] Bytes"
End Sub

Private Sub Text1_Change()
If Val(Text1) > Slider1.Min And Val(Text1) < Slider1.Max Then
LngChunkPackSize = Val(Text1.Text)
Else
End If
End Sub

Private Sub Text1_LostFocus()
If Val(Text1) > Slider1.Min And Val(Text1) < Slider1.Max Then
LngChunkPackSize = Val(Text1.Text)
Else
LngChunkPackSize = 4096
Slider1.Value = 4096
Text1 = 4096
End If
End Sub
