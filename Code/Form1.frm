VERSION 5.00
Object = "{6D9F7F71-9658-11D0-BDB5-00608CC9F9FB}#1.0#0"; "MILSystem.ocx"
Object = "{45BC0BC3-A6C5-11D0-BDD1-00608CC9F9FB}#1.0#0"; "MILDisplay.ocx"
Object = "{03985961-6B33-11D0-AB4A-00608CC9CA57}#1.0#0"; "MilBuffer.ocx"
Object = "{F2E7BDE3-B006-11D0-9162-00A024D24992}#1.0#0"; "MILGraphContext.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10344
   ClientLeft      =   132
   ClientTop       =   480
   ClientWidth     =   15204
   LinkTopic       =   "Form1"
   ScaleHeight     =   862
   ScaleMode       =   3  '像素
   ScaleWidth      =   1267
   StartUpPosition =   3  '系統預設值
   Begin MILBUFFERLib.Buffer BufBinarize2 
      Height          =   480
      Left            =   5040
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   1
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MILBUFFERLib.Buffer BufChainCd 
      Height          =   480
      Left            =   4440
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   1
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MILBUFFERLib.Buffer BufBinarize 
      Height          =   480
      Left            =   3480
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   1
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MILBUFFERLib.Buffer BufRed 
      Height          =   480
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   1
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MILBUFFERLib.Buffer Buffer 
      Height          =   480
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   3
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MILSYSTEMLib.System System 
      Height          =   480
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      SystemType      =   "VGA"
      ProcessingSystem=   1634756
      ProcessingSystemName=   "[Default]"
   End
   Begin MILGRAPHICCONTEXTLib.GraphicContext GraphicContext 
      Height          =   480
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      Buffer          =   "Buffer"
      ForegroundShade =   255
      BackgroundShade =   0
      ForegroundColor =   16777215
      BackgroundColor =   0
      ForegroundColorMode=   -1  'True
      BackgroundColorMode=   -1  'True
   End
   Begin MILBUFFERLib.Buffer Buffer1 
      Height          =   480
      Left            =   3960
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   1
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MILBUFFERLib.Buffer BufHue 
      Height          =   480
      Left            =   5640
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   1
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "控制台"
      Height          =   6852
      Left            =   10320
      TabIndex        =   5
      Top             =   960
      Width           =   3612
      Begin VB.Frame Frame5 
         Caption         =   "鍊碼總步數"
         Height          =   732
         Left            =   360
         TabIndex        =   21
         Top             =   2880
         Width           =   2532
         Begin VB.Label Label2 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1812
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "凸包點數目"
         Height          =   732
         Left            =   360
         TabIndex        =   19
         Top             =   3840
         Width           =   2532
         Begin VB.Label Label2 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1812
         End
         Begin VB.Label Label2 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   12
            Index           =   4
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   1812
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "辨識結果"
         Height          =   732
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   2532
         Begin VB.Label Label2 
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   2172
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "當前動作"
         Height          =   732
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   2652
         Begin VB.Label Label2 
            Caption         =   "Processing"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1812
         End
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   492
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   492
         Left            =   120
         TabIndex        =   13
         Top             =   4920
         Width           =   972
      End
      Begin VB.CommandButton cmdBlobs 
         Caption         =   "Blobs"
         Height          =   492
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdUnload 
         Caption         =   "Exit"
         Height          =   492
         Left            =   2280
         TabIndex        =   7
         Top             =   4920
         Width           =   852
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   492
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '透明
         Caption         =   "凸包"
         Height          =   252
         Left            =   2640
         TabIndex        =   29
         Top             =   6000
         Width           =   372
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         Height          =   852
         Index           =   2
         Left            =   2280
         Shape           =   3  '圓形
         Top             =   5760
         Width           =   1092
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "八方鍊碼"
         Height          =   252
         Left            =   1320
         TabIndex        =   28
         Top             =   6000
         Width           =   732
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         Height          =   852
         Index           =   0
         Left            =   1080
         Shape           =   3  '圓形
         Top             =   5760
         Width           =   1212
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "辨識"
         Height          =   252
         Left            =   480
         TabIndex        =   27
         Top             =   6000
         Width           =   372
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         Height          =   852
         Index           =   1
         Left            =   120
         Shape           =   3  '圓形
         Top             =   5760
         Width           =   1092
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7812
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   13780
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "Org_buffer"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Display1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Binarize"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Display1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "凸包線段資訊"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Display1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "凸包點鏈碼序號"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Display1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Test"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Display1(4)"
      Tab(4).ControlCount=   1
      Begin MILDISPLAYLib.Display Display1 
         Height          =   7080
         Index           =   3
         Left            =   -74880
         TabIndex        =   24
         Top             =   600
         Width           =   9600
         _Version        =   65536
         _ExtentX        =   16933
         _ExtentY        =   12488
         _StockProps     =   1
         BackColor       =   0
         AutomaticAllocation=   -1  'True
         OwnerSystem     =   "System"
         DeviceNumber    =   3
         Buffer          =   "BufChainCd"
         LUT             =   ""
         OverlayLUT      =   ""
         ExternalWindowTitle=   "Display1(3) - ActiveMIL External Window"
         FormatArrayIndex=   -1
         BackColor       =   0
         OverlayKeyColor =   268435456
         FormatArrayListSize=   24
         DisplayVisible  =   -1  'True
      End
      Begin MILDISPLAYLib.Display Display1 
         Height          =   7080
         Index           =   2
         Left            =   -74880
         TabIndex        =   23
         Top             =   600
         Width           =   9600
         _Version        =   65536
         _ExtentX        =   16933
         _ExtentY        =   12488
         _StockProps     =   1
         BackColor       =   0
         AutomaticAllocation=   -1  'True
         OwnerSystem     =   "System"
         DeviceNumber    =   2
         Buffer          =   "Buffer1"
         LUT             =   ""
         OverlayLUT      =   ""
         ExternalWindowTitle=   "Display1(2) - ActiveMIL External Window"
         FormatArrayIndex=   -1
         BackColor       =   0
         OverlayKeyColor =   268435456
         FormatArrayListSize=   24
         DisplayVisible  =   -1  'True
      End
      Begin MILDISPLAYLib.Display Display1 
         Height          =   7080
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   9600
         _Version        =   65536
         _ExtentX        =   16933
         _ExtentY        =   12488
         _StockProps     =   1
         BackColor       =   0
         AutomaticAllocation=   -1  'True
         OwnerSystem     =   "System"
         DeviceNumber    =   1
         Buffer          =   "BufBinarize"
         LUT             =   ""
         OverlayLUT      =   ""
         ExternalWindowTitle=   "Display1(1) - ActiveMIL External Window"
         FormatArrayIndex=   -1
         BackColor       =   0
         OverlayKeyColor =   268435456
         FormatArrayListSize=   24
         DisplayVisible  =   -1  'True
      End
      Begin MILDISPLAYLib.Display Display1 
         Height          =   7080
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   600
         Width           =   9600
         _Version        =   65536
         _ExtentX        =   16933
         _ExtentY        =   12488
         _StockProps     =   1
         BackColor       =   0
         AutomaticAllocation=   -1  'True
         OwnerSystem     =   "System"
         Buffer          =   "Buffer"
         LUT             =   ""
         OverlayLUT      =   ""
         ExternalWindowTitle=   "Display1(0) - ActiveMIL External Window"
         FormatArrayIndex=   -1
         BackColor       =   0
         OverlayKeyColor =   268435456
         FormatArrayListSize=   24
         DisplayVisible  =   -1  'True
      End
      Begin MILDISPLAYLib.Display Display1 
         Height          =   7080
         Index           =   4
         Left            =   -74880
         TabIndex        =   31
         Top             =   600
         Width           =   9600
         _Version        =   65536
         _ExtentX        =   16933
         _ExtentY        =   12488
         _StockProps     =   1
         BackColor       =   0
         AutomaticAllocation=   -1  'True
         OwnerSystem     =   "System"
         DeviceNumber    =   4
         Buffer          =   "BufRed"
         LUT             =   ""
         OverlayLUT      =   ""
         ExternalWindowTitle=   "Display1(4) - ActiveMIL External Window"
         FormatArrayIndex=   -1
         BackColor       =   0
         OverlayKeyColor =   268435456
         FormatArrayListSize=   24
         DisplayVisible  =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '全域變數---------------------------------------------------
Dim Mx() As Long
Dim My() As Long
Dim BoxXmin As Long
Dim BoxXMax As Long
Dim BoxYmin As Long
Dim BoxYMax As Long
'Bobs Value
Dim xpoint(10), ypoint(10) As Integer
Dim BlobArea As Long
Dim ZoomS As Integer

'八方鍊碼
'Dim cd() As Byte
'ReDim cd(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
'Dim CdNum As Integer
'Dim num As Integer
'Dim CdPtX(3000) As Integer
'Dim CdPtY(3000) As Integer

''凸包
'Dim CHnum As Long
'Dim CHx() As Long
'Dim CHy() As Long


Private Sub cmdBlobs_Click()
cmdSave.Enabled = True
'--------------RGB------------------------
Call MbufCopyColor(Buffer.MilID, BufRed.MilID, M_RED)
'Call MimBinarize(BufRed.MilID, BufBinarize.MilID, M_IN_RANGE, 40, 255)

Dim rd() As Byte
Dim Out() As Byte
ReDim Out(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
ReDim rd(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
Call MbufGet2d(BufRed.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, rd(0, 0))
'-------------------------OTSU-------------------------------------------

Dim f() As Byte
ReDim f(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
Dim pValue As Integer
Dim i, j, Num, k As Integer
Dim Sgm As Single
Dim SgmMax As Single
Dim mG As Single
Dim p() As Double
ReDim p(256) As Double
Dim P1(256) As Single
Dim Mk(256) As Single
Dim bufX, bufY As Single
bufX = Buffer.SizeX
bufY = Buffer.SizeY
Call MbufGet2d(BufRed.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, f(0, 0))
'       Dim ssss As Double
'       ssss = 0
        ' 正規化
                For i = 0 To Buffer.SizeX - 1
                    For j = 0 To Buffer.SizeY - 1
'                        If k = f(i, j) Then
                            p(f(i, j)) = p(f(i, j)) + (1) / (bufX * bufY)
'                            ssss = ssss + (1) / (bufX * bufY)
'                            d(k) = d(k) + 1
                           
'                        End If
                     Next j
                Next i
'                Debug.Print ssss
'                Debug.Print Sum
'                For i = 0 To 255
'                    p(i) = p(i) / (bufX * bufY)
'                Next i
                ' 計算mG
                For k = 0 To 255
                    mG = mG + CDbl(k) * p(k)
                Next k
'        Debug.Print "mG", mG
        P1(0) = p(0)
        Mk(0) = 0
        k = 1
        'OTSU公式
        Do While k <= 255
            P1(k) = P1(k - 1) + p(k)
'            Debug.Print P1(k)
            Mk(k) = Mk(k - 1) + k * p(k)
            'If P1(k) <> 0 Then
                Sgm = ((((P1(k) * mG) - Mk(k)) ^ 2) / (0.00001 + P1(k) * CDbl(1 - P1(k)))) ^ 0.5
            If SgmMax < Sgm Then
                SgmMax = Sgm
                pValue = k
            End If
            'End If
        k = k + 1
        Loop
    'Debug.Print pValue

'-----------------------------------HE-------------------------------------------

Dim r(256), h(256) As Long
r(256) = 0
h(256) = 0
 ' 計算每個灰階值的總次數
    For i = 0 To Buffer.SizeX - 1
        For j = 0 To Buffer.SizeY - 1
            If rd(i, j) > pValue Then
                Out(i, j) = 255
            Else
                Out(i, j) = CByte(((rd(i, j) - 0) / (pValue - 0)) * 255) ' 計算0-55的Normalized
            End If
        Next j
    Next i
  'Call MbufPut2d(BufRed.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, Out(0, 0))
  Call MbufPut2d(BufHue.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, Out(0, 0))

Call MimConvolve(BufHue.MilID, BufBinarize.MilID, M_EDGE_DETECT) 'M_EDGE_DETEECT 邊緣化
Call MimBinarize(BufBinarize.MilID, BufBinarize.MilID, M_OUT_RANGE, 0, pValue)
'Call MimClose(BufBinarize.MilID, BufBinarize.MilID, 5, M_BINARY)
Call MimDilate(BufBinarize.MilID, BufBinarize.MilID, 3, M_BINARY)
Call MimErode(BufBinarize.MilID, BufBinarize.MilID, 2, M_BINARY)
'Call MimBinarize(BufRed.MilID, BufRed.MilID, M_OUT_RANGE, 0, 50)

Call MimArith(BufRed.MilID, BufBinarize.MilID, BufBinarize.MilID, M_OR)

Call MimBinarize(BufBinarize.MilID, BufBinarize.MilID, M_OUT_RANGE, 0, 50)


'Call MgraColor(M_DEFAULT, 255)
'    Call MgraRect(M_DEFAULT, BufBinarize.MilID, 0, 0, 639, 479)
'    Call MgraRect(M_DEFAULT, BufBinarize.MilID, 1, 1, 638, 478)
'Call MgraFill(M_DEFAULT, BufBinarize.MilID, 3, 477)

'Call MimArith(BufBinarize.MilID, M_NULL, BufBinarize.MilID, M_NOT)
'     --------------Blobs1-----------------

        Dim FeatureList As Long ' 用來存儲Feature的項目
        Dim BlobResult As Long ' 用來存儲 Blob 計算的結果
        Dim Totalblobs As Long
        Dim area() As Long
        Dim Agl As Double

        'allocate a feature   分配系統內存記憶體給 FeatureList,BlobResult(STACK 狀態)所以需要Free
        Call MblobAllocFeatureList(System.MilID, FeatureList)
        Call MblobAllocResult(System.MilID, BlobResult)
        
        '補黑洞
        Call MblobControl(BlobResult, M_FOREGROUND_VALUE, M_ZERO) ' 所以設前景為0
        'Call MblobControl(BlobResult, M_LATTICE, M_4_CONNECTED)
    ' 選擇需要用的Feature method 放進List 內
    Call MblobSelectFeature(FeatureList, M_ALL_FEATURES)
    '經過List 內的NULL 的特徵值方法去運算buffer的結果放入BlobResult
    Call MblobCalculate(BufBinarize.MilID, M_NULL, FeatureList, BlobResult)
    
    '找 要補的的雜訊的SIZE Set blob size
    Call MblobSelect(BlobResult, M_EXCLUBE, M_AREA, M_LESS, 1500 * ZoomS, M_NULL)
     '補起來 fill blob
    Call MblobFill(BlobResult, BufBinarize.MilID, M_EXCLUBED_BLOBS, 255)
        ' 設定和選擇前景的基本灰階值  (有0 /255 )
        Call MblobControl(BlobResult, M_FOREGROUND_VALUE, M_NONZERO)
        'Call MblobControl(BlobResult, M_LATTICE, M_4_CONNECTED)
        ' 選擇需要用的Feature method 放進List 內
        Call MblobSelectFeature(FeatureList, M_ALL_FEATURES)
        '經過List 內的NULL 的特徵值方法去運算buffer的結果放入BlobResult
        Call MblobCalculate(BufBinarize.MilID, M_NULL, FeatureList, BlobResult)
        Call MblobGetNumber(BlobResult, Totalblobs)
        'Debug.Print Totalblobs
        '找 要補的的雜訊的SIZE Set blob size
        Call MblobSelect(BlobResult, M_EXCLUBE, M_AREA, M_LESS, 5000 * ZoomS, M_NULL)
         '補起來 fill blob
        Call MblobFill(BlobResult, BufBinarize.MilID, M_EXCLUBED_BLOBS, 0)

        '每次算完Fill 要再Calculate 特徵值
        Call MblobCalculate(BufBinarize.MilID, M_NULL, FeatureList, BlobResult)
        Call MblobFree(BlobResult) ' 因為後進需要先出(FILO原則)
        Call MblobFree(FeatureList)
        Call MimDilate(BufBinarize.MilID, BufBinarize.MilID, 2, M_BINARY)

        Dim buf() As Byte
        ReDim buf(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
        Call MbufGet2d(BufBinarize.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, buf(0, 0))

        For X = 0 To Buffer.SizeX - 1
            buf(X, 0) = 0
            buf(X, Buffer.SizeY - 1) = 0
        Next X
        For Y = 0 To Buffer.SizeY - 1
            buf(0, Y) = 0
            buf(Buffer.SizeX - 1, Y) = 0
        Next Y


'
'
    '找邊緣點
    Dim gtimg() As Byte
        ReDim gtimg(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
    Dim Xc, Xd, Yc, Yd As Integer
    Dim Dp(2) As Single
    Dim M, Slope As Double
    Dim a, b, c, openCt, count, way As Integer
    Dim Min As Integer
    
            Call MbufGet2d(BufBinarize.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, gtimg(0, 0))
            Call MbufPut2d(BufBinarize.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, gtimg(0, 0))
            'Call MbufPut2d(Buffer1.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, gtimg(0, 0))
            i = 0
            Call MgraColor(M_DEFAULT, 150)
                For Y = 1 To Buffer.SizeY - 2
                   If gtimg(20, Y) <> gtimg(20, Y + 1) Then
                   xpoint(i) = 20
                   ypoint(i) = Y + 1
                   'Debug.Print xpoint(i), ypoint(i)
                   Call MgraArc(M_DEFAULT, BufChainCd.MilID, 20, Y + 1, 10, 10, 0, 360)
                    i = i + 1
                   End If
                   If gtimg(Buffer.SizeX - 20, Y) <> gtimg(Buffer.SizeX - 20, Y + 1) Then
                   xpoint(i) = Buffer.SizeX - 20
                   ypoint(i) = Y + 1
                   'Debug.Print xpoint(i), ypoint(i)
                   Call MgraArc(M_DEFAULT, BufChainCd.MilID, Buffer.SizeX - 20, Y + 1, 10, 10, 0, 360)
                    i = i + 1
                   End If
               Next Y
               For X = 1 To Buffer.SizeX - 2
                   If gtimg(X, 20) <> gtimg(X + 1, 20) Then
                   xpoint(i) = X + 1
                   ypoint(i) = 20
                   'Debug.Print xpoint(i), ypoint(i)
                   Call MgraArc(M_DEFAULT, BufChainCd.MilID, X + 1, 20, 10, 10, 0, 360)
                   i = i + 1
                   End If
                   If gtimg(X, Buffer.SizeY - 20) <> gtimg(X + 1, Buffer.SizeY - 20) Then
                   xpoint(i) = X + 1
                   ypoint(i) = Buffer.SizeY - 20
                   'Debug.Print xpoint(i), ypoint(i)
                   Call MgraArc(M_DEFAULT, BufChainCd.MilID, X + 1, Buffer.SizeY - 20, 10, 10, 0, 360)
                   i = i + 1
                   End If
               Next X
                Call MgraColor(M_DEFAULT, 0)
                Call MgraLine(M_DEFAULT, BufBinarize.MilID, xpoint(0), ypoint(0), xpoint(1), ypoint(1))
                Call MgraLine(M_DEFAULT, BufBinarize.MilID, xpoint(0) - 1, ypoint(0) + 1, xpoint(1) - 1, ypoint(1) + 1)
                Call MgraLine(M_DEFAULT, BufBinarize.MilID, xpoint(0) - 1, ypoint(0) + 1, xpoint(1) - 1, ypoint(1) + 1)
                Call MgraLine(M_DEFAULT, BufBinarize.MilID, xpoint(0) - 1, ypoint(0) + 1, xpoint(1) - 1, ypoint(1) + 1)
                    
    
    
    '--------------Blobs2-----------------
    '先補黑洞
    Call MblobAllocFeatureList(System.MilID, FeatureList)
    Call MblobAllocResult(System.MilID, BlobResult)
    Call MimOpen(BufBinarize.MilID, BufBinarize.MilID, 3, M_BINARY)
    ' 設定和選擇前景的基本灰階值  (有0 /255 )
    Call MblobControl(BlobResult, M_FOREGROUND_VALUE, M_ZERO) ' 所以設前景為0
    'Call MblobControl(BlobResult, M_LATTICE, M_4_CONNECTED)
    ' 選擇需要用的Feature method 放進List 內
    Call MblobSelectFeature(FeatureList, M_ALL_FEATURES)
    '經過List 內的NULL 的特徵值方法去運算buffer的結果放入BlobResult
    Call MblobCalculate(BufBinarize.MilID, M_NULL, FeatureList, BlobResult)
    
    '找 要補的的雜訊的SIZE Set blob size
    Call MblobSelect(BlobResult, M_EXCLUBE, M_AREA, M_LESS, 1500 * ZoomS, M_NULL)
     '補起來 fill blob
    Call MblobFill(BlobResult, BufBinarize.MilID, M_EXCLUBED_BLOBS, 255)
   
   ' 設定和選擇前景的基本灰階值  (有0 /255 )
    Call MblobControl(BlobResult, M_FOREGROUND_VALUE, M_NONZERO)
    Call MblobControl(BlobResult, M_LATTICE, M_4_CONNECTED)
    ' 選擇需要用的Feature method 放進List 內
    Call MblobSelectFeature(FeatureList, M_ALL_FEATURES)
    '經過List 內的NULL 的特徵值方法去運算buffer的結果放入BlobResult
    Call MblobCalculate(BufBinarize.MilID, M_NULL, FeatureList, BlobResult)
    Call MblobGetNumber(BlobResult, Totalblobs)
    'Debug.Print Totalblobs

    
    '找 要補的的雜訊的SIZE Set blob size
    Call MblobSelect(BlobResult, M_EXCLUBE, M_AREA, M_LESS, 8000 * ZoomS, M_NULL)
     '補起來 fill blob
    Call MblobFill(BlobResult, BufBinarize.MilID, M_EXCLUBED_BLOBS, 0)
    
    '每次算完Fill 要再Calculate 特徵值
    Call MblobCalculate(BufBinarize.MilID, M_NULL, FeatureList, BlobResult)
     

    ' 把BlobResult 的值放進TotalBlobs
    Call MblobGetNumber(BlobResult, Totalblobs)
    
    'Debug.Print Totalblobs
   'txtBlobs.Text = Totalblobs
        If Totalblobs > 0 Then
            ReDim Mx(Totalblobs - 1) As Long
            ReDim My(Totalblobs - 1) As Long
'            Dim Barea(2) As Long
            'funtion
            Call MblobGetResult(BlobResult, M_BOX_X_MIN + M_TYPE_LONG, BoxXmin)
            Call MblobGetResult(BlobResult, M_BOX_X_MAX + M_TYPE_LONG, BoxXMax)
            Call MblobGetResult(BlobResult, M_BOX_Y_MIN + M_TYPE_LONG, BoxYmin)
            Call MblobGetResult(BlobResult, M_BOX_Y_MAX + M_TYPE_LONG, BoxYMax)
            Call MblobGetResult(BlobResult, M_CENTER_OF_GRAVITY_X + M_TYPE_LONG, Mx(0))
            Call MblobGetResult(BlobResult, M_CENTER_OF_GRAVITY_Y + M_TYPE_LONG, My(0))
           Call MblobGetResult(BlobResult, M_AREA + M_TYPE_LONG, BlobArea)
            Call MblobGetResult(BlobResult, M_AXIS_PRINCIPAL_ANGLE + M_TYPE_DOUBLE, Agl)
         End If
            Call MbufGet2d(BufBinarize.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, gtimg(0, 0))
            'Call MbufPut2d(BufBinarize2.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, gtimg(0, 0))
'
    '資料堆疊後~~需要釋放內存記憶體(FILO)     "Free"
    Call MblobFree(BlobResult) ' 因為後進需要先出(FILO原則)
    Call MblobFree(FeatureList)
    i = 0
'    Debug.Print Barea(0)
'    Debug.Print Barea(1)
End Sub

Private Sub cmdLoad_Click()
    Buffer.Clear
    CdNum = 0
    CHnum = 0
    
    Dim i As Integer
'    For i = 0 To 5
''    OptBands.Item(i) = False
'    Next i
    CommonDialog1.DialogTitle = "開啟舊檔"
    CommonDialog1.Filter = "*|*.*"
    CommonDialog1.ShowOpen
    
    Buffer.Load CommonDialog1.FileName, True
        For i = 0 To 3
        Display1(i).Width = Buffer.SizeX * 15 '把display大小固定為原圖的大小
        Display1(i).Height = Buffer.SizeY * 15
    Next i
    If Buffer.SizeX > 640 And Buffer.SizeY > 480 Then
    ZoomS = 4
    For i = 0 To 4
        Display1(i).ZoomX = -2
        Display1(i).ZoomY = -2
    Next i
    '----------------------------------
    BufRed.Free '釋放bufRed 的設定
    BufRed.SizeX = Buffer.SizeX
    BufRed.SizeY = Buffer.SizeY
    BufRed.NumberOfBands = 1
    BufRed.Allocate '建立
    '-----------------------------------
    BufBinarize.Free
    BufBinarize.SizeX = Buffer.SizeX
    BufBinarize.SizeY = Buffer.SizeY
    BufBinarize.NumberOfBands = 1
    BufBinarize.Allocate
    '-----------------------------------
    Buffer1.Free
    Buffer1.SizeX = Buffer.SizeX
    Buffer1.SizeY = Buffer.SizeY
    Buffer1.NumberOfBands = 1
    Buffer1.Allocate
    '-----------------------------------
    BufChainCd.Free
    BufChainCd.SizeX = Buffer.SizeX
    BufChainCd.SizeY = Buffer.SizeY
    BufChainCd.NumberOfBands = 1
    BufChainCd.Allocate
    '-----------------------------------
    BufHue.Free
    BufHue.SizeX = Buffer.SizeX
    BufHue.SizeY = Buffer.SizeY
    BufHue.NumberOfBands = 1
    BufHue.Allocate
    '-----------------------------------
    Else
    ZoomS = 1
    End If

End Sub

Private Sub cmdSave_Click()
    
    Do
        BufBinarize.Save App.Path + "/image/Img" + Str(1 + Num) + ".tif"
        BufRed.Save App.Path + "/image/ImgR" + Str(1 + Num) + ".tif"
         BufChainCd.Save App.Path + "/image/ImgCCD" + Str(1 + Num) + ".tif"
        cmdSave.Enabled = False
    Loop Until cmdSave.Enabled = False
        
        If cmdSave.Enabled = False Then
            Num = Num + 1
        End If
End Sub

Private Sub cmdStart_Click()
    Buffer1.Clear
    BufChainCd.Clear
    Label2(0).Caption = "processing"
    
    Dim ChainPx() As Long '鏈碼陣列
    Dim ChainPy() As Long
    Dim countChain As Long '鏈碼存取的步數(周長)
    Dim SerialNum() As Long '依鏈碼走訪順序給予序號
    Dim CHx() As Long '凸包點陣列
    Dim CHy() As Long
    Dim M As Long '凸包數目
    Dim f() As Byte
    Dim cd() As Byte
      ReDim cd(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
    Dim cx, cy, X, Y, Num, start As Integer
        Num = 0
    Dim ex()
    Dim ey()
        ex = Array(-1, 0, 1, 1, 1, 0, -1, -1)
        ey = Array(-1, -1, -1, 0, 1, 1, 1, 0)
    

    '將八方鍊碼走訪過的座標存成陣列形式
    '用自訂的結構型態存點座標
    Dim hullPointx() As Long
         ReDim hullPointx(2000 * ZoomS) As Long
    Dim hullPointy() As Long
         ReDim hullPointy(2000 * ZoomS) As Long
         ReDim SerialNum(2000 * ZoomS) As Long
         
         ReDim ChainPx(2000 * ZoomS) As Long
         ReDim ChainPy(2000 * ZoomS) As Long
         countChain = 0

'    Dim f() As Byte
    ' 設定 f 的範圍大小
        ReDim f(Buffer.SizeX - 1, Buffer.SizeY - 1) As Byte
    
    Call MbufGet2d(BufBinarize.MilID, 0, 0, Buffer.SizeX, Buffer.SizeY, f(0, 0))
    '掃描圖面以取得鏈碼起始點
    For X = 0 To Buffer.SizeX - 1
        For Y = 0 To Buffer.SizeY - 2
            If (f(X, Y) = 255) Then
                cd(X, Y) = f(X, Y)
                cx = X
                cy = Y
                hullPointx(0) = X
                hullPointy(0) = Y
                ChainPx(0) = X
                ChainPy(0) = Y
                SerialNum(0) = 0
                X = Buffer.SizeX * 1.1
                Y = Buffer.SizeY * 1.1 '直接修改值以離開迴圈
            End If
        Next Y
    Next X
    '記好起點
'    X = cx
'    Y = cy
    Call MgraColor(M_DEFAULT, 255)
    '在畫面上標記出鏈碼起點
    Call MgraArc(M_DEFAULT, BufChainCd.MilID, ChainPx(0), ChainPy(0), 8, 8, 0, 360)

    Call MgraColor(M_DEFAULT, 100)
    Do '走訪圖像邊緣
        If (f(cx + ex(Num), cy + ey(Num)) = 255) Then
            cx = cx + ex(Num)
            cy = cy + ey(Num)
            cd(cx, cy) = f(cx, cy)
            
            Num = Num - 3
            
            countChain = countChain + 1
            hullPointx(countChain) = cx
            hullPointy(countChain) = cy
            ChainPx(countChain) = cx
            ChainPy(countChain) = cy
            SerialNum(countChain) = countChain '鏈碼序號

            Call MgraDot(M_DEFAULT, BufChainCd.MilID, cx, cy)
            Call MgraDot(M_DEFAULT, Buffer1.MilID, cx, cy)
        End If
        
        If Num = 7 Then
                Num = 0
            ElseIf Num < 0 Then
                Num = Num + 8
            Else
                 Num = Num + 1
        End If
    Loop Until ((cx + ex(Num) = ChainPx(0)) And (cy + ey(Num)) = ChainPy(0))

'    For k = 0 To 2
'        Debug.Print k, hullPointx(k), hullPointy(k)
'    Next k
    
    Shape1(0).BackColor = RGB(255, 0, 0)
    '將資訊顯示在視窗
    Label2(2).Caption = countChain
'    Debug.Print "八方 DONE"

        ReDim CHx(countChain + 1) As Long
        ReDim CHy(countChain + 1) As Long

    Dim tempX, tempY, tempSN As Long
    
'   排列所有參考點(也就是指八方鍊碼所存的點) hullpointx(),hullpointy()
'   排列順序 -->   依照X座標再依照Y座標
    For i = 0 To countChain - 1
        For j = 0 To countChain - 2
                If hullPointx(j) > hullPointx(j + 1) Then
                    tempX = hullPointx(j)
                    hullPointx(j) = hullPointx(j + 1)
                    hullPointx(j + 1) = tempX
                    tempY = hullPointy(j)
                    hullPointy(j) = hullPointy(j + 1)
                    hullPointy(j + 1) = tempY
                End If
        Next j
    Next i

    For i = 0 To countChain - 1
        For j = 0 To countChain - 2
                If hullPointx(j) = hullPointx(j + 1) Then
                    If hullPointy(j) > hullPointy(j + 1) Then
                        tempY = hullPointy(j)
                        hullPointy(j) = hullPointy(j + 1)
                        hullPointy(j + 1) = tempY
                    End If
                End If
        Next j
    Next i

'    Debug.Print "排列 DONE"
    '標示出凸包起點位置(十字準星狀)
    Call MgraColor(M_DEFAULT, 255)
    Call MgraLine(M_DEFAULT, BufChainCd.MilID, hullPointx(0) - 5, hullPointy(0), hullPointx(0) + 5, hullPointy(0))
    Call MgraLine(M_DEFAULT, BufChainCd.MilID, hullPointx(0), hullPointy(0) - 5, hullPointx(0), hullPointy(0) + 5)

    Dim a, b, c, cosTH, TempCosTH As Double
    Dim hullN As Long
        M = 0 '凸包點數目
        hullN = 0 '凸包判斷時用以去除迴圈重複的變數，決定下一次凸包判斷起點
        CHx(0) = hullPointx(0)
        CHy(0) = hullPointy(0)

'   第一輪掃描 , 由左而右掃上半部
    For j = 1 To countChain - 1
'        利用餘弦定理，三個邊可決定指定夾角的特性
'        預設參考邊a為點之往y+ 之垂直線, 因此上半部由左往右掃描時, 凸包轉折邊緣為最大角度
'        張開角度越大則cosTH值越小 cos(0~180度)=1~0~-1
        TempCosTH = 1
        For i = j To countChain - 1
            '掃描所有尚未參考過的點，不斷更新至角度最大處
            
            a = 479 - CHy(M)
            b = ((CHx(M) - hullPointx(i)) ^ 2 + (CHy(M) - hullPointy(i)) ^ 2) ^ 0.5
            c = ((CHx(M) - hullPointx(i)) ^ 2 + (479 - hullPointy(i)) ^ 2) ^ 0.5
            
            If CHy(M) = hullPointy(i) Then
                cosTH = 0
            ElseIf CHx(M) = hullPointx(i) Then
                If CHy(M) < hullPointy(i) Then
                    cosTH = 1
                ElseIf CHy(M) > hullPointy(i) Then
                    cosTH = -1
                End If
            Else
                cosTH = (a ^ 2 + b ^ 2 - c ^ 2) / (2 * a * b)
            End If
            
            If (cosTH <= TempCosTH) Then
'             更新至最大角度即為凸包上部邊緣
                TempCosTH = cosTH
                hullN = i

            End If
        Next i

        M = M + 1
        CHx(M) = hullPointx(hullN)
        CHy(M) = hullPointy(hullN)
'         hullN用在此處 , 變更下一個凸包需掃描的起始點
         j = hullN + 1
    Next j
    
    '第二輪，掃下半部
    For j = countChain - 1 To 1 Step -1
'        Debug.Print m, j
        '預設參考邊a為點之往y-之垂直線，因此下半部由右往左掃描時，凸包轉折邊緣為最大角度
        '張開角度越大則cosTH值越小 cos(0~180度)=1~0~-1
        TempCosTH = 1
        For i = j To 1 Step -1
            '掃描所有尚未參考過的點，更新至角度最大處
            a = CHy(M)
            b = ((CHx(M) - hullPointx(i)) ^ 2 + (CHy(M) - hullPointy(i)) ^ 2) ^ 0.5
            c = ((CHx(M) - hullPointx(i)) ^ 2 + (hullPointy(i)) ^ 2) ^ 0.5
            If CHy(M) = hullPointy(i) Then
                cosTH = 0
            ElseIf CHx(M) = hullPointx(i) Then
                If CHy(M) > hullPointy(i) Then
                    cosTH = 1
                ElseIf CHy(M) < hullPointy(i) Then
                    cosTH = -1
                End If
            Else
                cosTH = (a ^ 2 + b ^ 2 - c ^ 2) / (2 * a * b)
            End If

            If (cosTH <= TempCosTH) Then
'            更新至最大角度即為凸包下部邊緣
                TempCosTH = cosTH
                hullN = i

            End If
        Next i

        M = M + 1
        CHx(M) = hullPointx(hullN)
        CHy(M) = hullPointy(hullN)
'        hullN變更下一個凸包需掃描的起始點
         j = hullN - 1
    Next j
        
'    第二輪結束
'    上下皆掃描完成，CHx()與CHy()即為凸包點座標陣列
    Shape1(1).BackColor = RGB(255, 0, 0)
    Label2(3).Caption = M  '在視窗顯示凸包點數目
'    Debug.Print "凸包 DONE"

'    標記凸包點
    For k = 0 To M - 1
        Call MgraArc(M_DEFAULT, BufChainCd.MilID, CHx(k), CHy(k), 2, 2, 0, 360)
    Next k
    
'   複製第一點至陣列尾 , 方便畫線
    M = M + 1
    CHx(M) = CHx(0)
    CHy(M) = CHy(0)
    '畫出凸包邊緣
    Call MgraColor(M_DEFAULT, 255)
    For k = 0 To M - 1
        Call MgraLine(M_DEFAULT, BufBinarize.MilID, CHx(k), CHy(k), CHx(k + 1), CHy(k + 1))
    Next k



    Dim hLength, Slope, Slope2 As Double
    Dim Linex() As Long
    Dim Liney() As Long

    '整合凸包邊緣線段與鏈碼座標關係
    
    'Dim jchang As Long
    
    Dim CHn() As Long
        ReDim CHn(M + 1) As Long

    For i = 0 To M - 1
        For j = Num To countChain - 1
            If (CHx(i) = ChainPx(j)) Then
                If (CHy(i) = ChainPy(j)) Then
                    CHn(i) = SerialNum(j)
                    'j = j
                End If
            End If
        Next j
'        Debug.Print CHn(i), i
    Next i
        

    '過濾出有興趣線段 --> 手指間距
    '除去轉折處短線
    '分析相較凸包，凹陷最深之處與該凸包線段的距離
    '並建立 凸包線 : 凹陷距離 的參考值
    '將手指中間部分萃取出來
    
    '直線 ax+by+c=0
    '計算點 Lx,Ly
    '點至線距離 d= abs(a*x0+b*y0+c)/(a^2+b^2)^0.5
    '直線用點斜式 mx-y+b=0
    'd=abs(m*Lx-Ly+b)/(m^2+1)^0.5
    
    Dim bConst As Long '點斜式常數
    Dim d, dMax As Double '點與線距離
    Dim ratio As Double
    Dim Gesture As Integer
      Gesture = 0
    Dim Wrist As Integer
      Wrist = 0
    Dim Uxpt, Uypt, UxSum, UySum, UwxSum, UwySum, UttX, UttY As Long
    Dim UwX() As Integer
    Dim UwY() As Integer
    ReDim UwX(M) As Integer
    ReDim UwY(M) As Integer
    Dim UcX() As Integer
    Dim UcY() As Integer
    ReDim UcX(M) As Integer
    ReDim UcY(M) As Integer

    UttX = 0
    UttY = 0
    For i = 0 To M - 1
        hLength = ((CHx(i) - CHx(i + 1)) ^ 2 + (CHy(i) - CHy(i + 1)) ^ 2) ^ 0.5
        If CHx(i) = CHx(i + 1) Then
            Slope = 9999
        Else
            Slope = (CHy(i) - CHy(i + 1)) / (CHx(i) - CHx(i + 1))
        End If
        
        If hLength >= 20 Then '去除過短的轉折處線段
            dMax = 0
            '用來計算的線段點斜式
            bConst = CHy(i) - Slope * CHx(i)
            For j = CHn(i) To CHn(i + 1) Step 1
                '點線距離
                d = Abs(Slope * ChainPx(j) - ChainPy(j) + bConst) / ((Slope ^ 2 + 1) ^ 0.5)
                If d >= dMax Then
                    dMax = d '更新至最大距離
                    Uxpt = ChainPx(j)
                    Uypt = ChainPy(j)
                    
                End If
            Next j
            
            
            
            ratio = dMax / hLength
            
            If ratio >= 0.3 Then
                UcX(Gesture) = Uxpt
                UxSum = UxSum + Uxpt
                UcY(Gesture) = Uypt
                UySum = UySum + Uypt
'                Debug.Print Gesture, UcX(Gesture), UcY(Gesture)
                Gesture = Gesture + 1
                
                Call MgraText(M_DEFAULT, BufBinarize.MilID, CHx(i) + 6, CHy(i), "Finger")
                Call MgraColor(M_DEFAULT, 255)
               '繪出被計算的凸包線段
                Call MgraLine(M_DEFAULT, BufBinarize.MilID, CHx(i), CHy(i), CHx(i + 1), CHy(i + 1))
                Call MgraLine(M_DEFAULT, BufChainCd.MilID, CHx(i), CHy(i), CHx(i + 1), CHy(i + 1))
                'If ratio >= 0.9 Then
                    Call MgraText(M_DEFAULT, BufBinarize.MilID, Uxpt, Uypt, "Upt")
                    Debug.Print ratio, "Upt"
                'End If
                
'            ElseIf ratio > 0.1 And ratio < 0.3 Then
'                Call MgraColor(M_DEFAULT, 255)
'                Call MgraText(M_DEFAULT, BufBinarize.MilID, Uxpt, Uypt, "Wpt")
'                Debug.Print ratio, "Wpt"
'            ElseIf ratio >= 0.02 And ratio < 0.3 Then
'                UwX(Wrist) = Uxpt
'                UwY(Wrist) = Uypt
'                   Wrist = Wrist + 1
''
'                    Debug.Print ratio, "Wpt"
'                    Call MgraText(M_DEFAULT, BufBinarize.MilID, Uxpt, Uypt, "Wpt")
'
            Else
                Call MgraText(M_DEFAULT, BufBinarize.MilID, Uxpt, Uypt, "Wpt")
                Call MgraText(M_DEFAULT, BufBinarize.MilID, CHx(i) + 6, CHy(i), "x")
                Debug.Print ratio, "X"
                Call MgraColor(M_DEFAULT, 100)
               '繪出被計算的凸包線段
                Call MgraLine(M_DEFAULT, BufBinarize.MilID, CHx(i), CHy(i), CHx(i + 1), CHy(i + 1))
                Call MgraLine(M_DEFAULT, BufChainCd.MilID, CHx(i), CHy(i), CHx(i + 1), CHy(i + 1))
'                End If
            End If
'            Debug.Print i, ratio
            
            Call MgraColor(M_DEFAULT, 255)
            '繪出被計算的凸包線段
            Call MgraLine(M_DEFAULT, Buffer1.MilID, CHx(i), CHy(i), CHx(i + 1), CHy(i + 1))
            Call MgraLine(M_DEFAULT, BufChainCd.MilID, CHx(i), CHy(i), CHx(i + 1), CHy(i + 1))
            '標示凸包點的序列號碼
            Call MgraText(M_DEFAULT, Buffer1.MilID, CHx(i) + 6, CHy(i), i)
            '標示出凸包點的鏈碼序列號
            Call MgraText(M_DEFAULT, BufChainCd.MilID, CHx(i) + 6, CHy(i), CHn(i))
        
        End If
        
    
    Next i
'    Debug.Print Wrist, Gesture
    If Gesture > 0 Then
        UttX = UxSum / Gesture
        UttY = UySum / Gesture
        Call MgraColor(M_DEFAULT, 150)
        Call MgraText(M_DEFAULT, BufBinarize.MilID, UttX, UttY, "O")
        Call MgraArc(M_DEFAULT, BufBinarize.MilID, UttX, UttY, 130, 130, 0, 360)
        Call MgraColor(M_DEFAULT, 0)
        Call MgraArc(M_DEFAULT, BufBinarize2.MilID, UttX, UttY, 130, 130, 0, 360)
    Else
      
        Call MgraColor(M_DEFAULT, 150)
        Call MgraText(M_DEFAULT, BufBinarize.MilID, Mx(0), My(0), "O")
'        Call MgraLine(M_DEFAULT, BufBinarize.MilID, Mx(0), My(0), Mx(0) + 50, My(0) + 50)
'        Call MgraLine(M_DEFAULT, BufBinarize.MilID, Mx(0), My(0), Mx(0) - 50, My(0) - 50)
'        Call MgraArc(M_DEFAULT, BufBinarize.MilID, Mx(0), My(0), 100, 100, 0, 360)
        Call MgraColor(M_DEFAULT, 0)
'        Call MgraArc(M_DEFAULT, BufBinarize2.MilID, Mx(0), My(0), 100, 100, 0, 360)
    End If
    If Gesture = 0 Then
        Label2(1).Caption = "STONE/ zero"
    ElseIf Gesture = 1 Then
        Label2(1).Caption = "SCISSORs"
    ElseIf Gesture = 2 Then
        Label2(1).Caption = "three"
    ElseIf Gesture = 3 Then
        Label2(1).Caption = "four"
    Else
        Label2(1).Caption = "PAPER/five"
    End If
    Label2(0).Caption = "Done"
    Shape1(2).BackColor = RGB(255, 0, 0)
        
        'Blobs 找手掌
'
'        Dim FeatureList As Long ' 用來存儲Feature的項目
'        Dim BlobResult As Long ' 用來存儲 Blob 計算的結果
'        Dim Totalblobs As Long
'
'        'allocate a feature   分配系統內存記憶體給 FeatureList,BlobResult(STACK 狀態)所以需要Free
'        Call MblobAllocFeatureList(System.MilID, FeatureList)
'        Call MblobAllocResult(System.MilID, BlobResult)
'        '去手腕
'
'    Call MblobAllocFeatureList(System.MilID, FeatureList)
'    Call MblobAllocResult(System.MilID, BlobResult)
'   ' 設定和選擇前景的基本灰階值  (有0 /255 )
'    Call MblobControl(BlobResult, M_FOREGROUND_VALUE, M_NONZERO)
'    Call MblobControl(BlobResult, M_LATTICE, M_4_CONNECTED)
'    ' 選擇需要用的Feature method 放進List 內
'    Call MblobSelectFeature(FeatureList, M_ALL_FEATURES)
'    '經過List 內的NULL 的特徵值方法去運算buffer的結果放入BlobResult
'    Call MblobCalculate(BufBinarize2.MilID, M_NULL, FeatureList, BlobResult)
'    Call MblobGetNumber(BlobResult, Totalblobs)
'    'Debug.Print Totalblobs
'    '找 要補的的雜訊的SIZE Set blob size
'    Call MblobSelect(BlobResult, M_EXCLUBE, M_AREA, M_LESS, BlobArea / 2, M_NULL)
'     '補起來 fill blob
'    Call MblobFill(BlobResult, BufBinarize2.MilID, M_EXCLUBED_BLOBS, 0)
'
'    '每次算完Fill 要再Calculate 特徵值
'    Call MblobCalculate(BufBinarize2.MilID, M_NULL, FeatureList, BlobResult)
'
'    Call MblobFree(BlobResult) ' 因為後進需要先出(FILO原則)
'    Call MblobFree(FeatureList)
'
'
'        Call MgraColor(M_DEFAULT, 255)
'    Call MgraLine(M_DEFAULT, Buffer.MilID, BoxXmin, BoxYmin, BoxXMax, BoxYmin)
'    Call MgraLine(M_DEFAULT, Buffer.MilID, BoxXmin, BoxYmin, BoxXmin, BoxYMax)
'    Call MgraLine(M_DEFAULT, Buffer.MilID, BoxXmin, BoxYMax, BoxXMax, BoxYMax)
'    Call MgraLine(M_DEFAULT, Buffer.MilID, BoxXMax, BoxYmin, BoxXMax, BoxYMax)
End Sub
Private Sub cmdUnload_Click()
Unload Me
End Sub

