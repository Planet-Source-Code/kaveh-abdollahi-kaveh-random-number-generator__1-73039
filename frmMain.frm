VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Kaveh Random Engine"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCmp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "1000000"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCmpSpd 
      Caption         =   "Speeds Ccompare"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   6720
      Width           =   1455
   End
   Begin VB.PictureBox picReverse 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   53.15
      ScaleMode       =   0  'User
      ScaleWidth      =   1328.294
      TabIndex        =   5
      Top             =   120
      Width           =   9255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   7
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "100"
      Top             =   2175
      Width           =   735
   End
   Begin VB.TextBox txtRndCount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "1000"
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CheckBox chkAuto 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto"
      Height          =   255
      Left            =   1500
      TabIndex        =   6
      Top             =   6840
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   6600
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS SystemEx"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   9255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS SystemEx"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   9255
   End
   Begin VB.CommandButton cmdMakeRnd 
      Caption         =   "Make Random"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSend2Ex 
      Caption         =   "Send To Excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "VB Random"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Kaveh Random"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private K_Cou As Double, K_Lrnd As Double, K_Tmp As Long


Function KRandom(Optional Range As Double) As Double
    K_Cou = K_Cou + 1
'    KRandom = ((K_Lrnd + 3) * 7 * K_Cou) / 5
    KRandom = (K_Lrnd * K_Cou * 2 + 5) / 7
    KRandom = (KRandom - Int(KRandom)) * Range
    K_Lrnd = KRandom
End Function

Private Sub cmdCmpSpd_Click()
Dim ti1 As Long, ti2 As Long, x As Long, tmp As Double, tmp2 As Long
    
    tmp2 = txtCmp
    ti1 = GetTickCount
    For x = 1 To tmp2
        tmp = Rnd(1)
    Next x
    ti1 = GetTickCount - ti1
    
    ti2 = GetTickCount
    For x = 1 To tmp2
        tmp = KRandom(1)
    Next x
    ti2 = GetTickCount - ti2
    List2.Clear
    List2.AddItem " VB Rnd   : " & ti1 / 1000
    List2.AddItem " K Random : " & ti2 / 1000

End Sub

Private Sub cmdMakeRnd_Click()
Dim Rc As Long, x As Long, y As Long, Telo1 As Double, Telo2 As Double, Sum1 As Double, Sum2 As Double              '''' only for test KRandom()
Dim tSgR1 As Long, tSgR2 As Long, CSgR1 As Long, CSgR2 As Long

List1.Visible = False
List1.Clear: List2.Clear
picReverse.Cls

    CSgR1 = 0: CSgR2 = 0: tSgR1 = 0: tSgR2 = 0
    Sum1 = 0: Sum2 = 0
    Rc = (txtRndCount)
    y = 1
    ReDim RndArray(0 To Rc, 0 To y)
    picReverse.ScaleWidth = Rc
    
    RndArray(1, 0) = 0
    RndArray(1, 1) = 0
    
    picReverse.Line (-101, 25)-(-2, 26)
    For x = 1 To Rc
        RndArray(x, 0) = KRandom(1)
        Telo1 = Telo1 + Abs(RndArray(x, 0) - RndArray(x - 1, 0))
        If Sgn(RndArray(x, 0) - RndArray(x - 1, 0)) <> Sgn(tSgR1) Then tSgR1 = Sgn(RndArray(x, 0) - RndArray(x - 1, 0)): CSgR1 = CSgR1 + 1
        Sum1 = Sum1 + RndArray(x, 0)
        
         picReverse.Line -(x, (RndArray(x, 0) - RndArray(x - 1, 0)) * 10 + 15), vbRed
         picReverse.Line -(x, (RndArray(x, 0) - RndArray(x - 1, 0)) * 10 + 15), vbRed
    
    Next x
    picReverse.Line (-101, 25)-(-2, 26)
        
        ''''''''''''''''''''''''''''''''''''''''''''
        
    For x = 1 To Rc
        RndArray(x, 1) = Rnd
        Telo2 = Telo2 + Abs(RndArray(x, 1) - RndArray(x - 1, 1))
        If Sgn(RndArray(x, 1) - RndArray(x - 1, 1)) <> Sgn(tSgR2) Then tSgR2 = Sgn(RndArray(x, 1) - RndArray(x - 1, 1)): CSgR2 = CSgR2 + 1
        Sum2 = Sum2 + RndArray(x, 1)
        
        picReverse.Line -(x, 40 - (RndArray(x, 1) - RndArray(x - 1, 1)) * 10), vbBlue
        picReverse.Line -(x, 40 - (RndArray(x, 1) - RndArray(x - 1, 1)) * 10), vbBlue
        
      List1.AddItem Format$((RndArray(x, 0)), "0.###,###,###,###,###0") & "    " & vbTab & Format$((RndArray(x, 1)), "0.################0")
    Next x
    picReverse.Line (-1, 25)-(-2, 26)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    List1.ListIndex = List1.ListCount - 1
    
    List2.AddItem "K-Random.. Telorance : " & Round(Telo1 / (x - 1), 4) & "    " & "Avrage : " & Round(Sum1 / (x - 1), 4) & "    " & "Reverse : " & CSgR1 ' (CSgR1 / Rc)
    List2.AddItem "VB-RND.....  Telorance : " & Round(Telo2 / (x - 1), 4) & "    " & "Avrage : " & Round(Sum2 / (x - 1), 4) & "    " & "Reverse : " & CSgR2  ' (CSgR2 / Rc)
    List2.AddItem "--------------------------------------------------------------------------------------------------------------------------------"
    
    List2.ListIndex = List2.ListCount - 1
    
  
cmdSend2Ex.Enabled = True
List1.Visible = True

End Sub

Private Sub cmdSend2Ex_Click()
  ExcelSaveArray ex_Num
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 1 And Timer1.Interval > 10 Then Timer1.Interval = Timer1.Interval - 10
    If Index = 0 And Timer1.Interval < 60000 Then Timer1.Interval = Timer1.Interval + 10
    txtInterval = Timer1.Interval
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then cmdMakeRnd_Click
End Sub

Private Sub Form_Load()
    K_Cou = (Timer / 1.123456789 + 1)
    K_Lrnd = K_Cou / (0.123456789 * 0.31)
End Sub

Private Sub Timer1_Timer()
    If chkAuto Then cmdMakeRnd_Click
End Sub

Private Sub txtRndCount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdMakeRnd_Click
    DoEvents
End Sub

