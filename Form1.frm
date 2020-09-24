VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Spirals :-: ZATRiX@load.com:-:http://zatrix.i8.com"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "STOP"
      Height          =   195
      Left            =   0
      TabIndex        =   26
      Top             =   4530
      Width           =   4935
   End
   Begin VB.OptionButton Option9 
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   5400
      Width           =   255
   End
   Begin VB.OptionButton Option8 
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton Option7 
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   4920
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pen Width"
      Height          =   1125
      Left            =   5040
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   360
         X2              =   1320
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   360
         X2              =   1320
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   1320
         Y1              =   405
         Y2              =   405
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   21
      Top             =   3840
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   20
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pen Color"
      Height          =   3375
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
      Begin VB.OptionButton Option6 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   1440
      Max             =   100
      Min             =   -100
      TabIndex        =   3
      Top             =   5520
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   1440
      Max             =   100
      Min             =   -100
      TabIndex        =   2
      Top             =   5160
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   1440
      Max             =   100
      Min             =   -100
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4455
      Left            =   0
      ScaleHeight     =   219.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   243.75
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Inner Offset: 0"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inner Radius: 0"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Outer Radius: 0"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
PenColor = RGB(255, 255, 255)
SStop = False
DrawRoullette
End Sub

Private Sub Command2_Click()
Picture1.Cls
End Sub

Private Sub Command3_Click()
SStop = True
End Sub

Private Sub Form_Click()
Picture1.Cls

End Sub

Private Sub Form_Load()

MsgBox "The speed of this program is dependent on the speed of your computer. Please be patient if your computer is slow. Thank you!", vbExclamation, "ZATRiX:-:ZATRiX@load.com"
SStop = False
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "Outer Radius: " & HScroll1.Value

End Sub

Private Sub HScroll2_Change()
Label2.Caption = "Inner Radius: " & HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Label3.Caption = "Inner Offset: " & HScroll3.Value
End Sub

Private Sub Option1_Click()
re = 255
gr = 0
bl = 0
End Sub

Private Sub Option2_Click()
re = 0
gr = 0
bl = 255
End Sub

Private Sub Option3_Click()
re = 0
gr = 255
bl = 0
End Sub

Private Sub Option4_Click()
re = 255
gr = 255
bl = 0
End Sub

Private Sub Option5_Click()
re = 255
gr = 255
bl = 255
End Sub

Private Sub Option6_Click()
re = 0
gr = 0
bl = 0
End Sub

Private Sub Option7_Click()
Picture1.DrawWidth = Line1.BorderWidth
End Sub

Private Sub Option8_Click()
Picture1.DrawWidth = Line2.BorderWidth
End Sub

Private Sub Option9_Click()
Picture1.DrawWidth = Line3.BorderWidth
End Sub
