VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "MaxSpeed"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   2220
      TabIndex        =   0
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MaxSpeed = Not MaxSpeed
    Command1.FontBold = MaxSpeed
End Sub

Private Sub Form_Load()
    Randomize
    DoEvents
    Me.Show
    Men(1).tX = 50
    Men(1).tY = 50
    Men(1).Atr = 7
    For A = 2 To UBound(Men)
        Men(A).X = Int(Rnd * FrmMain.ScaleWidth)
        Men(A).Y = Int(Rnd * FrmMain.ScaleHeight)
        Men(A).tX = Men(A).X: Men(A).tY = Men(A).Y
        Men(A).Atr = 7
    Next A
    MainLoop
End Sub

Sub MainLoop()
Dim TickTag As Long
    Do
        DoEvents
        t = GetTickCount
        'Framelimiter
        Do Until GetTickCount > TickTag Or MaxSpeed
            DoEvents
        Loop: TickTag = GetTickCount + 10 'TickTag = IIf(Faster, GetTickCount + 25, GetTickCount + 50)
        

        
        
        DoStuffMen
        
        PaintBoard
        
        
        
        
        Label1 = GetTickCount - t
        
    Loop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Men(1).tX = X
    Men(1).tY = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
