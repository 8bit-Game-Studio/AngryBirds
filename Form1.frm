VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "愤怒的小鸟"
   ClientHeight    =   8580
   ClientLeft      =   4650
   ClientTop       =   1005
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   11910
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   735
      Left            =   10560
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   10080
      Top             =   8040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   10560
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   10560
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10680
      Top             =   8040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replay"
      Height          =   735
      Left            =   10560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11280
      Top             =   8040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim SlingX As Integer
Dim SlingY As Integer
Dim BirdX As Integer
Dim BirdY As Integer
Dim TempX As Integer
Dim TempY As Integer
Dim MoveX As Integer
Dim MoveY As Integer
Dim Mark As Integer
Dim R As Integer
Dim Roll As Integer
Dim Time As Integer
Dim Vx As Single
Dim Vy As Single
Dim G As Long
Dim V As Long
Dim Bate As Integer
Dim Cloud1 As Integer
Dim Cloud2 As Integer
Dim Cloud3 As Integer


Private Sub Command1_Click()
    Form1.Width = 800 * Screen.TwipsPerPixelX
    Form1.Height = 600 * Screen.TwipsPerPixelY
    SlingX = 100
    SlingY = 431
    TempX = 0
    TempY = 0
    MoveX = 0
    MoveY = 0
    Mark = 0
    R = 1100
    Roll = 0
    Time = 0
    Bate = 100
    Vx = 0
    Vy = 0
    V = 0
    G = 1
End Sub


Private Sub Form_Load()
    SlingX = 100
    SlingY = 431
    TempX = 0
    TempY = 0
    MoveX = 0
    MoveY = 0
    Mark = 0
    R = 1100
    Roll = 0
    Time = 0
    Bate = 100
    Vx = 0
    Vy = 0
    V = 0
    G = 1
    Cloud1 = 1
    Cloud2 = 10
    Cloud3 = 600
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub Form_Activate()
    Call PaintPng(App.Path & "\Fire.png", Me.hdc, SlingX - 5, SlingY - 7)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1_Click
    If Mark = 0 And Button = 1 And X / Screen.TwipsPerPixelX > SlingX - 5 And X / Screen.TwipsPerPixelX < SlingX + 25 And Y / Screen.TwipsPerPixelY > SlingY - 5 And Y / Screen.TwipsPerPixelY < SlingY + 25 Then
        Mark = 1
        MoveX = X
        MoveY = Y
        BirdX = X / Screen.TwipsPerPixelX - 18
        BirdY = Y / Screen.TwipsPerPixelY - 18
    End If
    Command1.Caption = Down
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mark = 1 Then
        If (X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 11) * Screen.TwipsPerPixelY) ^ 2 < R ^ 2 Then
            BirdX = X / Screen.TwipsPerPixelX - 18
            BirdY = Y / Screen.TwipsPerPixelY - 18
            MoveX = X
            MoveY = Y
        Else
            BirdX = ((SlingX + 13) * Screen.TwipsPerPixelX + R / Sqr((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 13) * Screen.TwipsPerPixelY) ^ 2) * (X - (SlingX + 13) * Screen.TwipsPerPixelX)) / Screen.TwipsPerPixelX - 18
            BirdY = ((SlingY + 11) * Screen.TwipsPerPixelY + R / Sqr((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 13) * Screen.TwipsPerPixelY) ^ 2) * (Y - (SlingY + 13) * Screen.TwipsPerPixelY)) / Screen.TwipsPerPixelY - 18
            MoveX = (SlingX + 13) * Screen.TwipsPerPixelX + R / Sqr((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 13) * Screen.TwipsPerPixelY) ^ 2) * (X - (SlingX + 13) * Screen.TwipsPerPixelX)
            MoveY = (SlingY + 11) * Screen.TwipsPerPixelY + R / Sqr((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 13) * Screen.TwipsPerPixelY) ^ 2) * (Y - (SlingY + 13) * Screen.TwipsPerPixelY)
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mark = 1 Then
        BirdX = SlingX - 5
        BirdY = SlingY - 7
        Cls
        Call PaintPng(App.Path & "\Fire.png", Me.hdc, SlingX - 5, SlingY - 7)
        If (X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 11) * Screen.TwipsPerPixelY) ^ 2 < R ^ 2 Then
            V = ((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 11) * Screen.TwipsPerPixelY) ^ 2) / 1000
        Else
            V = (R ^ 2) / 1000
        End If
        Vx = (-V / Sqr((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 13) * Screen.TwipsPerPixelY) ^ 2) * (X - (SlingX + 13) * Screen.TwipsPerPixelX)) / Bate
        Vy = (-V / Sqr((X - (SlingX + 13) * Screen.TwipsPerPixelX) ^ 2 + (Y - (SlingY + 13) * Screen.TwipsPerPixelY) ^ 2) * (Y - (SlingY + 13) * Screen.TwipsPerPixelY)) / Bate
        Mark = 2
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
   Command4.Caption = Mark
    If Mark = 0 Then
        Cls
        Call PaintPng(App.Path & "\Fire.png", Me.hdc, SlingX - 5, SlingY - 7)
        Call PaintPng(App.Path & "\Cloud1.png", Me.hdc, Cloud1, 10)
        Call PaintPng(App.Path & "\Cloud2.png", Me.hdc, Cloud2, 15)
        Call PaintPng(App.Path & "\Cloud3.png", Me.hdc, Cloud3, 17)
    Else
        If Mark = 1 Then
            Cls
            Form1.Line ((SlingX + 25) * Screen.TwipsPerPixelX, (SlingY + 13) * Screen.TwipsPerPixelY)-(MoveX, MoveY)
            Call PaintPng(App.Path & "\Fire.png", Me.hdc, BirdX, BirdY)
            Form1.Line ((SlingX + 5) * Screen.TwipsPerPixelX, (SlingY + 10) * Screen.TwipsPerPixelY)-(MoveX, MoveY)
            Call PaintPng(App.Path & "\Cloud1.png", Me.hdc, Cloud1, 10)
            Call PaintPng(App.Path & "\Cloud2.png", Me.hdc, Cloud2, 15)
            Call PaintPng(App.Path & "\Cloud3.png", Me.hdc, Cloud3, 17)
        Else
            If Mark = 2 Then
                Cls
                Call PaintPng(App.Path & "\Fire.png", Me.hdc, BirdX, BirdY)
                Call PaintPng(App.Path & "\Cloud1.png", Me.hdc, Cloud1, 10)
                Call PaintPng(App.Path & "\Cloud2.png", Me.hdc, Cloud2, 15)
                Call PaintPng(App.Path & "\Cloud3.png", Me.hdc, Cloud3, 17)
            Else
                If Mark = 3 Or Mark = 4 Then
                    Select Case ((BirdX / 10) Mod 8)
                    Case 0
                        Cls
                        Call PaintPng(App.Path & "\Fire.png", Me.hdc, BirdX, BirdY)
                    Case 1
                        Cls
                        Call PaintPng(App.Path & "\Fire1.png", Me.hdc, BirdX, BirdY)
                    Case 2
                        Cls
                        Call PaintPng(App.Path & "\Fire2.png", Me.hdc, BirdX, BirdY)
                    Case 3
                        Cls
                        Call PaintPng(App.Path & "\Fire3.png", Me.hdc, BirdX, BirdY)
                    Case 4
                        Cls
                        Call PaintPng(App.Path & "\Fire4.png", Me.hdc, BirdX, BirdY)
                    Case 5
                        Cls
                        Call PaintPng(App.Path & "\Fire5.png", Me.hdc, BirdX, BirdY)
                    Case 6
                        Cls
                        Call PaintPng(App.Path & "\Fire6.png", Me.hdc, BirdX, BirdY)
                    Case 7
                        Cls
                        Call PaintPng(App.Path & "\Fire7.png", Me.hdc, BirdX, BirdY)
                    End Select
                    Call PaintPng(App.Path & "\Cloud1.png", Me.hdc, Cloud1, 10)
                    Call PaintPng(App.Path & "\Cloud2.png", Me.hdc, Cloud2, 15)
                    Call PaintPng(App.Path & "\Cloud3.png", Me.hdc, Cloud3, 17)
                End If
            End If
        End If
    End If
    Command4.Caption = Mark
End Sub

Private Sub Timer2_Timer()
    If Mark = 2 Then
        Time = Time + 1
        BirdX = ((Time + 1) * Vx + (SlingX - 5))
        BirdY = ((Time + 1) * Vy + G * (Time ^ 2) / 10 + (SlingY - 7))
        If BirdY > 480 Then
            Roll = 0
            Mark = 3
        Else
            If BirdX > 810 Or BirdX < -60 Or BirdY < 0 Then
             '   Timer1.Enabled = False
                Mark = 4
             '   Timer2.Enabled = False
             '   Command1_Click
            End If
        End If
    Else
        If Mark = 3 Then
            If Roll < 40 Then
                If Vx > 0.2 Then
                    Vx = Vx - 0.19
                End If
                BirdX = BirdX + Vx
            Else
             '   Timer1.Enabled = False
                Mark = 4
            '    Timer2.Enabled = False
            '    Command1_Click
            End If
            Roll = Roll + 1
        End If
    End If
    If Cloud1 < 700 Then
        Cloud1 = Cloud1 + 1
    Else
        Cloud1 = -950
    End If
    If Cloud2 < 780 Then
        Cloud2 = Cloud2 + 2
    Else
        Cloud2 = -120
    End If
    If Cloud3 < 780 Then
        Cloud3 = Cloud3 + 1
    Else
        Cloud3 = -90
    End If
    Command2.Caption = BirdX
    Command3.Caption = BirdY
End Sub
