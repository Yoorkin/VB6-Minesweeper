VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "骚雷               By方程"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===========游戏功能变量=======================================
Dim isBomb(-1 To 10, -1 To 10) As Boolean
Dim HasTested As Integer, BombCount As Integer
'===========绘图变量===========================================
Private Enum SheetTypeEnum
    None = -1
    Num0
    Num1
    Num2
    Num3
    Num4
    Num5
    Num6
    Num7
    Num8
    Num9
    Flag
    Bomb
End Enum

Dim Button As Integer, x As Single, y As Single
Dim SheetType(9, 9) As SheetTypeEnum
Dim Sence As New BitmapBuffer
Dim Number(11) As New FrameManager, Bg As New FrameManager, Cursor As New FrameManager



Private Sub Form_Load()
    Me.Move Me.Left, Me.Top, 500 * Screen.TwipsPerPixelX, 528 * Screen.TwipsPerPixelY
    mGdip.InitGDIPlus
    Sence.Create Me.Hdc, 500, 500
    Dim i, j
    For i = 0 To 10
        Number(i).LoadFromFile Me.Hdc, App.Path & "\Resources\" & i & ".png"
    Next
    Number(11).LoadFromFile Me.Hdc, App.Path & "\Resources\Boom.png"
    Bg.LoadFromFile Me.Hdc, App.Path & "\Resources\bg.png"
    Cursor.LoadFromFile Me.Hdc, App.Path & "\Resources\cursor.png"

    LoadGame 1
End Sub

Public Sub LoadGame(Optional GameBombCount As Byte = 10)
    HasTested = 0
    BombCount = GameBombCount
    Dim i, j
    For i = 0 To 9
        For j = 0 To 9
            SheetType(i, j) = -1
            isBomb(i, j) = False
        Next
    Next
      

    For i = 1 To BombCount
        Randomize
        x = Rnd * 9
        y = Rnd * 9
        If isBomb(x, y) Then i = i - 1 Else isBomb(x, y) = True
    Next
 
End Sub

Public Sub CheckSheet(x As Byte, y As Byte)
    Dim BombCount As Long
    If SheetType(x, y) <> -1 Then Exit Sub
    
    If isBomb(x, y) Then
       SheetType(x, y) = Bomb
       Exit Sub
    End If
    
    If isBomb(x - 1, y - 1) Then BombCount = BombCount + 1
    If isBomb(x, y - 1) Then BombCount = BombCount + 1
    If isBomb(x + 1, y - 1) Then BombCount = BombCount + 1
    If isBomb(x - 1, y) Then BombCount = BombCount + 1
    If isBomb(x + 1, y) Then BombCount = BombCount + 1
    If isBomb(x - 1, y + 1) Then BombCount = BombCount + 1
    If isBomb(x, y + 1) Then BombCount = BombCount + 1
    If isBomb(x + 1, y + 1) Then BombCount = BombCount + 1
    
    SheetType(x, y) = BombCount
    HasTested = HasTested + 1
    If 100 - HasTested = BombCount Then MsgBox "胜利"
    
    If BombCount = 0 Then
        If y < 9 Then CheckSheet x, y + 1
        If y > 0 Then CheckSheet x, y - 1
        If x > 0 Then CheckSheet x - 1, y
        If x < 9 Then CheckSheet x + 1, y
        If y > 0 And x > 0 Then CheckSheet x - 1, y - 1
        If y > 0 And x < 9 Then CheckSheet x + 1, y - 1
        If x > 0 And y < 9 Then CheckSheet x - 1, y + 1
        If x < 9 And y < 9 Then CheckSheet x + 1, y + 1
    End If


End Sub

Public Sub UpDate()
    Dim shadow As Byte
    Bg.NextFrame.Present Sence.CompatibleDC, 0, 0
    If Button = 1 Then shadow = 100 Else shadow = 255
    Dim i, j
    For i = 0 To 9
        For j = 0 To 9
            If SheetType(i, j) >= 0 Then Number(SheetType(i, j)).NextFrame.Present Sence.CompatibleDC, i * 50, j * 50
        Next
    Next
    Cursor.NextFrame.Present Sence.CompatibleDC, Int(x / 50) * 50 - 25, Int(y / 50) * 50 - 25, shadow
    Sence.Present Me.Hdc, 0, 0
End Sub

Private Sub Form_MouseDown(btn As Integer, Shift As Integer, x As Single, y As Single)
    Button = btn
    UpDate
End Sub

Private Sub Form_MouseMove(btn As Integer, Shift As Integer, mX As Single, mY As Single)
     x = mX: y = mY
    UpDate
End Sub

Private Sub Form_MouseUp(btn As Integer, Shift As Integer, x As Single, y As Single)
    Button = 0
    If btn = 1 Then CheckSheet Int(x / 50), Int(y / 50)
    If btn = 2 Then
       If SheetType(Int(x / 50), Int(y / 50)) = None Then
           SheetType(Int(x / 50), Int(y / 50)) = Flag
       ElseIf SheetType(Int(x / 50), Int(y / 50)) = Flag Then
           SheetType(Int(x / 50), Int(y / 50)) = None
       End If
    End If
    UpDate
End Sub

Private Sub Form_Paint()
    UpDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mGdip.TerminateGDIPlus
End Sub
