VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Frames() As BitmapBuffer, Index As Long, FixedIndex As Long

Public Sub LoadFromFile(Hdc As Long, ParamArray Paths())
Dim Path As Long, Image As Long, Graphics As Long, Rect As RECTL, i As Long
ReDim Frames(UBound(Paths))
For i = 0 To UBound(Frames)
    GdipLoadImageFromFile StrPtr(Paths(i)), Image
    GdipGetImageWidth Image, Rect.Right
    GdipGetImageHeight Image, Rect.Bottom
    Set Frames(i) = New BitmapBuffer
    Frames(i).Create Hdc, Rect.Right, Rect.Bottom
    GdipCreateFromHDC Frames(i).CompatibleDC, Graphics
    GdipDrawImage Graphics, Image, 0, 0
    GdipDeleteGraphics Graphics
    GdipDisposeImage Image
Next
End Sub

Public Property Get NextFrame() As BitmapBuffer
  Index = Index + 1
 Set NextFrame = Frames(Index Mod (UBound(Frames) + 1))
  
End Property

'Public Property Get NextFrameFixed() As BitmapBuffer
'
'End Property

