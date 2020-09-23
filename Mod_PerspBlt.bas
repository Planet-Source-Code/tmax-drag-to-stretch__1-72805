Attribute VB_Name = "modPerspBlt"
' Original PerspBltX & PerspBltY  from Mike D Sutton (http://edais.mvps.org)
' PerspBltX & PerspBltY - modified by Tmax to fit the application.

Option Explicit
Const COLORONCOLOR = 3
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Declare Function StretchBlt Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal nSrcWidth As Long, _
      ByVal nSrcHeight As Long, _
      ByVal dwRop As Long) As Long
      
Sub PerspBltX(ByVal outDC As Long, ByVal outX As Long, ByVal outY As Long, _
    ByVal OutWidth As Long, ByVal outStartHeight As Long, ByVal outEndHeight As Long, _
    ByVal outYOff As Long, ByVal inDC As Long, ByVal inWidth As Long, ByVal inHeight As Long)
Dim loopx As Long
Dim InterpPos As Single
Dim InterpH As Long
Dim StartLoop As Long
Dim EndLoop As Long
If OutWidth = 0 Then Exit Sub
StartLoop = 0
EndLoop = OutWidth
If OutWidth < 0 Then
    StartLoop = OutWidth
    EndLoop = 0
End If
SetStretchBltMode outDC, COLORONCOLOR
For loopx = StartLoop To EndLoop
    InterpPos = loopx / OutWidth
    InterpH = InterpPos * (outEndHeight - outStartHeight)
    StretchBlt outDC, loopx + outX, outY + (InterpPos * outYOff), 1, outStartHeight + InterpH, inDC, (InterpPos * inWidth), 0, 1, inHeight, vbSrcCopy
Next loopx
End Sub

Sub PerspBltY(ByVal outDC As Long, ByVal outX As Long, ByVal outY As Long, _
    ByVal outStartWidth As Long, ByVal outEndWidth As Long, ByVal OutHeight As Long, _
    ByVal outXOff As Long, ByVal inDC As Long, ByVal inWidth As Long, ByVal inHeight As Long)
Dim LoopY As Long
Dim InterpPos As Single
Dim InterpW As Long
Dim StartLoop As Long
Dim EndLoop As Long
If OutHeight = 0 Then Exit Sub
StartLoop = 0
EndLoop = OutHeight
If OutHeight < 0 Then
    StartLoop = OutHeight
    EndLoop = 0
End If
SetStretchBltMode outDC, COLORONCOLOR
For LoopY = StartLoop To EndLoop
    InterpPos = LoopY / OutHeight
    InterpW = InterpPos * (outEndWidth - outStartWidth)
    StretchBlt outDC, outX + (InterpPos * outXOff), LoopY + outY, outStartWidth + InterpW, 1, inDC, 0, (InterpPos * inHeight), inWidth, 1, vbSrcCopy
Next LoopY
End Sub

