Attribute VB_Name = "PaintMod"
Public intDrawWidth As Integer
Public intStyle As Integer
Public intRed As Integer
Public intBlue As Integer
Public intGreen As Integer
Public intPaint As Integer
Public intPaintRed As Integer
Public intPaintBlue As Integer
Public intPaintGreen As Integer
Public blnErase As Boolean
Public blnPaintBuck As Boolean
Public blnDraw As Boolean
Public blnVertM As Boolean
Public blnHorM As Boolean
Public blnBothM As Boolean
Public oldx%
Public oldy%
Public CurrentChoice
Public x1%, x2%, y1%, y2%
Public blnRainboPaint As Boolean
Public blnLeft As Boolean
Public blnRight As Boolean
Public BCRed As Integer
Public BCBlue As Integer
Public BCGreen As Integer
Public SamBlue As Integer
Public SamRed As Integer
Public SamGreen As Integer
Public blnCav1 As Boolean
Public blnCav2 As Boolean
Public blnCav3 As Boolean
Public blnCav4 As Boolean

Sub MirrorDraw(x As Single, y As Single, old1 As Integer, old2 As Integer, mirror As Byte)
    
    With frmPaint.picMainPic

        
        'Vertical
        If blnVertM Then
                .picMainPic.Line (.picMainPic.ScaleWidth - old1, old2)-(.picMainPic.ScaleWidth - x, y), RGB(intRed, intGreen, intBlue)
        End If
    
        'Horisontal
        If blnHorM Then
               .picMainPic.Line (old1, .picMainPic.ScaleHeight - old2)-(x, .picMainPic.ScaleWidth - y), RGB(intRed, intGreen, intBlue)
        End If
        
        'Diagonal
        If blnBothM Then
            .picMainPic.Line (.picMainPic.ScaleWidth - old1, .picMainPic.ScaleHeight - old2)-(.picMainPic.ScaleWidth - x, .picMainPic.ScaleWidth - y), RGB(intRed, intGreen, intBlue)
        End If
    End With

End Sub

Sub RainboPaint()
    
    Dim color, value As Single
    Dim colorchange As Integer
    
    Randomize
    color = Int(Rnd * 3)

    Randomize
    value = Int(Rnd * 2)
    
    Randomize
    colorchange = Int(Rnd * 20)
                
    If color = 1 And value = 0 And intRed + colorchange <= 255 Then intRed = intRed + colorchange
 
    If color = 1 And value = 1 And intRed - colorchange >= 0 Then intRed = intRed - colorchange

    If color = 0 And value = 0 And intGreen + colorchange <= 255 Then intGreen = intGreen + colorchange

    If color = 0 And value = 1 And intGreen - colorchange >= 0 Then intGreen = intGreen - colorchange

    If color = 2 And value = 0 And intBlue + colorchange <= 255 Then intBlue = intBlue + colorchange

    If color = 2 And value = 1 And intBlue - colorchange >= 0 Then intBlue = intBlue - colorchange
    
End Sub

