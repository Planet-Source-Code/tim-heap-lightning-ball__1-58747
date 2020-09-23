Attribute VB_Name = "mdlModule"
Option Explicit

    'Sine, cos and tan functions
Function Sine(AngleInDegrees As Single) As Single
Sine = Sin(Deg2Rad(AngleInDegrees))
End Function
Function Cosine(AngleInDegrees As Single) As Single
Cosine = Cos(Deg2Rad(AngleInDegrees))
End Function
Function Tangent(AngleInDegrees As Single) As Single
Tangent = Tan(Deg2Rad(AngleInDegrees))
End Function
Function ArcTangent(Opposite As Single, Adjacent As Single) As Single
If Adjacent = 0 Then
    If Opposite < 0 Then
        ArcTangent = 270
    Else
        ArcTangent = 90
    End If
Else
    ArcTangent = Rad2Deg(Atn(Opposite / Adjacent))
End If
End Function

    'Converts Degrees to radians and vice versa.
Function Deg2Rad(AngleInDegrees As Single) As Single
Deg2Rad = AngleInDegrees * sngPIdiv180
End Function
Function Rad2Deg(AngleInRadians As Single) As Single
Rad2Deg = AngleInRadians * sng180divPI
End Function
    
    'Will resize the form's scalesize to the specified numbers
Sub ResizeForm(Form As Form, WidthInPixels As Integer, HeightInPixels As Integer)
Dim FormWidth As Integer
Dim FormHeight As Integer
Dim CurrentScaleMode As Integer
Dim CurrentScaleHeight As Integer
Dim CurrentScaleWidth As Integer
CurrentScaleMode = Form.ScaleMode
If CurrentScaleMode = 0 Then
    CurrentScaleHeight = Form.ScaleHeight
    CurrentScaleWidth = Form.ScaleWidth
End If
Form.ScaleMode = 1 'Twips
Form.Width = (Form.Width - Form.ScaleWidth) + (WidthInPixels * Screen.TwipsPerPixelX)
Form.Height = (Form.Height - Form.ScaleHeight) + (HeightInPixels * Screen.TwipsPerPixelY)
Form.ScaleMode = CurrentScaleMode
If CurrentScaleMode = 0 Then
    Form.ScaleHeight = CurrentScaleHeight
    Form.ScaleWidth = CurrentScaleWidth
End If

End Sub

