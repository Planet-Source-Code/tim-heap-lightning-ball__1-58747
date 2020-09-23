VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Lightning Ball"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long


Dim blnRunning As Boolean   'If the game is running
Dim blnPaused As Boolean    'If the game is Paused
Dim lngTimeOfStart As Long

    'Used in a framerate counter
Dim lngTimer As Long
Dim lngFrameRate As Long
Dim lngCounter As Long

    'Used when clicking the mouse
Dim blnMouseIsDown As Boolean
Dim MouseX As Integer
Dim MouseY As Integer

    'Used in calculatoins
Dim sngAdjacent As Single
Dim sngOpposite As Single
Dim sngAngle As Single


    'When you clikc on the form
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
blnMouseIsDown = True
End Sub

    'When you move the mouse
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseX = x
MouseY = y
End Sub

    'When you releas the mouse button
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
blnMouseIsDown = False
End Sub

    'Pressing a key
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then blnRunning = False
End Sub

    'When the progrma starts
Private Sub Form_Load()

On_Load     'Initialises the Program

Reset       'Resets the Data

Main_Loop   'Runs the game

Un_load     'Closes down the game

Unload Me
End         'Closes the program
End Sub

    'Initialises the Program
Sub On_Load()

Randomize       'Sets up the random numbers

    'Resizes the form
ResizeForm Me, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
    'Moves the form
Me.Left = 0
Me.Top = 0
    'Makes the form transparent
'See mdlTransparent for Credit
Dim CLR As Long
Dim RET As Long
CLR = RGB(0, 0, 255) 'this color is the color that will be transparent
'Set the window style to 'Layered'
RET = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
RET = RET Or WS_EX_LAYERED
SetWindowLong Me.hWnd, GWL_EXSTYLE, RET
'Set the opacity of the layered window to 128
SetLayeredWindowAttributes Me.hWnd, CLR, 0, LWA_COLORKEY
    
    'Sets up the back buffer
picBackBuffer.Width = Me.ScaleWidth
picBackBuffer.Height = Me.ScaleHeight
picBackBuffer.BackColor = CLR

    'Initialises varialbes
HalfScaleWidth = Me.ScaleWidth / 2
HalfScaleHeight = Me.ScaleHeight / 2
BallRadius = HalfScaleHeight
For a = 1 To UBound(Lightning)
    Make_lightning a
    Lightning(a).LifeSpan = Rnd * 100
    Lightning(a).CurrentLife = Rnd * Lightning(a).LifeSpan
Next a

blnRunning = True

Me.Show 'Shows the form
    'We need to do this because the code never leaves the form_load!

End Sub

    'Resets the game
Sub Reset()

End Sub
    
    '!!!!!!!!!!!!!!!!!!!!!!!!
    'BASIC STRUCTURE OF GAME!
    '!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!DO NOT TOUCH!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!
Sub Main_Loop()
    'Declares
Dim lngTimeOfLastRender As Long
    'Initialise Variables
lngTimeOfLastRender = GetTickCount
lngTimeOfStart = lngTimeOfLastRender
    'Game Loop
Do While blnRunning 'Runs while blnRunning = true


        'Goes into here every so many milliseconds
            'eg. If Delay_Time = 25, go into it every 25 ms
    If lngTimeOfLastRender + DELAY_TIME < GetTickCount Then
        lngTimeOfLastRender = GetTickCount
            'If the games not paused
            
        If Not blnPaused Then Move_Stuff 'Call Subroutines
        Draw_Stuff  'Call Subroutines
                
        Dim lngTemp As Long
        lngTemp = GetTickCount - lngTimeOfLastRender
        SleepEx CLng(IIf(DELAY_TIME - lngTemp - 2 >= 0, DELAY_TIME - lngTemp - 2, 0)), False
    End If  'lngTOLR + Delay_Time < GetTickCount

DoEvents    'Lets other applications run

Loop    'While blnRunning
End Sub 'Main_Loop

Sub Move_Stuff()
Move_Lightning
End Sub



    'Resets spawns and moves te lightning
Sub Move_Lightning()
For a = 1 To UBound(Lightning)
        'Counts the life up
    Lightning(a).CurrentLife = Lightning(a).CurrentLife + 1
        'Respawns the lightning if its 'dead'
    If Lightning(a).CurrentLife >= Lightning(a).LifeSpan Then
        Make_lightning a
    Else
        'makes the lighting 'wriggle'
        With Lightning(a)
            For i = 1 To 10
                    'Moves all the secitons by a random amount
                .Section(i).x = .Section(i).x + ((Rnd * 5) - 2.5)
                .Section(i).y = .Section(i).y + ((Rnd * 5) - 2.5)
                .SubSection(i).x = .SubSection(i).x + ((Rnd * 5) - 2.5)
                .SubSection(i).y = .SubSection(i).y + ((Rnd * 5) - 2.5)
                .SubSubSection(i).x = .SubSubSection(i).x + ((Rnd * 5) - 2.5)
                .SubSubSection(i).y = .SubSubSection(i).y + ((Rnd * 5) - 2.5)
            Next i
        End With
    End If
Next a
End Sub


Sub Make_lightning(LightningNumber As Integer)
    'Spawns the lightning
With Lightning(LightningNumber)
        'sets its start in the middle
    .Start.x = HalfScaleWidth
    .Start.y = HalfScaleHeight

        'Sets the end at a random angle
    If blnMouseIsDown Then
        
        If Int((Rnd * 2) + 1) = 2 Then
                'Calculates the angle to the mouse
            sngAdjacent = (HalfScaleWidth) - MouseX
            sngOpposite = (HalfScaleHeight) - MouseY
            If sngOpposite = 0 Then sngOpposite = 1
            sngAngle = Atn(sngAdjacent / sngOpposite)
            If sngOpposite < 0 Then sngAngle = sngAngle + Deg2Rad(180)
                'Sets the end at the mouse
            .End.x = -(Sin(sngAngle) * (BallRadius - 10)) + .Start.x
            .End.y = -(Cos(sngAngle) * (BallRadius - 10)) + .Start.y
           .AttachedToMouse = True
        
        Else
            GoTo NormalLightning
        End If
    Else
    
NormalLightning:
            'Makes normal, random lightning
        Dim sngTemp As Single
        sngTemp = Deg2Rad(Rnd * 360)
        .End.x = (Sin(sngTemp) * (BallRadius - 10)) + .Start.x
        .End.y = (Cos(sngTemp) * (BallRadius - 10)) + .Start.y
        .AttachedToMouse = False

    End If
    
        'Sets the sections between the start and end, with a small randomness
    For i = 1 To 10
        .Section(i).x = ((.End.x - .Start.x) / 12 * i) + ((Rnd * (LightningVariation / 2)) - (LightningVariation / 4)) + .Start.x
        .Section(i).y = ((.End.y - .Start.y) / 12 * i) + ((Rnd * (LightningVariation / 2)) - (LightningVariation / 4)) + .Start.y
    Next i
        'Attaches the little sparks to the lightning
    For i = 1 To 10
        .SubSectionNumber(i) = Int(10 * Rnd + 1)
        .SubSection(i).x = .Section(.SubSectionNumber(i)).x + ((Rnd * LightningVariation) - (LightningVariation / 2))
        .SubSection(i).y = .Section(.SubSectionNumber(i)).y + ((Rnd * LightningVariation) - (LightningVariation / 2))
    Next i
    For i = 1 To 10
        .SubSubSectionNumber(i) = Int(10 * Rnd + 1)
        .SubSubSection(i).x = .SubSection(.SubSubSectionNumber(i)).x + ((Rnd * LightningVariation) - (LightningVariation / 2))
        .SubSubSection(i).y = .SubSection(.SubSubSectionNumber(i)).y + ((Rnd * LightningVariation) - (LightningVariation / 2))
    Next i
        'Sets its life
    .LifeSpan = 10 + (Rnd * 50)
    .CurrentLife = 0
End With
End Sub

    'Draws the whole thing
Sub Draw_Stuff()
    
    'Clears the back buffer
picBackBuffer.Cls

    Draw_Ball   'Draws the big ball
    
    Draw_Lightning  'Drawes the lightning

Me.Picture = picBackBuffer.Image
End Sub

Sub Draw_Ball()
picBackBuffer.FillStyle = 0
For i = 1 To 20
    Dim intColour As Integer
    intColour = 255 - ((255 / 20) * i)
    picBackBuffer.FillColor = RGB(0, intColour, intColour)
    picBackBuffer.Circle (HalfScaleWidth, HalfScaleHeight), BallRadius - i, picBackBuffer.FillColor
Next i
picBackBuffer.FillStyle = 1
End Sub

Sub Draw_Lightning()
picBackBuffer.FillStyle = 0
Dim ColourR As Integer
Dim ColourG As Integer
Dim ColourB As Integer
ColourR = 150
ColourG = 0
ColourB = 150
'For i = 1 To 5
'    picBackBuffer.FillColor = RGB((255 / 5) * i, (255 / 5) * i, (255 / 5) * i)
'
'    picBackBuffer.Circle (HalfScaleWidth, HalfScaleheight), (5 - i) * 3, picBackBuffer.FillColor
'Next i
For a = 1 To UBound(Lightning)
    With Lightning(a)
        'If .CurrentLife < 10 Then
        
            picBackBuffer.FillStyle = 0
            For i = 1 To 5
                picBackBuffer.FillColor = RGB((255 / 5) * i, (255 / 5) * i, (255 / 5) * i)
                
                picBackBuffer.Circle (.End.x, .End.y), (5 - i) * 3, picBackBuffer.FillColor
            Next i
            picBackBuffer.FillStyle = 1
        
            If .AttachedToMouse Then
                picBackBuffer.DrawWidth = 10
            Else
                picBackBuffer.DrawWidth = 5
            End If
            picBackBuffer.ForeColor = RGB(ColourR / 3, ColourG / 3, ColourB / 3)
            
            picBackBuffer.Line (.Start.x, .Start.y)-(.Section(1).x, .Section(1).y)
            For i = 1 To 9
                picBackBuffer.Line (.Section(i).x, .Section(i).y)-(.Section(i + 1).x, .Section(i + 1).y)
            Next i
            picBackBuffer.Line (.Section(10).x, .Section(10).y)-(.End.x, .End.y)
            
            
            If .AttachedToMouse Then
                picBackBuffer.DrawWidth = 4
            Else
                picBackBuffer.DrawWidth = 2
            End If
            picBackBuffer.ForeColor = RGB(ColourR, ColourG, ColourB)
            
            picBackBuffer.Line (.Start.x, .Start.y)-(.Section(1).x, .Section(1).y)
            For i = 1 To 9
                picBackBuffer.Line (.Section(i).x, .Section(i).y)-(.Section(i + 1).x, .Section(i + 1).y)
            Next i
            picBackBuffer.Line (.Section(10).x, .Section(10).y)-(.End.x, .End.y)
            
            If .AttachedToMouse Then
                picBackBuffer.DrawWidth = 2
            Else
                picBackBuffer.DrawWidth = 1
            End If
            picBackBuffer.ForeColor = RGB(ColourR / 2, ColourG / 2, ColourB / 2)
            For i = 1 To 10
                picBackBuffer.Line (.SubSection(i).x, .SubSection(i).y)-(.Section(.SubSectionNumber(i)).x, .Section(.SubSectionNumber(i)).y)
            Next i
            
            picBackBuffer.ForeColor = RGB(ColourR / 3, ColourG / 3, ColourB / 3)
            For i = 1 To 10
                picBackBuffer.Line (.SubSubSection(i).x, .SubSubSection(i).y)-(.SubSection(.SubSubSectionNumber(i)).x, .SubSection(.SubSubSectionNumber(i)).y)
            Next i
                    
        'End If
    End With
Next a
End Sub


    'Unloads everything
Sub Un_load()
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnRunning = False
End Sub

