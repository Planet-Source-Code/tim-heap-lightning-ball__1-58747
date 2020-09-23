Attribute VB_Name = "mdlDeclares"
Public Const DELAY_TIME As Integer = 20 'Used in the loop


Public Const PI As Single = 3.1419254
Public Const sng180divPI = 180 / PI
Public Const sngPIdiv180 = PI / 180

    'The offset of the lightning bits
Public Const LightningVariation As Integer = 100
    
'Type Statments
    
Private Type XY
    x As Single
    y As Single
End Type

Public Type Lightnings
    Start As XY                                 'position of the Start
    End As XY                                   'Position of the End
    Section(1 To 10) As XY                      'Position of the middle bits
    SubSection(1 To 10) As XY                   'the larger sparks
    SubSectionNumber(1 To 10) As Integer        'which section its attached to
    SubSubSection(1 To 10) As XY                'the little sparks
    SubSubSectionNumber(1 To 10) As Integer     'which subsection its on
    LifeSpan As Integer                         'How long it lasts for
    CurrentLife As Integer                      'How long its lasted for
    AttachedToMouse As Boolean                  'If its attached to the mouse
End Type

Public Lightning(1 To 10) As Lightnings

'Declares
    'For...Next loops
Public i As Integer, a As Integer, b As Integer
    'the size of the Ball
Public BallRadius As Integer

Public HalfScaleWidth As Integer
Public HalfScaleHeight As Integer
