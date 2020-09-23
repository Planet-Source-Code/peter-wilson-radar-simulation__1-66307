VERSION 5.00
Begin VB.Form frmRadar 
   Caption         =   "RADAR"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   Icon            =   "frmRadar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Animation 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmRadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' How many aircraft do we want?
Private Const m_MaxAircraft = 19

' *** Important ***
' You must balance the speed of rotation with the accuracy of the beam.
Private Const m_sngRadarSpeed = 5             '   Any value from -179 to + 179. (Try values between +1 to +10)
Private Const m_sngRadarAccuracy = 0.999       '   Any value from -1 to +1 (realistic radar beam = +0.9999)



Private Const m_Radius = 100

' The zoom value is just for fun. It's not really needed, but it's so simple to do I usually put it in every graphics project anyway.
Private m_sngZoomRequired   As Single
Private m_sngZoomActual     As Single

Private m_sngRadarAngle As Single

Private Sub DoDrawCircle(StepSize As Single, Radius As Single, Color As OLE_COLOR, Optional XPosition As Single, Optional YPosition As Single)

    Dim sngDegrees  As Single
    Dim sngRadians  As Single
    Dim sngX        As Single
    Dim sngY        As Single
    
    For sngDegrees = 45 To 360 + 45 Step StepSize
    
        ' Convert degrees to radians
        sngRadians = sngDegrees * 0.01745329
        
        sngX = Radius * Sin(sngRadians) + XPosition
        sngY = Radius * -Cos(sngRadians) + YPosition
        
        If sngDegrees = 45 Then
            Me.CurrentX = sngX
            Me.CurrentY = sngY
        Else
            Me.Line -(sngX, sngY), Color
        End If
        
    Next sngDegrees
        
End Sub

Private Sub DoDrawCrossHairs(StepSize As Single, Radius As Single, LineColor As OLE_COLOR, TextColor As OLE_COLOR)

    Dim sngDegrees  As Single
    Dim sngRadians  As Single
    Dim sngX        As Single
    Dim sngY        As Single
    
    For sngDegrees = 0 To (360 - StepSize) Step StepSize
    
        ' Convert degrees to radians
        sngRadians = sngDegrees * 0.01745329
        
        sngX = Radius * Sin(sngRadians)
        sngY = Radius * -Cos(sngRadians)
        
        Me.Line (0, 0)-(sngX, sngY), LineColor
        Me.ForeColor = TextColor
        Me.Print sngDegrees
        
    Next sngDegrees

End Sub

Private Sub DoDrawRadarBeam(Radius As Single, LineColor As OLE_COLOR)

    Dim sngRadians As Single
    Dim sngX As Single
    Dim sngY As Single
    
    ' Convert degrees to radians
    sngRadians = m_sngRadarAngle * 0.01745329
    
    sngX = Radius * Sin(sngRadians)
    sngY = Radius * -Cos(sngRadians)
    
    Me.Line (0, 0)-(sngX, sngY), LineColor
    
End Sub


Private Function DotProduct(VectorU As virtualAircraft, VectorV As virtualAircraft) As Single

    ' Determines the dot-product of two vectors.
    DotProduct = (VectorU.XPosition * VectorV.XPosition) + (VectorU.YPosition * VectorV.YPosition)
    
End Function


Private Sub DoDrawRadarReturns(Radius As Single)

    ' Plan of action
    ' --------------
    '
    '   1) Loop through all aircraft objects.
    '   2) Calculate the dot-product between the beam location and the aircraft's location.
    '
    '      The "dot-product" is THE most useful thing you can learn. Once you understand how
    '      easy and simple it is, you'll be using it everywhere!!!
    '
    '      Any time you need to know the difference between two angles (or two object positions)
    '      in 3D or 2D, you will use the 'dot-product'.
    '
    '   3) If the "dot-product" is close to zero, then display the aircraft on the radar.
    
    Dim sngRadians As Single
    Dim beamVector As virtualAircraft   ' I'm reusing this object for the "radar beam's location"
    Dim sngDotProduct As Single
    Dim intIndex As Integer
    
    ' Convert degrees to radians
    sngRadians = m_sngRadarAngle * 0.01745329
    beamVector.XPosition = Radius * Sin(sngRadians)
    beamVector.YPosition = Radius * -Cos(sngRadians)
    

    ' Loop through all aircraft and move them around.
    For intIndex = 0 To m_MaxAircraft
            
        ' I use the 'With' keyword to save on typing.
        With myVirtualAircraftArray(intIndex)
        
            ' Get the cosine of the angle between the radar beam and the aircraft.
            sngDotProduct = DotProduct(Vec3Normalize(beamVector), Vec3Normalize(myVirtualAircraftArray(intIndex)))
                    
            ' TODO: This is a *nice* one to adjust. Try any value from -1 to +1
            ' For a *VERY* narrow radar beam use 0.99999
            ' Just like a real radar unit, if you make the beam very narrow, you MUST also slow down it's rotation speed.
            If (sngDotProduct) > m_sngRadarAccuracy Then
                
                ' Aircraft has been hit by the radar.
                .RadarEnergy = 255
                
                ' Remember the position that the radar last hit the aircraft.
                .XHitPosition = .XPosition
                .YHitPosition = .YPosition
                
            End If
    
            ' Display the last known position of the aircraft.
            Me.PSet (.XHitPosition, .YHitPosition), RGB(.RadarEnergy, .RadarEnergy, 0)

            Call DoDrawCircle(120, 2, RGB(.RadarEnergy, .RadarEnergy, 0), .XHitPosition, .YHitPosition)
            
            ' TODO: Tweak this value so that the energy fades.
            Dim sngFadeAmount As Single
            sngFadeAmount = 190 / (360 / m_sngRadarSpeed) ' The 190 should be 255 (to match the initial radar energy. ie. 255)
            .RadarEnergy = .RadarEnergy - sngFadeAmount
            
            If .RadarEnergy < 0 Then .RadarEnergy = 0
            
        End With

    Next intIndex
    
    
End Sub

Private Sub DoResetAircraftPositions()

    Dim intIndex As Integer
    
    ' Create 20 virtual aircraft objects in memory (0 to 19 = 20)
    ReDim myVirtualAircraftArray(m_MaxAircraft)

    ' Loop through all aircraft and set some properties.
    For intIndex = 0 To m_MaxAircraft
        
        ' Randomly position each aircraft within plus or minus 500 units.
        myVirtualAircraftArray(intIndex).XPosition = (Rnd * m_Radius) - (m_Radius / 2)
        myVirtualAircraftArray(intIndex).YPosition = (Rnd * m_Radius) - (m_Radius / 2)
        
        ' Randomly adjust each aircraft's velocity within plus or minus 5 units.
        myVirtualAircraftArray(intIndex).XVelocity = (Rnd * (m_Radius / 1000)) - (m_Radius / 2000)
        myVirtualAircraftArray(intIndex).YVelocity = (Rnd * (m_Radius / 1000)) - (m_Radius / 2000)
        
        ' Reset the amount of "radar energy" that the aircraft has "absorbed"
        ' Ok... this is a quick and nasty hack to make the aircraft "fade away" on the radar.
        myVirtualAircraftArray(intIndex).RadarEnergy = 0
    Next intIndex
    
End Sub

Private Sub DoAircraftSimulation()

    Dim intIndex As Integer

    ' Loop through all aircraft and move them around.
    For intIndex = 0 To m_MaxAircraft
    
        ' I use the with statement to save on typing, that is all.
        With myVirtualAircraftArray(intIndex)
        
            ' Move the aircraft according to it's velocity/heading.
            .XPosition = .XPosition + .XVelocity
            .YPosition = .YPosition + .YVelocity
        
            ' If aircraft goes outside our area, just bounce it back in (very unrealistic!!)
            If Abs(.XPosition) > m_Radius Then .XVelocity = -.XVelocity
            If Abs(.YPosition) > m_Radius Then .YVelocity = -.YVelocity
            
        End With
    
    Next intIndex
    
End Sub

Private Sub DoDrawRadar()

    Call DoDrawCircle(5, m_Radius, RGB(255, 0, 0))
    Call DoDrawCircle(10, m_Radius * 0.75, RGB(128, 0, 0))
    Call DoDrawCircle(18, m_Radius * 0.5, RGB(64, 0, 0))
    
    Call DoDrawCrossHairs(45, 105, RGB(92, 0, 0), RGB(255, 128, 128))
    
End Sub

Private Sub DoZoomWindow()
        
    ' =================================================================================================
    ' There's 2 ways to make an object appear bigger on the screen.
    '   1) Make the object/geometry bigger.
    '       or
    '   2) Leave the geometry the same size, but make the coordinate system smaller.
    '
    ' Most professional games don't adjust the geometry at all, instead they change the
    ' coordinate system by an opposite amount. That is what I have done in this application
    ' by chaning the VB form's ScaleWidth and ScaleHeight properties.
    '
    ' Not only can you zoom using this method, but you can also scroll and pan around your
    ' world, and the only things you have to adjust are the Scale* properties of the form.
    ' Pretty simply hey? Now, get cracking and build a 2D Role Playing Game with scrolling/zooming Map!
    ' It should take you about 10-30 lines of code for the pan/scroll/zoom.
    ' Key Press -> Update the Required/Actual Values -> Resize the coordinate system.
    ' =================================================================================================
    
    ' Calculate the difference between what the user wants, and what the current zoom value.
    Dim sngDifference As Single
    sngDifference = (m_sngZoomRequired - m_sngZoomActual)
    
    ' Only perform an adjustment if the two values are different.
    If sngDifference <> 0 Then
    
        ' Adjust the actual zoom value by a fraction of the difference.
        m_sngZoomActual = m_sngZoomActual + (sngDifference / 8)
        
        ' This step isn't really required. It just stops the adjustments when they are really close to each other.
        If (Abs(m_sngZoomRequired - m_sngZoomActual) < 1) Then m_sngZoomActual = m_sngZoomRequired
        
        ' Resize the coordinate system.
        Call Form_Resize
    End If

End Sub

Private Sub Animation_Timer()

    ' This timer routine should get called as often as possible.
    '
    ' These comments here, are the first things I wrote for the application.
    ' Once I've got my plan, then I start.
    '
    '   Step 1a)    Clear the screen
    '   Step 1b)    Adjust the zoom value to match what the user requests.
    '   Step 2)     Draw the radar screen
    '   Step 3)     Draw the radar's current beam location
    '   Step 4)     Draw the aircraft within the radar beam (will have to do some simple calculations for this)
    '   Step 5)     Move the virtual aircraft around (basically a very simple simulation of aircraft movement)
    
    Me.Cls
    
    Call DoZoomWindow
    
    Call DoDrawRadar
    
    m_sngRadarAngle = m_sngRadarAngle + m_sngRadarSpeed
    If m_sngRadarAngle > 360 Then m_sngRadarAngle = m_sngRadarAngle - 360
    
    Call DoDrawRadarBeam(m_Radius, RGB(0, 255, 0))

    Call DoDrawRadarReturns(m_Radius)

    Call DoAircraftSimulation
    
End Sub

Private Sub Form_DblClick()
    If Me.WindowState <> vbNormal Then
        Me.WindowState = vbNormal
    End If
    Me.Height = Me.Width
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then Unload Me
    If (KeyCode = vbKeyUp) Then m_sngZoomRequired = m_sngZoomRequired / 1.1
    If (KeyCode = vbKeyDown) Then m_sngZoomRequired = m_sngZoomRequired * 1.1

End Sub


Private Sub Form_Load()

    ' Reset the zoom values. This will determin the height and width of the drawing area.
    m_sngZoomActual = 150
    m_sngZoomRequired = 150
    
    ' Background color to black
    Me.BackColor = vbBlack
    
    ' Prevent flicker.
    Me.AutoRedraw = True
    
    Call DoResetAircraftPositions
    
End Sub

Private Sub Form_Resize()
   
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Get the aspect ratio of the form (the user might have resized it)
    Dim sngAspectRatio As Single
    sngAspectRatio = Me.Width / Me.Height
    
    ' Reset the height and width of the form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier.
    Me.ScaleWidth = 2 * m_sngZoomActual * sngAspectRatio
    Me.ScaleLeft = -ScaleWidth / 2
    
    Me.ScaleHeight = 2 * m_sngZoomActual
    Me.ScaleTop = -Me.ScaleHeight / 2
    
End Sub

