Attribute VB_Name = "mMaths"
Option Explicit

' This is a virtual aircraft object.
Public Type virtualAircraft
    
    ' This is the simulated position of the aircraft.
    XPosition   As Single
    YPosition   As Single
       
    ' This is the aircraft's direction and speed vector.
    XVelocity   As Single
    YVelocity   As Single
    
    RadarEnergy As Single   ' This is a quick hack to made the signal fade away... actually, it works quite nicely!
    
    ' This is the last known position of the aircraft when it was hit by the radar.
    XHitPosition   As Single
    YHitPosition   As Single
    
End Type

' This empty array will hold our virtual aircraft.
Public myVirtualAircraftArray() As virtualAircraft

Public Function Vec3Length(V1 As virtualAircraft) As Single

    ' Returns the length of a vector.
    Vec3Length = Sqr((V1.XPosition ^ 2) + (V1.YPosition ^ 2))
    
End Function


' Returns the normalized version of a vector.
Public Function Vec3Normalize(V1 As virtualAircraft) As virtualAircraft
    
    Dim sngLength As Single
    
    sngLength = Vec3Length(V1)
    
    If sngLength = 0 Then sngLength = 1
    
    Vec3Normalize.XPosition = V1.XPosition / sngLength
    Vec3Normalize.YPosition = V1.YPosition / sngLength
    
End Function

