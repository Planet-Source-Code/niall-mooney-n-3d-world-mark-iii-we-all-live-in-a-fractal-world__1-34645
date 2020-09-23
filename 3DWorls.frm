VERSION 5.00
Begin VB.Form frm3DWorld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ni-Star Enterprises 3DWorld Mark III - We all live in a Fractal World"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C000&
   Icon            =   "3DWorls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLandColour 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   3885
      Left            =   1560
      Picture         =   "3DWorls.frx":030A
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C000&
      Height          =   5850
      Left            =   6795
      ScaleHeight     =   386
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   7740
   End
End
Attribute VB_Name = "frm3DWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ni-Star Enterprises Agree Classes Seven
'Started - 16 March 2002
'3D World Mark III: We all live in a cubic world


Const VPW = 512 'Viewport Dimentions
Const VPH = 384
Const HVPW = 255
Const HVPH = 195
Const Pi = 3.14159265358979
Const Pi2 = 6.28318530717959
Const LandWidth = 25
Const LandLength = 25
Const LandHeight = 6.66

Private Type Point2D
    X As Long
    Y As Long
End Type

Private Type Point3D
    X As Double
    Y As Double
    Z As Double
End Type

Private Type Camera
    X As Double
    Y As Double
    Z As Double
    RotationX As Double
    RotationY As Double
    RotationZ As Double
End Type

Private Type TriIndexia
    One As Long
    Two As Long
    Three As Long
    Texture As Integer
    Colour As Long
    Lit As Long
End Type

Private Type Object3D
    Verticies As Long
    Vertex() As Point3D
    VertexProcessed() As Point3D
    VertexScreen() As Point2D
    Triangles As Long
    Triangle() As TriIndexia
    ZBuffTri() As Integer
    TriVisible() As Boolean
    WorldCoords As Point3D
    ScaleFactor As Double
    RotationX As Double
    RotationY As Double
    RotationZ As Double
End Type

Dim WorldObjects() As Object3D
Dim ViewPoint As Camera
Dim DemoGoing As Boolean
Dim Frames As Long, FramesL As Long, FramesPS As Single
Dim LastTime As Long
Dim TimeGap As Long
Dim Blur As Boolean
Dim RenderMode As Integer
Dim BackFaceRemove As Integer
Dim Outline As Boolean

Dim PlasmaLandscape(LandWidth + 2, LandLength + 2) As Long

Dim Xadd As Double, Yadd As Double, Zadd As Double

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Point2D, ByVal nCount As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlending Lib "msimg32.dll" Alias "AlphaBlend" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BF As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_Load()
Me.Show
Call Main
End Sub

Private Sub Main()
Dim LastF As Long, LastFCT As Long  'Frame/s Calculation Tick

LastFCT = GetTickCount
Frames = 1
TimeGap = 15
DemoGoing = True

Call SetupWorld(0) 'Sets up 3D World Engine

'Insert Code to Load Objects of 3D World
'Call LoadCube(WorldObjects(0), -100, 0, 100, 5)
'Call LoadTear(WorldObjects(0), -100, 100, 100, 5, 13)
'Call LoadCube(WorldObjects(1), 100, -100, 100, 5)
'Call LoadCube(WorldObjects(2), -100, -100, 100, 5)
'Call LoadCube(WorldObjects(3), 100, 100, 100, 5)
Call LoadLandscape(WorldObjects(0), 0, 0, 100, LandWidth, LandLength, LandHeight, 20)

Do
If LastTime + TimeGap < GetTickCount Then 'Time for next Frame

If LastFCT + 1000 < GetTickCount Then 'Update FPS readout
FramesPS = (Frames - FramesL) / ((GetTickCount - LastFCT) / 1000)
LastFCT = GetTickCount
FramesL = Frames
Me.Caption = "Ni-Star Enterprises 3DWorld Mark III - We all live in a Fractal World -" + Str$(Round(FramesPS, 2)) + "fps -" + Str$(Frames) + " Frames"
End If

Call Graphics 'Do all the mathes and drawing
Call UserInput 'Allow user to move point of view

'Call MadRotate(WorldObjects(1)) 'Rotate everything madly
'Call MadRotate(WorldObjects(2)) 'Rotate everything madly
'Call MadRotate(WorldObjects(3)) 'Rotate everything madly

Frames = Frames + 1
LastTime = GetTickCount
End If
DoEvents
Loop While DemoGoing
End Sub

Private Sub LoadLandscape(object As Object3D, Xpos As Long, Ypos As Long, Zpos As Long, PointsWide As Long, PointsFar As Long, Heighty As Double, Scaler As Double)
Dim X As Long, Y As Long
Dim Indexia As Long
Dim Jedexia As Long

object.Verticies = (PointsWide + 1) * (PointsFar + 1)
ReDim object.Vertex(object.Verticies)
ReDim object.VertexScreen(object.Verticies)
ReDim object.VertexProcessed(object.Verticies)

Call RandomLandscape
Call SmoothLevel
Call NormalizeInvert
Call SmoothLevel
Call SmoothLevel
Call SmoothLevel
Call RemoveEdges

For X = 0 To PointsFar
For Y = 0 To PointsWide
object.Vertex(Indexia).X = Y - (PointsWide / 2)
object.Vertex(Indexia).Y = Heighty * PlasmaLandscape(Y, X) / 255
object.Vertex(Indexia).Z = X - (PointsFar / 2)
Indexia = Indexia + 1
Next
Next

'MsgBox Str$(Indexia)

object.Triangles = Indexia * 2
ReDim object.Triangle(object.Triangles)
ReDim object.TriVisible(object.Triangles)
ReDim object.ZBuffTri(object.Triangles)

Indexia = 0
For Y = 0 To PointsFar - 1
For X = 0 To PointsWide - 1
object.Triangle(Indexia).One = X + (Y * (PointsWide + 1))
object.Triangle(Indexia).Two = X + (PointsWide + 1) + (Y * (PointsWide + 1))
object.Triangle(Indexia).Three = X + 1 + (Y * (PointsWide + 1))
object.Triangle(Indexia).Colour = GetReliefColour((object.Vertex(object.Triangle(Indexia).One).Y + object.Vertex(object.Triangle(Indexia).Two).Y + object.Vertex(object.Triangle(Indexia).Three).Y) / (3 * Heighty))
Indexia = Indexia + 1
object.Triangle(object.Triangles / 2 + Indexia).One = X + (PointsWide + 2) + (Y * (PointsWide + 1))
object.Triangle(object.Triangles / 2 + Indexia).Two = X + (PointsWide + 1) + (Y * (PointsWide + 1))
object.Triangle(object.Triangles / 2 + Indexia).Three = X + 1 + (Y * (PointsWide + 1))
object.Triangle(object.Triangles / 2 + Indexia).Colour = GetReliefColour((object.Vertex(object.Triangle(object.Triangles / 2 + Indexia).One).Y + object.Vertex(object.Triangle(object.Triangles / 2 + Indexia).Two).Y + object.Vertex(object.Triangle(object.Triangles / 2 + Indexia).Three).Y) / (3 * Heighty))
Next X
Next Y

object.WorldCoords.X = Xpos
object.WorldCoords.Y = Ypos
object.WorldCoords.Z = Zpos

object.ScaleFactor = Scaler
End Sub

Private Sub Graphics()
Dim Xee As Long

For Xee = 0 To UBound(WorldObjects)
Call TranslateAndRotateObject(WorldObjects(Xee))
Call ZSortObject(WorldObjects(Xee))
Call ProjectToScreen(WorldObjects(Xee))
Call BackFaceCull(WorldObjects(Xee))
Call Lighting(WorldObjects(Xee))
Call RenderObject(WorldObjects(Xee))
Next Xee

For i = 0 To UBound(WorldObjects)
Xee = Xee + WorldObjects(i).Triangles
Next i

picBackBuffer.Print "3DWorld rendering" + Str$(Int(Xee * FramesPS)) + " polys per second"
picBackBuffer.Print "3DWorld with" + Str$(UBound(WorldObjects) + 1) + " objects," + Str$(Xee) + " polygons"
picBackBuffer.Print "3DWorld Z-Sorting per object"
If Blur Then picBackBuffer.Print "Motion blur On" Else picBackBuffer.Print "Motion blur Off"
If Outline Then picBackBuffer.Print "Outlined Filled Polys On" Else picBackBuffer.Print "Outlined Filled Polys Off"
If BackFaceRemove Then picBackBuffer.Print "BackFaceCulling On" Else picBackBuffer.Print "BackFaceCulling Off"
If RenderMode = 0 Then picBackBuffer.Print "Render Mode = WireFrame"
If RenderMode = 1 Then picBackBuffer.Print "Render Mode = Flat Filled Polys"
If RenderMode = 2 Then picBackBuffer.Print "Render Mode = Flat Filled Lit Polys"
If RenderMode = 3 Then picBackBuffer.Print "Render Mode = Dot Verticies"
If RenderMode = 4 Then picBackBuffer.Print "Render Mode = Outlined Polys"

If Blur Then AlphaBlending frm3DWorld.hdc, 0, 0, 512, 384, picBackBuffer.hdc, 0, 0, 512, 384, &HD00000 Else BitBlt frm3DWorld.hdc, 0, 0, 512, 384, picBackBuffer.hdc, 0, 0, vbSrcCopy
picBackBuffer.Cls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DemoGoing = False
End
End Sub

Private Sub UserInput()

If GetKeyState(vbKeyShift) And 4096 Then
If GetKeyState(vbKeyUp) And 4096 Then ViewPoint.RotationX = ViewPoint.RotationX - 0.1
If GetKeyState(vbKeyRight) And 4096 Then ViewPoint.RotationY = ViewPoint.RotationY + 0.1
If GetKeyState(vbKeyDown) And 4096 Then ViewPoint.RotationX = ViewPoint.RotationX + 0.1
If GetKeyState(vbKeyLeft) And 4096 Then ViewPoint.RotationY = ViewPoint.RotationY - 0.1
If GetKeyState(vbKeyPageUp) And 4096 Then ViewPoint.RotationZ = ViewPoint.RotationZ - 0.1
If GetKeyState(vbKeyPageDown) And 4096 Then ViewPoint.RotationZ = ViewPoint.RotationZ + 0.1
Else
If GetKeyState(vbKeyUp) And 4096 Then ViewPoint.Y = ViewPoint.Y - 1
If GetKeyState(vbKeyRight) And 4096 Then ViewPoint.X = ViewPoint.X + 1
If GetKeyState(vbKeyDown) And 4096 Then ViewPoint.Y = ViewPoint.Y + 1
If GetKeyState(vbKeyLeft) And 4096 Then ViewPoint.X = ViewPoint.X - 1
If GetKeyState(vbKeyPageUp) And 4096 Then ViewPoint.Z = ViewPoint.Z + 1
If GetKeyState(vbKeyPageDown) And 4096 Then ViewPoint.Z = ViewPoint.Z - 1
End If

If GetKeyState(vbKeyW) And 4096 Then WorldObjects(0).RotationX = WorldObjects(0).RotationX - 0.1 'WorldObjects.WorldCoords.Y = WorldObjects.WorldCoords.Y - 1
If GetKeyState(vbKeyD) And 4096 Then WorldObjects(0).RotationY = WorldObjects(0).RotationY + 0.1 'WorldObjects.WorldCoords.X = WorldObjects.WorldCoords.X - 1
If GetKeyState(vbKeyS) And 4096 Then WorldObjects(0).RotationX = WorldObjects(0).RotationX + 0.1 'WorldObjects.WorldCoords.Y = WorldObjects.WorldCoords.Y + 1
If GetKeyState(vbKeyA) And 4096 Then WorldObjects(0).RotationY = WorldObjects(0).RotationY - 0.1 'WorldObjects.WorldCoords.X = WorldObjects.WorldCoords.X + 1
If GetKeyState(vbKeyE) And 4096 Then WorldObjects(0).RotationZ = WorldObjects(0).RotationZ - 0.1
If GetKeyState(vbKeyQ) And 4096 Then WorldObjects(0).RotationZ = WorldObjects(0).RotationZ + 0.1


If GetKeyState(vbKeySubtract) And 4096 Then
WorldObjects(0).ScaleFactor = WorldObjects(0).ScaleFactor - 0.2
End If
If GetKeyState(vbKeyAdd) And 4096 Then
WorldObjects(0).ScaleFactor = WorldObjects(0).ScaleFactor + 0.2
End If

If GetKeyState(vbKeyB) And 4096 Then
RenderMode = RenderMode + 1
If RenderMode > 4 Then RenderMode = 0
Sleep 250
End If
If GetKeyState(vbKeyN) And 4096 Then
BackFaceRemove = BackFaceRemove + 1
If BackFaceRemove > 1 Then BackFaceRemove = 0
Sleep 250
End If

If GetKeyState(vbKeyM) And 4096 Then
If Blur Then Blur = False Else Blur = True
Sleep 250
End If

If GetKeyState(vbKeyO) And 4096 Then
If Outline Then Outline = False Else Outline = True
Sleep 250
End If

If GetKeyState(vbKeyR) And 4096 Then
Call LoadLandscape(WorldObjects(0), 0, 0, 100, LandWidth, LandLength, LandHeight, 20)
End If

End Sub

Private Sub TranslateAndRotateObject(object As Object3D)
Dim i As Integer, RotationBuffer As Point3D, s As Double

For i = 0 To object.Verticies
object.VertexProcessed(i).X = object.Vertex(i).X
object.VertexProcessed(i).Y = object.Vertex(i).Y
object.VertexProcessed(i).Z = object.Vertex(i).Z
Next

If object.RotationX > Pi2 Then object.RotationX = object.RotationX - Pi2
If object.RotationZ > Pi2 Then object.RotationY = object.RotationY - Pi2
If object.RotationY > Pi2 Then object.RotationZ = object.RotationZ - Pi2

'Rotate About Own Axis
For i = 0 To object.Verticies
s = object.VertexProcessed(i).Y * Cos(object.RotationX) - object.VertexProcessed(i).Z * Sin(object.RotationX)
object.VertexProcessed(i).Z = object.VertexProcessed(i).Z * Cos(object.RotationX) + object.VertexProcessed(i).Y * Sin(object.RotationX)
object.VertexProcessed(i).Y = s
s = object.VertexProcessed(i).Z * Cos(object.RotationY) - object.VertexProcessed(i).X * Sin(object.RotationY)
object.VertexProcessed(i).X = object.VertexProcessed(i).X * Cos(object.RotationY) + object.VertexProcessed(i).Z * Sin(object.RotationY)
object.VertexProcessed(i).Z = s
s = object.VertexProcessed(i).X * Cos(object.RotationZ) - object.VertexProcessed(i).Y * Sin(object.RotationZ)
object.VertexProcessed(i).Y = object.VertexProcessed(i).Y * Cos(object.RotationZ) + object.VertexProcessed(i).X * Sin(object.RotationZ)
object.VertexProcessed(i).X = s

object.VertexProcessed(i).X = object.VertexProcessed(i).X * object.ScaleFactor
object.VertexProcessed(i).Y = object.VertexProcessed(i).Y * object.ScaleFactor
object.VertexProcessed(i).Z = object.VertexProcessed(i).Z * object.ScaleFactor
Next


'Translate into 3D World
For i = 0 To object.Verticies
object.VertexProcessed(i).X = object.VertexProcessed(i).X + object.WorldCoords.X + ViewPoint.X
object.VertexProcessed(i).Y = object.VertexProcessed(i).Y + object.WorldCoords.Y + ViewPoint.Y
object.VertexProcessed(i).Z = object.VertexProcessed(i).Z + object.WorldCoords.Z + ViewPoint.Z
Next i

'Rotate About Camera
For i = 0 To object.Verticies
s = object.VertexProcessed(i).Y * Cos(ViewPoint.RotationX) - object.VertexProcessed(i).Z * Sin(ViewPoint.RotationX)
object.VertexProcessed(i).Z = object.VertexProcessed(i).Z * Cos(ViewPoint.RotationX) + object.VertexProcessed(i).Y * Sin(ViewPoint.RotationX)
object.VertexProcessed(i).Y = s
s = object.VertexProcessed(i).Z * Cos(ViewPoint.RotationY) - object.VertexProcessed(i).X * Sin(ViewPoint.RotationY)
object.VertexProcessed(i).X = object.VertexProcessed(i).X * Cos(ViewPoint.RotationY) + object.VertexProcessed(i).Z * Sin(ViewPoint.RotationY)
object.VertexProcessed(i).Z = s
s = object.VertexProcessed(i).X * Cos(ViewPoint.RotationZ) - object.VertexProcessed(i).Y * Sin(ViewPoint.RotationZ)
object.VertexProcessed(i).Y = object.VertexProcessed(i).Y * Cos(ViewPoint.RotationZ) + object.VertexProcessed(i).X * Sin(ViewPoint.RotationZ)
object.VertexProcessed(i).X = s
Next

End Sub

Private Sub ProjectToScreen(object As Object3D)
On Error Resume Next

Dim i As Integer

For i = 0 To object.Verticies
object.VertexScreen(i).X = HVPW + (object.VertexProcessed(i).X) / (0.005 * object.VertexProcessed(i).Z) '* Object.ScaleFactor))
object.VertexScreen(i).Y = HVPH + (object.VertexProcessed(i).Y) / (0.005 * object.VertexProcessed(i).Z) '* Object.ScaleFactor))
Next i
End Sub

Private Sub ZSortObject(object As Object3D)
'On Error Resume Next

Dim i As Integer, Index As Integer
Dim ThisValue As Long

'( face(0).y - face(2).y ) * ( face(1).x - face(0).x ) - ( face(0).x - face(2).x ) * ( face(1).y - face(0).y )

ReDim TriangleZ(object.Triangles) As Double  'Z of Tri
ReDim TriangleI(object.Triangles) As Long 'Index of tri
Dim a As Integer
For i = 0 To object.Triangles 'Get Triangle Distances
TriangleZ(i) = (object.VertexProcessed(object.Triangle(i).One).Z + object.VertexProcessed(object.Triangle(i).Two).Z + object.VertexProcessed(object.Triangle(i).Three).Z) / 3
Next
For Index = 0 To object.Triangles
For i = 0 To object.Triangles
If TriangleZ(i) > TriangleZ(a) Then a = i
Next i
TriangleZ(a) = 0
TriangleI(Index) = a
a = object.Triangles
Next Index
For i = 0 To object.Triangles
object.ZBuffTri(i) = TriangleI(i)
Next i
End Sub

Private Sub RenderObject(object As Object3D)
Dim i As Integer, TriAgain(2) As Point2D

If RenderMode = 0 Then 'WireFrame
For i = 0 To object.Triangles
If object.TriVisible(object.ZBuffTri(i)) = True Then
picBackBuffer.DrawStyle = 0
picBackBuffer.FillColor = object.Triangle(object.ZBuffTri(i)).Colour
picBackBuffer.FillStyle = 1
TriAgain(0).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).X
TriAgain(0).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).Y
TriAgain(1).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).X
TriAgain(1).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).Y
TriAgain(2).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).X
TriAgain(2).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).Y
Polygon picBackBuffer.hdc, TriAgain(0), 3
End If
Next i
ElseIf RenderMode = 1 Then 'Flat Shaded
For i = 0 To object.Triangles
If object.TriVisible(object.ZBuffTri(i)) = True Then
If Outline Then picBackBuffer.DrawStyle = 0 Else picBackBuffer.DrawStyle = 5
picBackBuffer.FillColor = object.Triangle(object.ZBuffTri(i)).Colour
picBackBuffer.FillStyle = 0
TriAgain(0).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).X
TriAgain(0).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).Y
TriAgain(1).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).X
TriAgain(1).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).Y
TriAgain(2).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).X
TriAgain(2).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).Y
Polygon picBackBuffer.hdc, TriAgain(0), 3
End If
Next i
ElseIf RenderMode = 2 Then 'Flat Shaded Lit
For i = 0 To object.Triangles
If object.TriVisible(object.ZBuffTri(i)) = True Then
If Outline Then picBackBuffer.DrawStyle = 0 Else picBackBuffer.DrawStyle = 5
picBackBuffer.FillColor = object.Triangle(object.ZBuffTri(i)).Lit
picBackBuffer.FillStyle = 0
TriAgain(0).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).X
TriAgain(0).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).Y
TriAgain(1).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).X
TriAgain(1).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).Y
TriAgain(2).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).X
TriAgain(2).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).Y
Polygon picBackBuffer.hdc, TriAgain(0), 3
End If
Next i
ElseIf RenderMode = 3 Then 'Draw Points
For i = 0 To object.Verticies
SetPixelV picBackBuffer.hdc, object.VertexScreen(i).X, object.VertexScreen(i).Y, &HC0C000
picBackBuffer.PSet (object.VertexScreen(i).X, object.VertexScreen(i).Y), &HC0C000
If i > 0 And i < 50 Then picBackBuffer.Print i
Next i
ElseIf RenderMode = 4 Then 'Flat non shaded
For i = 0 To object.Triangles
If object.TriVisible(object.ZBuffTri(i)) = True Then
picBackBuffer.DrawStyle = 0
picBackBuffer.FillColor = 0 'object.Triangle(object.ZBuffTri(i)).Colour
picBackBuffer.FillStyle = 0
TriAgain(0).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).X
TriAgain(0).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).One).Y
TriAgain(1).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).X
TriAgain(1).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Two).Y
TriAgain(2).X = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).X
TriAgain(2).Y = object.VertexScreen(object.Triangle(object.ZBuffTri(i)).Three).Y
Polygon picBackBuffer.hdc, TriAgain(0), 3
End If
Next i
End If
End Sub

Private Sub SetupWorld(ObjectPopulation As Long)
Randomize GetTickCount

ReDim WorldObjects(ObjectPopulation) As Object3D

ViewPoint.X = 0
ViewPoint.Y = 0
ViewPoint.Z = 100
ViewPoint.RotationX = 0
ViewPoint.RotationY = 0
ViewPoint.RotationZ = 0

Xadd = Rnd * (Pi / 32)
Yadd = Rnd * (Pi / 32)
Zadd = Rnd * (Pi / 32)
End Sub

Private Sub BackFaceCull(object As Object3D)
For i = 0 To object.Triangles
object.TriVisible(i) = True
If (object.VertexScreen(object.Triangle(i).One).X > VPW Or object.VertexScreen(object.Triangle(i).One).X < 0 Or object.VertexScreen(object.Triangle(i).One).Y > VPH Or object.VertexScreen(object.Triangle(i).One).Y < 0) And (object.VertexScreen(object.Triangle(i).Two).X > VPW Or object.VertexScreen(object.Triangle(i).Two).X < 0 Or object.VertexScreen(object.Triangle(i).Two).Y > VPH Or object.VertexScreen(object.Triangle(i).Two).Y < 0) And (object.VertexScreen(object.Triangle(i).Three).X > VPW Or object.VertexScreen(object.Triangle(i).Three).X < 0 Or object.VertexScreen(object.Triangle(i).Three).Y > VPH Or object.VertexScreen(object.Triangle(i).Three).Y < 0) Then object.TriVisible(i) = False
If (object.VertexProcessed(object.Triangle(i).One).Z < 1 Or object.VertexProcessed(object.Triangle(i).Two).Z < 1 Or object.VertexProcessed(object.Triangle(i).Three).Z < 1) Then object.TriVisible(i) = False
Next i
End Sub

Private Sub Lighting(object As Object3D)
Dim Red As Byte, Green As Byte, Blue As Byte
Dim R As Byte, G As Byte, B As Byte
Dim Indexia As Long
Dim Dist As Single

For Indexia = 0 To object.Triangles
Red = object.Triangle(Indexia).Colour Mod 256
Green = ((object.Triangle(Indexia).Colour And &HFF00) / 256&) Mod 256&
Blue = (object.Triangle(Indexia).Colour And &HFF0000) / 65536
Dist = ((object.VertexProcessed(object.Triangle(Indexia).One).Z + object.VertexProcessed(object.Triangle(Indexia).Two).Z + object.VertexProcessed(object.Triangle(Indexia).Three).Z) / 350)
If Dist < 1 Then Dist = 1
R = Red / Dist
G = Green / Dist
B = Blue / Dist
'If R > 255 Or R < 0 Or G > 255 Or G < 0 Or B > 255 Or B < 0 Then R = Red: G = Green: B = Blue
object.Triangle(Indexia).Lit = RGB(R, G, B)
Next Indexia

ShadeColor = RGB(R, G, B)
End Sub

Private Sub RandomLandscape()
Dim X As Long
Dim Y As Long
Dim Z As Long

Randomize Timer

For X = 0 To LandWidth + 2
For Y = 0 To LandLength + 2
Z = Rnd * 255
PlasmaLandscape(X, Y) = Z
Next Y
Next X
End Sub

Private Sub SmoothLevel()
Dim X As Long
Dim Y As Long
Dim Z As Long

For X = 1 To LandWidth + 1
For Y = 1 To LandLength + 1
Z = PlasmaLandscape(X - 1, Y - 1) + PlasmaLandscape(X - 1, Y) + PlasmaLandscape(X - 1, Y + 1) _
    + PlasmaLandscape(X, Y - 1) + PlasmaLandscape(X, Y) + PlasmaLandscape(X, Y + 1) + _
    PlasmaLandscape(X + 1, Y - 1) + PlasmaLandscape(X + 1, Y) + PlasmaLandscape(X + 1, Y + 1)
Z = Z \ 9
PlasmaLandscape(X, Y) = Z
Next Y
Next X
End Sub

Private Sub NormalizeInvert()
Dim Lowest As Long
Dim Highest As Long
Dim X As Long
Dim Y As Long
Dim Z As Long

Randomize Timer

Lowest = 255
Highest = 0

For X = 0 To LandWidth + 2
For Y = 0 To LandLength + 2
Z = PlasmaLandscape(X, Y)
If Z < Lowest Then Lowest = Z
If Z > Highest Then Highest = Z
Next Y
Next X

If Lowest = 0 Then Lowest = 1
If Highest = 0 Then Highest = 1

For X = 0 To LandWidth + 2
For Y = 0 To LandLength + 2
Z = PlasmaLandscape(X, Y)
If Z < 127 Then Z = -((Z - 127) * (255 / Lowest))
If Z > 127 Then Z = (Z - 127) * (255 / Highest)
If Z > 255 Then Z = 255
If Z < 255 Then Z = 0
PlasmaLandscape(X, Y) = Z
Next Y
Next X
End Sub

Private Sub RemoveEdges()
Dim X As Long
Dim Y As Long
Dim Z As Long

Randomize Timer

For X = 0 To LandWidth + 1
For Y = 0 To LandLength + 1
PlasmaLandscape(X, Y) = PlasmaLandscape(X + 1, Y + 1)
Next Y
Next X
End Sub

Private Function GetReliefColour(Relief As Double)
GetReliefColour = GetPixel(picLandColour.hdc, 0, Relief * 255)
If GetReliefColour = -1 Then MsgBox "What on earth?"
End Function
