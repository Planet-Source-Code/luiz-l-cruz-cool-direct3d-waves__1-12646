Attribute VB_Name = "Module1"
Public DX As New DirectX7 ' the main DirectX object

Public DDRAW As DirectDraw7 ' the DirectDraw object

Public SurfDesc As DDSURFACEDESC2 ' used for surface description
Public Primary As DirectDrawSurface7 ' the primary surface - represents the screen
Public Backbuffer As DirectDrawSurface7 ' the backbuffer surface - used for drawing on

Public Clipper As DirectDrawClipper ' the clip - to contain drawing in just the forms window

Public DestRect As RECT ' the source and destination rectangles
Public SrcRect As RECT

Public D3D As Direct3D7 ' the main Direct3D object
Public D3Ddevice As Direct3DDevice7 ' the rendering device

Public Viewport(0) As D3DRECT ' the viewport rectangle for the rendering device
Public VPdesc As D3DVIEWPORT7 ' the viewport description

Public Material As D3DMATERIAL7 ' the material for our polygon surface

Public Type tLight
    Light As D3DLIGHT7
    Tag As String
End Type

Public Light() As tLight

Function MakeLight(LightType As CONST_D3DLIGHTTYPE, Ambient As D3DCOLORVALUE, Diffuse As D3DCOLORVALUE, Specular As D3DCOLORVALUE, Pos As D3DVECTOR, Dir As D3DVECTOR, Optional attenuation0 As Single = 0, Optional attenuation1 As Single = 0, Optional attenuation2 As Single = 0, Optional FallOff As Single = 0, Optional phi As Single = 0, Optional theta As Single = 0, Optional Range As Single = 0) As D3DLIGHT7
With MakeLight
  .Ambient = Ambient
  .attenuation0 = attenuation0
  .attenuation1 = attenuation1
  .attenuation2 = attenuation2
  .Diffuse = Diffuse
  .direction = Dir
  .dltType = LightType
  .FallOff = FallOff
  .phi = phi
  .position = Pos
  .Range = Range
  .Specular = Specular
  .theta = theta
End With
End Function
Function MakeD3DCOLORVALUE(a As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
With MakeD3DCOLORVALUE
  .a = a
  .r = r
  .g = g
  .b = b
End With
End Function
Sub AddLight(NewLight As D3DLIGHT7, Tag As String, Enabled As Boolean)
On Error Resume Next
Dim UBL As Long
UBL = UBound(Light) + 1
ReDim Preserve Light(0 To UBL)
Light(UBL).Light = NewLight
Light(UBL).Tag = Tag

D3Ddevice.SetLight UBL, Light(UBL).Light
D3Ddevice.LightEnable UBL, Enabled
End Sub

Function MakeVector(x As Single, y As Single, z As Single) As D3DVECTOR
' copy x, y and z into the return value
With MakeVector
    .x = x
    .y = y
    .z = z
End With
End Function
