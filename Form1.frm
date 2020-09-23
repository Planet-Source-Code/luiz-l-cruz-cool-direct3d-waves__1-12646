VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]
'
'                          By LUIZ LAYDNER
'
' I made this program based on the Direct 3D tutorial by
' Simon Price, available at www.planet-soure-code.com
' Thanks Simon for that Tutorial!!!!
'
'[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]




' matrices we will need to describe transformations
Dim matWorld As D3DMATRIX ' world matrix
Dim matView  As D3DMATRIX ' View matrix
Dim matProj As D3DMATRIX 'projection matrix

Dim matSpin As D3DMATRIX ' this will be used to rotate stuff
Dim matSpin2 As D3DMATRIX
Dim matSpin3 As D3DMATRIX

Dim Vertices() As D3DVERTEX

Dim cont As Double ' used to count vertices
Dim Counter As Long ' used to rotate the world matrix
Dim wave As Integer ' used to make a wave

Dim EndNow As Boolean ' this tells the program when to end

Const DETAIL = 0.3 ' if you decrease  this, there will be more triangles
Const AMPLITUDE = 0.5 'wave amplitude
Const FREQUENCY = 2 'frequency
Const SIZE = 1


' DirectDrawInit - Initializes DirectDraw objects
Function DirectDrawInit() As Long

Set DDRAW = DX.DirectDrawCreate("") ' create the directdraw object

DDRAW.SetCooperativeLevel hWnd, DDSCL_NORMAL ' set the cooperative level, we only need normal

' set the properties of the primary surface
SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE

Set Primary = DDRAW.CreateSurface(SurfDesc) ' create the primary surface

' set up the backbuffer surface (which will be where we render the 3D view)
SurfDesc.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE

' use the size of the form to determine the size of the render target
' and viewport rectangle
DX.GetWindowRect hWnd, DestRect ' set the dimensions of the surface description

SurfDesc.lWidth = DestRect.Right - DestRect.Left
SurfDesc.lHeight = DestRect.Bottom - DestRect.Top

Set Backbuffer = DDRAW.CreateSurface(SurfDesc) ' create the backbuffer surface

' cache the size of the render target for later use
With SrcRect
        .Left = 0: .Top = 0
        .Bottom = SurfDesc.lHeight
        .Right = SurfDesc.lWidth
End With

' create a DirectDrawClipper and attach it to the primary surface.
Set Clipper = DDRAW.CreateClipper(0)
Clipper.SetHWnd hWnd
Primary.SetClipper Clipper


End Function

'Direct3DInit - Initializes Direct3D objects

Function Direct3DInit() As Long


Set D3D = DDRAW.GetDirect3D
' create the direct3d object

' create the rendering device - we are using software emulation only
Set D3Ddevice = D3D.CreateDevice("IID_IDirect3DRGBDevice", Backbuffer)

' set the viewport rectangle.
VPdesc.lWidth = DestRect.Right - DestRect.Left
VPdesc.lHeight = DestRect.Bottom - DestRect.Top
VPdesc.minz = 0
VPdesc.maxz = 1
D3Ddevice.SetViewport VPdesc

' cache the viewport rectangle for later use
With Viewport(0)
    .X1 = 0: .Y1 = 0
    .X2 = VPdesc.lWidth
    .Y2 = VPdesc.lHeight
End With
    


' disable culling
D3Ddevice.SetRenderState D3DRENDERSTATE_CULLMODE, D3DCULL_NONE
'wireframe mode - try using D3DFILL_SOLID too
D3Ddevice.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_WIREFRAME
'flat shading
D3Ddevice.SetRenderState D3DRENDERSTATE_SHADEMODE, D3DSHADE_FLAT

' set the material
Material.Ambient.r = 1: Material.Ambient.g = 1: Material.Ambient.b = 1
D3Ddevice.SetMaterial Material

' the world matrix - all polygons in world space are transformed by this matrix
DX.IdentityMatrix matWorld ' used to initiate the matrix
D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matWorld

' the view matrix - basically the camera position is at -5
' (although it's really just making the whole world at +5)
DX.IdentityMatrix matView ' used to initiate the matrix
DX.ViewMatrix matView, MakeVector(0, 0, -5), MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
D3Ddevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView

' the projection matrix - decides how the 3D scene is projected onto the 2D surface
DX.IdentityMatrix matProj ' used to initiate the matrix
DX.ProjectionMatrix matProj, 1, 1000, 3.14 / 2
D3Ddevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj

'lights
Dim Cor As D3DCOLORVALUE
Dim MainLight As D3DLIGHT7

Cor = MakeD3DCOLORVALUE(10, 10, 200, 10)
MainLight = MakeLight(D3DLIGHT_POINT, Cor, Cor, Cor, MakeVector(10, 0, 5), MakeVector(0, 0, 0), 0, 1, 0, 0, 0, 0, 1000)
AddLight MainLight, "Sun", True


End Function

' RenderIt - renders the scene
Sub RenderIt()


 CreateWaves SIZE, Counter, FREQUENCY, AMPLITUDE
 
Do While EndNow = False
    ' increase the counter
    Counter = Counter + 1
     CreateWaves SIZE, Counter, FREQUENCY, AMPLITUDE
   
    
    ' clear the viewport
    D3Ddevice.Clear 1, Viewport(), D3DCLEAR_TARGET, vbBlack, 0, 0
    
    D3Ddevice.BeginScene
    D3Ddevice.DrawPrimitive D3DPT_TRIANGLELIST, D3DFVF_VERTEX, Vertices(0), cont, D3DDP_DEFAULT
    D3Ddevice.EndScene

    ' rotate the matrix
    DX.RotateYMatrix matSpin, Counter / 60
    DX.RotateZMatrix matSpin2, Counter / 60
    DX.RotateXMatrix matSpin3, Counter / 60
           
        
    DX.MatrixMultiply matSpin2, matSpin2, matSpin
    DX.MatrixMultiply matSpin3, matSpin3, matSpin2

    

    ' set the new world transform matrix
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matSpin2
    
    ' copy the backbuffer to the screen
    DX.GetWindowRect hWnd, DestRect
    Primary.Blt DestRect, Backbuffer, SrcRect, DDBLT_WAIT
    
    ' look for window messages - we need to know when the escape key is pressed
    DoEvents

Loop
End Sub


Private Sub Form_Load()
Form1.Width = Screen.Width
Form1.Height = Screen.Height
Form1.Left = 0
Form1.Top = 0
' show the form
Show
Call DirectDrawInit
Call Direct3DInit

' call rendering loop
RenderIt

MsgBox "3D Waves, made by Luiz Laydner", vbOKOnly
Unload Me
End Sub

'The Form_Unload event - Stops the main loop

Private Sub Form_Unload(Cancel As Integer)
EndNow = True
End Sub

' The Form_KeyDown event - Also stops the main loop

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' end program if escape is pressed
If KeyCode = vbKeyEscape Then EndNow = True
End Sub

Private Sub CreateWaves(Raio As Single, var As Long, freq, amp)
Dim x As Double, y As Double, z As Double

Dim temp(0 To 3) As Double

cont = 0

NumberOfVert = 6 * ((9 * Raio) / DETAIL) ^ 2 'number of vertices needed

Erase Vertices ' clear array
ReDim Vertices(0 To NumberOfVert)


For x = 3 * -Raio To 3 * Raio Step DETAIL
    For z = 3 * -Raio To 3 * Raio Step DETAIL

    temp(0) = (x ^ 2) + (z ^ 2)
    temp(1) = ((x + DETAIL) ^ 2 + (z) ^ 2)
    temp(2) = ((x) ^ 2 + (z + DETAIL) ^ 2)
    temp(3) = ((x + DETAIL) ^ 2 + (z + DETAIL) ^ 2)
    
    If temp(0) > 0 And temp(1) > 0 And temp(2) > 0 And temp(3) > 0 Then
        
        y = Sqr(temp(0))
        
        Vertices(cont).x = x
        Vertices(cont).z = z
        Vertices(cont).y = amp * Sin(freq * ((y) - (var) / 10))
        cont = cont + 1
        
        y = Sqr(temp(1))
        Vertices(cont).x = x + DETAIL
        Vertices(cont).z = z
        Vertices(cont).y = amp * Sin(freq * ((y) - (var) / 10))
        cont = cont + 1
        
        y = Sqr(temp(2))
        Vertices(cont).x = x
        Vertices(cont).z = z + DETAIL
        Vertices(cont).y = amp * Sin(freq * ((y) - (var) / 10))
        cont = cont + 1
        
        y = Sqr(temp(1))
        Vertices(cont).x = x + DETAIL
        Vertices(cont).z = z
        Vertices(cont).y = amp * Sin(freq * ((y) - (var) / 10))
        cont = cont + 1
        
        y = Sqr(temp(2))
        Vertices(cont).x = x
        Vertices(cont).z = z + DETAIL
        Vertices(cont).y = amp * Sin(freq * ((y) - (var) / 10))
        cont = cont + 1
        
        y = Sqr(temp(3))
        Vertices(cont).x = x + DETAIL
        Vertices(cont).z = z + DETAIL
        Vertices(cont).y = amp * Sin(freq * ((y) - (var) / 10))
        cont = cont + 1
    End If


    Next
Next

End Sub
