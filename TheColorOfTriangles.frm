VERSION 5.00
Begin VB.Form frmTheColorOfTriangles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Color Of Triangles"
   ClientHeight    =   4695
   ClientLeft      =   135
   ClientTop       =   990
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4575
   Begin VB.CheckBox chkUseDifferentColors 
      Caption         =   "Use Different Colors"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmTheColorOfTriangles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI As Single = 3.14159265

Private m_dx As DirectX7
Private m_dd As DirectDraw7
Private m_ddPrimarySurface As DirectDrawSurface7
Private m_ddRenderSurface As DirectDrawSurface7
Private m_d3d As Direct3D7
Private m_d3dDevice As Direct3DDevice7
Private m_CameraR As Double
Private m_CameraTheta As Double
Private m_CameraDTheta As Double
Private m_CameraY As Double
' This is an array because we need to pass an array to
' m_d3dDevice.Clear.
Private m_ViewportRect(0) As D3DRECT

' True while the program should redraw the triangle.
Private m_Running As Boolean

' The vertices we will draw.
Private m_NumVertices As Integer
Private m_Vertex() As D3DVERTEX

' Picture dimensions.
Private m_PictureRect As RECT

' Rendering surface dimensions.
Private m_RenderRect As RECT
' Add a triangle to the list.
Private Sub MakeTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal z3 As Single)
    ' Add room for the triangle's three vertices.
    m_NumVertices = m_NumVertices + 3
    ReDim Preserve m_Vertex(1 To m_NumVertices)

    ' Make the vertices.
    With m_Vertex(m_NumVertices - 2)
        .x = x1
        .y = y1
        .z = z1
    End With
    With m_Vertex(m_NumVertices - 1)
        .x = x2
        .y = y2
        .z = z2
    End With
    With m_Vertex(m_NumVertices)
        .x = x3
        .y = y3
        .z = z3
    End With
End Sub

' Display the scene.
Private Sub RenderLoop()
Dim status As Long

    m_Running = True
    Do While m_Running
        ' Draw the objects.
        RenderObjects

        ' Display the results.
        status = m_ddPrimarySurface.Blt(m_PictureRect, m_ddRenderSurface, m_RenderRect, DDBLT_WAIT)
        If status <> DD_OK Then
            MsgBox "Error " & Format$(status) & " displaying the scene."
            m_Running = False
        End If
        DoEvents
    Loop
End Sub

Private Sub Form_Load()
    ' Initialize DirectDraw.
    InitializeDirectDraw

    ' Initialize Direct3D.
    InitializeDirect3D

    ' Initialize the scene.
    InitializeScene

    ' Initialize the objects we will display.
    InitializeObjects

    Show

    ' Display the scene rotating.
    RenderLoop

    ' End.
    Unload Me
End Sub
' Draw the objects.
Private Sub RenderObjects()
Dim clr As Single
Dim dclr As Single
Dim i As Integer
Dim num_triangles As Integer
Dim Surface As Integer
Dim camera_x As Double
Dim camera_z As Double
Dim matrix_Camera As D3DMATRIX
Static Shape
Shape = "Cube"
    ' Clear the viewport.
    m_d3dDevice.Clear 1, m_ViewportRect(), D3DCLEAR_TARGET, _
        m_dx.CreateColorRGB(0#, 0#, 0.5), 1, 0

    ' Begin the scene.
    m_d3dDevice.BeginScene

    ' Set the color for the first triangle.
    clr = 0.9
' set viewing position
camera_x = m_CameraR * Cos(m_CameraTheta)
camera_z = m_CameraR * Sin(m_CameraTheta)
m_dx.ViewMatrix matrix_Camera, _
    MakeVector(camera_x, m_CameraY, camera_z), _
    MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
    m_d3dDevice.SetTransform D3DTRANSFORMSTATE_VIEW, matrix_Camera
    'revolve viewing position
    m_CameraTheta = m_CameraTheta + 0.01
    ' See if we should use different colors for the different
    ' triangles.
    If chkUseDifferentColors.Value = vbChecked Then
        ' Use different colors. Set the color increment.
        dclr = -0.05
    Else
        ' Use the same colors. Set the color increment to zero.
        dclr = 0
    End If

    ' Draw the triangles.
        num_triangles = m_NumVertices \ 3
    For i = 1 To num_triangles
        ' Set the ambient color to the next shade of orange.
        m_d3dDevice.SetRenderState D3DRENDERSTATE_AMBIENT, _
            m_dx.CreateColorRGB(clr, clr / 2, 0#)
        ' Increment the color for the next triangle.
        clr = clr + dclr
        ' Draw the triangle.
        m_d3dDevice.DrawPrimitive D3DPT_TRIANGLELIST, _
            D3DFVF_VERTEX, m_Vertex((i - 1) * 3 + 1), 3, D3DDP_DEFAULT
    Next i
    ' End the scene.
    On Error Resume Next
    m_d3dDevice.EndScene
End Sub
Private Sub Form_Unload(Cancel As Integer)
    m_Running = False
End Sub


' Initalize DirectDraw.
Private Sub InitializeDirectDraw()
Dim surf_desc As DDSURFACEDESC2

    ' Create the DirectDraw object and set cooperative level.
    Set m_dx = New DirectX7
    Set m_dd = m_dx.DirectDrawCreate("")
    m_dd.SetCooperativeLevel Picture1.hWnd, DDSCL_NORMAL

    ' Create the primary drawing surface.
    surf_desc.lFlags = DDSD_CAPS
    surf_desc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set m_ddPrimarySurface = m_dd.CreateSurface(surf_desc)

    ' Save the picture's size for later use.
    m_dx.GetWindowRect Picture1.hWnd, m_PictureRect

    ' Create the render surface making it fit Picture1.
    ' Specify system memory because we may use the RGB rasterizer.
    surf_desc.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    surf_desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_SYSTEMMEMORY
    surf_desc.lWidth = m_PictureRect.Right - m_PictureRect.Left
    surf_desc.lHeight = m_PictureRect.Bottom - m_PictureRect.Top
    Set m_ddRenderSurface = m_dd.CreateSurface(surf_desc)

    ' Save the size of the render surface for later use.
    With m_RenderRect
        .Left = 0
        .Top = 0
        .Bottom = surf_desc.lHeight
        .Right = surf_desc.lWidth
    End With

    ' Save a reference to the Direct3D object.
    Set m_d3d = m_dd.GetDirect3D
End Sub
' Initalize Direct3D.
Private Sub InitializeDirect3D()
Dim surf_desc As DDSURFACEDESC2
Dim viewport_desc As D3DVIEWPORT7

    ' Ensure that the display mode uses greater than 8-bit color.
    m_dd.GetDisplayMode surf_desc

    If surf_desc.ddpfPixelFormat.lRGBBitCount <= 8 Then
        MsgBox "This program requires a color mode higher than 8-bit."
        End
    End If

    ' Create the Direct3D device. Try for IID_IDirect3DHALDevice
    ' first and IID_IDirect3DRGBDevice if it isn't available.
    On Error Resume Next
    Set m_d3dDevice = m_d3d.CreateDevice("IID_IDirect3DHALDevice", m_ddRenderSurface)
    If m_d3dDevice Is Nothing Then
        Set m_d3dDevice = m_d3d.CreateDevice("IID_IDirect3DRGBDevice", m_ddRenderSurface)
    End If
    If m_d3dDevice Is Nothing Then
        ' We failed to create a device.
        MsgBox "Could not create a Direct3D device."
        End
    End If

    ' Define the viewport rectangle.
    With viewport_desc
        .lWidth = m_PictureRect.Right - m_PictureRect.Left
        .lHeight = m_PictureRect.Bottom - m_PictureRect.Top
        .minz = 0#
        .maxz = 1#
    End With
    m_d3dDevice.SetViewport viewport_desc

    ' Save the viewport rectangle for later use.
    With m_ViewportRect(0)
        .x1 = 0
        .y1 = 0
        .x2 = viewport_desc.lWidth
        .y2 = viewport_desc.lHeight
    End With
End Sub
' Initalize the scene (lighting, material, etc).
Private Sub InitializeScene()
Dim matrix_projection As D3DMATRIX
Dim matrix_Camera As D3DMATRIX
Dim material As D3DMATERIAL7
m_CameraR = Sqr(4 * 4 + 20 * 20)
m_CameraTheta = 0
m_CameraY = 3
    ' Set the device's material so it reflects all light.
    With material
        .Ambient.r = 1#
        .Ambient.g = 1#
        .Ambient.B = 1#
    End With
    m_d3dDevice.SetMaterial material

    ' Define the projection's clipping planes.
    m_dx.ProjectionMatrix matrix_projection, 1, 10000, PI / 2
    m_d3dDevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matrix_projection

    ' Set the viewing position to (4, 3, -20).
'    m_dx.ViewMatrix matrix_camera, MakeVector(4, 3, -20), _
'        MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
'    m_d3dDevice.SetTransform D3DTRANSFORMSTATE_VIEW, matrix_camera
End Sub
' Initalize the objects we will display.
Private Sub InitializeObjects()
'    MakeTriangle 0, 10, 0, 0, 0, 10, 10, 0, 0
'    MakeTriangle 0, 10, 0, 10, 0, 0, 0, 0, -10
'    MakeTriangle 0, 10, 0, 0, 0, -10, -10, 0, 0
'    MakeTriangle 0, 10, 0, -10, 0, 0, 0, 0, 10
'    MakeTriangle 0, -10, 0, 0, 0, -10, 10, 0, 0
'    MakeTriangle 0, -10, 0, 10, 0, 0, 0, 0, 10
'    MakeTriangle 0, -10, 0, 0, 0, 10, -10, 0, 0
'    MakeTriangle 0, -10, 0, -10, 0, 0, 0, 0, -10
MakeTriangle -10, -10, 0, -10, 10, 0, 10, -10, 0
MakeTriangle -10, 10, 0, 10, 10, 0, 10, -10, 0
MakeTriangle -10, -10, 20, -10, 10, 0, -10, -10, 0
MakeTriangle -10, -10, 20, -10, 10, 20, -10, 10, 0
MakeTriangle 10, 10, 20, -10, 10, 20, -10, -10, 20
MakeTriangle 10, -10, 20, 10, 10, 20, -10, -10, 20
MakeTriangle 10, 10, 0, 10, 10, 20, 10, -10, 20
MakeTriangle 10, -10, 0, 10, 10, 0, 10, -10, 20
MakeTriangle -10, 10, 0, -10, 10, 20, 10, 10, 20
MakeTriangle -10, 10, 0, 10, 10, 20, 10, 10, 0
MakeTriangle 10, -10, 20, -10, -10, 20, -10, -10, 0
MakeTriangle -10, -10, 0, 10, -10, 0, 10, -10, 20
End Sub
' Make a vector with the given components.
Private Function MakeVector(a As Double, B As Double, C As Double) As D3DVECTOR
Dim result As D3DVECTOR

    With result
        .x = a
        .y = B
        .z = C
    End With

    MakeVector = result
End Function

