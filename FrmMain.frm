VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   1395
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clicked As Single
Const AngleP As Single = 0.5
Const iScale = 1
Const Xx = 5 * iScale
Const Yy = 5 * iScale
Const Zz = 50 * iScale
Dim Dif As Single
Dim currentAngle As Single
Dim currentAngle2 As Single
Dim matWorld As D3DMATRIX
Dim matView As D3DMATRIX
Dim matProj As D3DMATRIX
Dim ItemCount As Integer
Dim CCount As Integer
Dim Dx As DirectX8
Dim D3D As Direct3D8
Dim D3DDevice As Direct3DDevice8
Dim D3DX As D3DX8
Dim InMenu As Boolean
Dim Material As D3DMATERIAL8
Dim MatCol As D3DCOLORVALUE
Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
Const Lit_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
Const pi As Single = 3.14159265358979
Const DegTrans As Single = pi / 180
Private Type LITVERTEX
    X As Single
    Y As Single
    Z As Single
    color As Long
    specular As Long
    Tu As Single
    TV As Single
End Type
Dim Points() As LITVERTEX
Private Type iCube
    FrontP(5) As LITVERTEX
End Type
Dim Cubes() As iCube
Private Function SetDispMode()
    DispMode.Format = D3DFMT_X8R8G8B8
    DispMode.Width = 800
    DispMode.Height = 600
    D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
    D3DWindow.Windowed = 0
    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferFormat = D3DFMT_X8R8G8B8
    D3DWindow.BackBufferWidth = 800
    D3DWindow.BackBufferHeight = 600
    D3DWindow.hDeviceWindow = FrmMain.hWnd
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
End Function
Public Function CreateDevice()
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
    SetDispMode
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FrmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    D3DDevice.SetVertexShader Lit_FVF
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
End Function
Public Function SetTheMatrix(From As D3DVECTOR, Too As D3DVECTOR)
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    D3DXMatrixLookAtLH matView, From, Too, MakeVektor(0, 1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    D3DXMatrixPerspectiveFovLH matProj, pi / 4, 1, 0.1, 500
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Function
Public Function Initialise() As Boolean
    CreateDevice
    Call SetTheMatrix(MakeVektor(0, 10, 100), MakeVektor(0, 0, Zz))
    If InitialiseGeometry() = True Then
        Initialise = True
        Exit Function
    End If
    Initialise = False
End Function
Private Function InitialiseGeometry() As Boolean
    CreateMaterial
    Call CreateAllCubes(CCount)
    ReDim Points((ItemCount + 1) * 6) As LITVERTEX
    For k% = 0 To ItemCount
        With Cubes(k%)
            For t% = 0 To 5
                Points(k% * 6 + t%) = .FrontP(t%)
            Next t%
        End With
    Next k%
    InitialiseGeometry = True
End Function
Public Function MakeVektor(X As Single, Y As Single, Z As Single) As D3DVECTOR
    With MakeVektor
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function
Public Function CreateMaterial()
        With MatCol
            .a = 1
            .r = 1
            .g = 1
            .b = 1
        End With
        Material.Ambient = MatCol
        Material.diffuse = MatCol
        D3DDevice.SetMaterial Material
End Function
Public Function Render()
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0

D3DDevice.BeginScene
    
    D3DDevice.SetRenderState D3DRS_AMBIENT, RGB(100, 10, 10)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, (ItemCount + 1) * 2, Points(0), Len(Points(0))

D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Private Function MakeUVertex(X As Single, Y As Single, Z As Single) As LITVERTEX
    With MakeUVertex
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function
Public Function CreateAllCubes(Count As Integer)
    ItemCount = Count - 1
    ReDim Cubes(ItemCount) As iCube
    For k% = 0 To ItemCount
        With Cubes(k%)
            .FrontP(0) = MakeUVertex(Xx, -Yy, Zz)
            .FrontP(1) = MakeUVertex(-Xx, -Yy, Zz)
            .FrontP(2) = MakeUVertex(Xx, Yy, Zz)
            .FrontP(3) = MakeUVertex(Xx, Yy, Zz)
            .FrontP(4) = MakeUVertex(-Xx, Yy, Zz)
            .FrontP(5) = MakeUVertex(-Xx, -Yy, Zz)
        End With
    Next k%
    Dim ThisOne As Single
    Dim xxx As Single, zzz As Single
    Dim CosA As Single, SinA As Single
    Dif = 360 / (ItemCount + 1)
    For k% = 0 To ItemCount
        CosA = Cos(ThisOne * DegTrans)
        SinA = Sin(ThisOne * DegTrans)
        With Cubes(k%)
            For t% = 0 To 5
                xxx = .FrontP(t%).X
                zzz = .FrontP(t%).Z
                .FrontP(t%).X = zzz * SinA + xxx * CosA
                .FrontP(t%).Z = zzz * CosA - xxx * SinA
            Next t%
        End With
        ThisOne = ThisOne + Dif
    Next k%
End Function
Public Function RotatePoints(Deg As Single)
Dim xxx As Single, zzz As Single
Dim CosA As Single, SinA As Single
CosA = Cos(Deg * DegTrans)
SinA = Sin(Deg * DegTrans)
    For k% = 0 To (ItemCount + 1) * 6 - 1
        xxx = Points(k%).X
        zzz = Points(k%).Z
        Points(k%).X = zzz * SinA + xxx * CosA
        Points(k%).Z = zzz * CosA - xxx * SinA
    Next k%
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            InMenu = False
            DoEvents
            Set D3DDevice = Nothing
            Set D3D = Nothing
            End
        Case vbKeyLeft
            If clicked = 0 Then
                clicked = 2
            End If
        Case vbKeyRight
            If clicked = 0 Then
                clicked = 1
            End If
        Case vbKeyUp
            If clicked = 0 Then
                InMenu = False
            End If
    End Select
End Sub
Private Sub Form_Load()
Me.Show
DoEvents
CCount = 10 ' how many objects there are...
clicked = 0
InMenu = Initialise()
currentAngle2 = Dif
RotatePoints 0.001
Do While InMenu
    If clicked = 1 Then
        If currentAngle <= Dif Then
            RotatePoints AngleP
            currentAngle = currentAngle + AngleP
        Else
            clicked = 0
            currentAngle = currentAngle - Dif
        End If
    ElseIf clicked = 2 Then
        If currentAngle2 >= 0 Then
            RotatePoints -AngleP
            currentAngle2 = currentAngle2 - AngleP
        Else
            clicked = 0
            currentAngle2 = currentAngle2 + Dif
        End If
    End If
    Render
    DoEvents
Loop
End Sub

