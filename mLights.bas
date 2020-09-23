Attribute VB_Name = "mLights"
Option Explicit

'#Requires
'# Direct3DVertexBuffer8
'# That initD3D has been run


'# Constant for untransformed and unlit verticies
Public Const NORMAL_FVF = (D3DFVF_XYZ Or D3DFVF_NORMAL)

'# Type for untransformed and unlit verticies.
'# Requires a normal to be calculated for each
'# vertex if lighting is to work
Type NORMALVERTEX
    X As Single
    Y As Single
    Z As Single
    nX As Single
    nY As Single
    nZ As Single
End Type

'#initLightingGeo

'# Creates an octahedron as in mGeometry, except
'# that normals are now calculated for each triangle
'# and that the verticies are now in clockwise order:
'# 0-------1
'#        /
'#       /
'#      /
'#     /
'#    /
'#   /
'#  /
'# 2

'#GenerateTriangleNormals

'# Calculates the normals for each triangle by creating
'# a vector from 0-1 and from 0-2.  It then works out the
'# cross product of these vectors and normalises it.
'# The returned value is then applied to each vertex
'# in the triangle.

'#SetupLights

'# Sets the material of the scene and
'# creates a coloured point light.


Sub initLightingGeo()
Dim normVec As D3DVECTOR
Dim rt2 As Single
Dim verts(23) As NORMALVERTEX
    
    rt2 = Sqr(2)
    
    verts(0) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(1) = createNormVert(1, 0, -1, 0, 0, 0)
    verts(2) = createNormVert(-1, 0, -1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(0), verts(1), verts(2))
        verts(0).nX = normVec.X: verts(0).nY = normVec.Y: verts(0).nZ = normVec.Z
        verts(1).nX = normVec.X: verts(1).nY = normVec.Y: verts(1).nZ = normVec.Z
        verts(2).nX = normVec.X: verts(2).nY = normVec.Y: verts(2).nZ = normVec.Z
        
        
    verts(3) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(4) = createNormVert(-1, 0, 1, 0, 0, 0)
    verts(5) = createNormVert(1, 0, 1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(3), verts(4), verts(5))
        verts(3).nX = normVec.X: verts(3).nY = normVec.Y: verts(3).nZ = normVec.Z
        verts(4).nX = normVec.X: verts(4).nY = normVec.Y: verts(4).nZ = normVec.Z
        verts(5).nX = normVec.X: verts(5).nY = normVec.Y: verts(5).nZ = normVec.Z
    

    verts(6) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(7) = createNormVert(1, 0, 1, 0, 0, 0)
    verts(8) = createNormVert(1, 0, -1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(6), verts(7), verts(8))
        verts(6).nX = normVec.X: verts(6).nY = normVec.Y: verts(6).nZ = normVec.Z
        verts(7).nX = normVec.X: verts(7).nY = normVec.Y: verts(7).nZ = normVec.Z
        verts(8).nX = normVec.X: verts(8).nY = normVec.Y: verts(8).nZ = normVec.Z
    

    verts(9) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(10) = createNormVert(-1, 0, -1, 0, 0, 0)
    verts(11) = createNormVert(-1, 0, 1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(9), verts(10), verts(11))
        verts(9).nX = normVec.X: verts(9).nY = normVec.Y: verts(9).nZ = normVec.Z
        verts(10).nX = normVec.X: verts(10).nY = normVec.Y: verts(10).nZ = normVec.Z
        verts(11).nX = normVec.X: verts(11).nY = normVec.Y: verts(11).nZ = normVec.Z
    
    
    verts(12) = createNormVert(-1, 0, -1, 0, 0, 0)
    verts(13) = createNormVert(1, 0, -1, 0, 0, 0)
    verts(14) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(12), verts(13), verts(14))
        verts(12).nX = normVec.X: verts(12).nY = normVec.Y: verts(12).nZ = normVec.Z
        verts(13).nX = normVec.X: verts(13).nY = normVec.Y: verts(13).nZ = normVec.Z
        verts(14).nX = normVec.X: verts(14).nY = normVec.Y: verts(14).nZ = normVec.Z
    
    
    verts(15) = createNormVert(1, 0, 1, 0, 0, 0)
    verts(16) = createNormVert(-1, 0, 1, 0, 0, 0)
    verts(17) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(15), verts(16), verts(17))
        verts(15).nX = normVec.X: verts(15).nY = normVec.Y: verts(15).nZ = normVec.Z
        verts(16).nX = normVec.X: verts(16).nY = normVec.Y: verts(16).nZ = normVec.Z
        verts(17).nX = normVec.X: verts(17).nY = normVec.Y: verts(17).nZ = normVec.Z
    
    
    verts(18) = createNormVert(1, 0, -1, 0, 0, 0)
    verts(19) = createNormVert(1, 0, 1, 0, 0, 0)
    verts(20) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(18), verts(19), verts(20))
        verts(18).nX = normVec.X: verts(18).nY = normVec.Y: verts(18).nZ = normVec.Z
        verts(19).nX = normVec.X: verts(19).nY = normVec.Y: verts(19).nZ = normVec.Z
        verts(20).nX = normVec.X: verts(20).nY = normVec.Y: verts(20).nZ = normVec.Z
        
        
    verts(21) = createNormVert(-1, 0, 1, 0, 0, 0)
    verts(22) = createNormVert(-1, 0, -1, 0, 0, 0)
    verts(23) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(21), verts(22), verts(23))
        verts(21).nX = normVec.X: verts(21).nY = normVec.Y: verts(21).nZ = normVec.Z
        verts(22).nX = normVec.X: verts(22).nY = normVec.Y: verts(22).nZ = normVec.Z
        verts(23).nX = normVec.X: verts(23).nY = normVec.Y: verts(23).nZ = normVec.Z
    
        
    D3DDevice.SetVertexShader NORMAL_FVF
    
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    D3DDevice.SetRenderState D3DRS_AMBIENT, &H101010
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 24, _
                                                    0, _
                                                    NORMAL_FVF, _
                                                    D3DPOOL_DEFAULT)

    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 24, 0, verts(0)


End Sub

Private Function GenerateTriangleNormals(p0 As NORMALVERTEX, p1 As NORMALVERTEX, p2 As NORMALVERTEX) As D3DVECTOR

Dim vNorm As D3DVECTOR
Dim temp1 As D3DVECTOR
Dim temp2 As D3DVECTOR
    
    temp1.X = p1.X - p0.X
    temp1.Y = p1.Y - p0.Y
    temp1.Z = p1.Z - p0.Z
            
    temp2.X = p2.X - p0.X
    temp2.Y = p2.Y - p0.Y
    temp2.Z = p2.Z - p0.Z
       
    D3DXVec3Cross vNorm, temp1, temp2

    D3DXVec3Normalize vNorm, vNorm

    GenerateTriangleNormals.X = vNorm.X
    GenerateTriangleNormals.Y = vNorm.Y
    GenerateTriangleNormals.Z = vNorm.Z

End Function

Sub SetupLights()


Dim Material As D3DMATERIAL8
Dim Colour As D3DCOLORVALUE

    With Colour
        .a = 1
        .r = 1
        .g = 1
        .b = 1
    End With

    Material.Ambient = Colour
    Material.diffuse = Colour

    D3DDevice.SetMaterial Material

    With Light
        .Type = D3DLIGHT_POINT
        
        .Position = D3DVec(0, 3, -5)
        
        .diffuse.r = Rnd() * 2
        .diffuse.g = Rnd() * 2
        .diffuse.b = Rnd() * 2
                
        .Direction = D3DVec(0, 0, 0)
        
        .Range = 10
        .Attenuation1 = 0.3   'Set the Linear Attenuation to 0.05
    End With

    D3DDevice.SetLight 0, Light
    D3DDevice.LightEnable 0, 1

End Sub

Private Function createNormVert(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal nX As Single, ByVal nY As Single, ByVal nZ As Single) As NORMALVERTEX

With createNormVert
    .X = X
    .Y = Y
    .Z = Z
    .nX = nX
    .nY = nY
    .nZ = nZ
End With

End Function

