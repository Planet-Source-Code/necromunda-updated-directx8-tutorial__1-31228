Attribute VB_Name = "mGeometry"
Option Explicit

'#Requires
'# Direct3DVertexBuffer8
'# That initD3D has been run

'# Constant for untransformed and lit vertices - XYZ and DIFFUSE
Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'# Type for lit vertex
Type VERTEX
  X As Single
  Y As Single
  Z As Single
  Colour As Long
End Type

'-------------------------
'#initGeometry

'# Creates 24 vertices (8 triangles) in an octahedral
'# shape.  The vertices are also coloured.
'# Then loads the vertices into a Vertex Buffer

'#initMatrix

'# Initialises the view and projection matrices.
'# Sets the viewport (eye) to look at (0,0,0) with up set
'# as (0,1,0).  The projection is set to a ratio of pi/4
'# with a range of 1 to 10 metres.

'#createVertex
'# Creates a lit vertex

'#D3DVec
'# Creates a D3DVECTOR
'-------------------------

'Now, here's where we make up our luvly little shapes, using equally nice
'vectors.  Vectors are great wee things that tell us where we are in
'3D space, using an X, Y and Z value. A point in 3D space (like the corner
'of a cube) is known as a vertex.

'Eg.
'Using the normal type co-ordinate system (X across the screen,
'Y up the screen, Z into the screen):

'verts(0) = createVertex(-1, 0, -1, vbBlue) means go -1 in the X-axis,
'0 in the Y-axis, and -1 in the Z-axis, then make the point there.
'And make it blue.

'As we are making an octahedron (think of 2 pyramids
'stuck base to base) here we need to make 24 vectors. This works out as
'8 triangles - 3 verticies per triangle.

'For a cube, we would need 36 verticies
'(think about it - cube has 6 faces, but each face is made up of 2 triangles.
'so that makes 12 triangles = 36 verticies). It really helps if you work out
'the points that you need beforehand, using pen and paper...


'Right. That's the basic theory explained...now the VB stuff.
'In the declarations section you'll notice a type (FVF_VERTEX) and
'a constant (VERTEX).  These describe what type of vertex that you're
'gonna be using, using the Flexible Vertex Format system in DirectX.
'This lovely system means that you can mix and match your vertex types.
'








Sub initGeometry()
Dim verts(23) As VERTEX
    Dim rt2 As Single
    
    rt2 = Sqr(2)

    verts(0) = createVertex(-1, 0, -1, vbBlue)
    verts(1) = createVertex(0, rt2, 0, vbCyan)
    verts(2) = createVertex(1, 0, -1, vbRed)
    
    verts(3) = createVertex(-1, 0, 1, vbRed)
    verts(4) = createVertex(0, rt2, 0, vbCyan)
    verts(5) = createVertex(1, 0, 1, vbBlue)
    
    verts(6) = createVertex(1, 0, 1, vbBlue)
    verts(7) = createVertex(0, rt2, 0, vbCyan)
    verts(8) = createVertex(1, 0, -1, vbRed)
    
    verts(9) = createVertex(-1, 0, 1, vbRed)
    verts(10) = createVertex(0, rt2, 0, vbCyan)
    verts(11) = createVertex(-1, 0, -1, vbBlue)
    
    verts(12) = createVertex(-1, 0, -1, vbBlue)
    verts(13) = createVertex(0, -rt2, 0, vbCyan)
    verts(14) = createVertex(1, 0, -1, vbRed)
    
    verts(15) = createVertex(-1, 0, 1, vbRed)
    verts(16) = createVertex(0, -rt2, 0, vbCyan)
    verts(17) = createVertex(1, 0, 1, vbBlue)
    
    verts(18) = createVertex(1, 0, 1, vbBlue)
    verts(19) = createVertex(0, -rt2, 0, vbCyan)
    verts(20) = createVertex(1, 0, -1, vbRed)
    
    verts(21) = createVertex(-1, 0, 1, vbRed)
    verts(22) = createVertex(0, -rt2, 0, vbCyan)
    verts(23) = createVertex(-1, 0, -1, vbBlue)
    
    D3DDevice.SetVertexShader FVF_VERTEX
    
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
           
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 24, _
                                                    0, _
                                                    FVF_VERTEX, _
                                                    D3DPOOL_DEFAULT)
        
    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 24, 0, verts(0)
    

End Sub

Sub initMatrix()
    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX
        
    D3DXMatrixLookAtLH matView, _
        D3DVec(0#, 3#, -5.5), _
        D3DVec(0#, 0#, 0#), _
        D3DVec(0#, 1#, 0#)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    
    D3DXMatrixPerspectiveFovLH matProj, _
        PI / 4, _
        1, _
        1, 10
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj

End Sub

Private Function createVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Colour As Long) As VERTEX

    With createVertex
        .X = X
        .Y = Y
        .Z = Z
        .Colour = Colour
    End With

End Function

Function D3DVec(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR

    With D3DVec
        .X = Y
        .Y = Y
        .Z = Z
    End With

End Function
