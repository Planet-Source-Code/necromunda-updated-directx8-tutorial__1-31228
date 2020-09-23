Attribute VB_Name = "mTextures"
Option Explicit

Public Const FVF_TEX = (D3DFVF_XYZ Or D3DFVF_TEX1)

Type TEXVERTEX
  X As Single
  Y As Single
  Z As Single
  tu As Single
  tv As Single
End Type

Sub initTexCube()
    Dim verts(35) As TEXVERTEX

    verts(0) = createTexVert(-1, 1, -1, 0, 0)
    verts(1) = createTexVert(1, 1, -1, 1, 0)
    verts(2) = createTexVert(-1, 1, 1, 0, 1)
        
    verts(3) = createTexVert(1, 1, -1, 1, 0)
    verts(4) = createTexVert(1, 1, 1, 1, 1)
    verts(5) = createTexVert(-1, 1, 1, 0, 1)
    
    verts(6) = createTexVert(-1, -1, -1, 0, 0)
    verts(7) = createTexVert(1, -1, -1, 1, 0)
    verts(8) = createTexVert(-1, -1, 1, 0, 1)
        
    verts(9) = createTexVert(1, -1, -1, 1, 0)
    verts(10) = createTexVert(1, -1, 1, 1, 1)
    verts(11) = createTexVert(-1, -1, 1, 0, 1)
    
    verts(12) = createTexVert(-1, 1, -1, 0, 0)
    verts(13) = createTexVert(-1, 1, 1, 1, 0)
    verts(14) = createTexVert(-1, -1, -1, 0, 1)
        
    verts(15) = createTexVert(-1, 1, 1, 1, 0)
    verts(16) = createTexVert(-1, -1, 1, 1, 1)
    verts(17) = createTexVert(-1, -1, -1, 0, 1)
    
    verts(18) = createTexVert(1, 1, -1, 0, 0)
    verts(19) = createTexVert(1, 1, 1, 1, 0)
    verts(20) = createTexVert(1, -1, -1, 0, 1)
        
    verts(21) = createTexVert(1, 1, 1, 1, 0)
    verts(22) = createTexVert(1, -1, 1, 1, 1)
    verts(23) = createTexVert(1, -1, -1, 0, 1)
        
    verts(24) = createTexVert(-1, 1, 1, 0, 0)
    verts(25) = createTexVert(1, 1, 1, 1, 0)
    verts(26) = createTexVert(-1, -1, 1, 0, 1)
    
    verts(27) = createTexVert(1, 1, 1, 1, 0)
    verts(28) = createTexVert(1, -1, 1, 1, 1)
    verts(29) = createTexVert(-1, -1, 1, 0, 1)
    
    verts(30) = createTexVert(-1, 1, -1, 0, 0)
    verts(31) = createTexVert(1, 1, -1, 1, 0)
    verts(32) = createTexVert(-1, -1, -1, 0, 1)
        
    verts(33) = createTexVert(1, 1, -1, 1, 0)
    verts(34) = createTexVert(1, -1, -1, 1, 1)
    verts(35) = createTexVert(-1, -1, -1, 0, 1)
    
    D3DDevice.SetVertexShader FVF_TEX
    
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
           
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 36, _
                                                    0, _
                                                    FVF_TEX, _
                                                    D3DPOOL_DEFAULT)
            
    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 36, 0, verts(0)
    
    
    
    Dim DispMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
            
    Set Texture = D3DX.CreateTextureFromFileEx(D3DDevice, _
                            App.Path & "\green.jpeg", _
                            128, 128, _
                            1, 0, _
                            DispMode.Format, _
                            D3DPOOL_MANAGED, _
                            D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                            0, ByVal 0, ByVal 0)
    
    D3DDevice.SetTexture 0, Texture

End Sub

Private Function createTexVert(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal tv As Single, ByVal tu As Single) As TEXVERTEX

    With createTexVert
        .X = X
        .Y = Y
        .Z = Z
        .tu = tu
        .tv = tv
    End With

End Function

