Attribute VB_Name = "mRender"
Option Explicit

Public Sub Render()
    Dim vertexSize As Single
    
    Dim sizeVert As VERTEX
    Dim sizeNormVert As NORMALVERTEX
    Dim sizeTexVert As TEXVERTEX
        
    If currentApp = Normal Then
        vertexSize = Len(sizeVert)
    ElseIf currentApp = lighting Then
        vertexSize = Len(sizeNormVert)
    ElseIf currentApp = Texturing Then
        vertexSize = Len(sizeTexVert)
    End If
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(225, 100, 100), 1#, 0
    D3DDevice.BeginScene
         
    D3DDevice.SetStreamSource 0, vertexBuffer, vertexSize
        
    If currentApp <> Texturing Then
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 8
    Else
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
    End If
    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
End Sub


