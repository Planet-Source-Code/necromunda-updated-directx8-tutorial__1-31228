VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectX"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8640
      Top             =   4920
   End
   Begin VB.ListBox lstReport 
      Height          =   1425
      Left            =   6240
      TabIndex        =   7
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Frame fmeExtra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extra Options"
      Height          =   2775
      Left            =   6120
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
      Begin VB.CheckBox chkDrag 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Drag Mode"
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkZ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Z"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Y"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&X"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1920
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.CheckBox chkMove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Translate"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox chkScale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Scale"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkRotate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Rotate in:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame fmeSetup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Setup Options"
      Height          =   1215
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
      Begin VB.CheckBox chkTexture 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Textures"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkLights 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lighting"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optWindow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Window"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fullscreen"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Unload Objects and Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "St&art Scene"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picDirectX 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   2
      ToolTipText     =   "Click and hold to drag!"
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim FPS_LastCheck As Long
Dim FPS_Count As Long
Dim FPS_Current As Integer


Private Sub Command1_Click()
    Dim vert As TEXVERTEX

    InitD3D True, picDirectX.hWnd
    appRunning = True
    
    Set picDirectX = Nothing
    addReport ("Window DX Scene Loaded!")
    addReport ("--------")
            
    initTexCube
    initMatrix
    
    Timer1.Enabled = True
        
    
    Do
    
    DoEvents
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(225, 100, 100), 1#, 0
    D3DDevice.BeginScene
         
    D3DDevice.SetStreamSource 0, vertexBuffer, Len(vert)
        
    D3DDevice.SetVertexShader FVF_TEX
            
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
       
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

    Loop Until appRunning = False

End Sub

Private Sub Form_Load()
        
    Call Randomize

End Sub
Private Sub Form_Unload(Cancel As Integer)
        
    Call unloadObjects

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
        Select Case KeyCode
            Case Is = vbKeyEscape
                Call unloadObjects
            Case Is = vbKeyR
                chkRotate.Value = IIf(chkRotate.Value = 1, 0, 1)
            Case Is = vbKeyS
                chkScale.Value = IIf(chkScale.Value = 1, 0, 1)
            Case Is = vbKeyT
                chkMove.Value = IIf(chkMove.Value = 1, 0, 1)
            Case Is = vbKeyX
                chkX.Value = IIf(chkX.Value = 1, 0, 1)
            Case Is = vbKeyY
                chkY.Value = IIf(chkY.Value = 1, 0, 1)
            Case Is = vbKeyZ
                chkZ.Value = IIf(chkZ.Value = 1, 0, 1)
        End Select

End Sub


Private Sub cmdChange_Click()

    Call initGeometry

End Sub
Private Sub cmdStart_Click()
                
    lstReport.Clear
    currentApp = notRunning
                
    If InitD3D(optWindow.Value, IIf(optWindow.Value, picDirectX.hWnd, frmFullDisp.hWnd)) = True Then
        'appRunning = True
        'appLighting = IIf((chkLights.Value = 0), False, True)
                
        Set picDirectX = Nothing
        addReport ("Window DX Scene Loaded!")
        addReport ("--------")
        
        If chkLights.Value = vbChecked Then currentApp = lighting
        If chkTexture.Value = vbChecked Then currentApp = Texturing
        If currentApp = notRunning Then currentApp = Normal
                
                
        If currentApp = Normal Then
            initGeometry
        
        ElseIf currentApp = lighting Then
            initLightingGeo
            SetupLights
            
        ElseIf currentApp = Texturing Then
            initTexCube
                          
        End If
                
        initMatrix
        
        If chkDrag.Value = 0 Then Timer1.Enabled = True
        
        Do While Not currentApp = notRunning
            Call Render
        
            If GetTickCount() - FPS_LastCheck >= 100 Then
                FPS_Current = FPS_Count * 10
                FPS_Count = 0
                FPS_LastCheck = GetTickCount()
            End If
        
            FPS_Count = FPS_Count + 1
            Me.Caption = FPS_Current & "fps"
                     
            'Let Windows take a breath
            DoEvents
        
        Loop
    Else
        MsgBox "Initialisation failed!"
    End If
    
    
End Sub
Private Sub cmdExit_Click()

    Call unloadObjects

End Sub


Private Sub chkDrag_Click()
    
    If chkDrag.Value = 1 Then
        Timer1.Enabled = False
        picDirectX.Enabled = True
    
        chkRotate.Enabled = False
        chkScale.Enabled = False
        chkMove.Enabled = False
        
        chkX.Enabled = False
        chkY.Enabled = False
        chkZ.Enabled = False
    Else
        If currentApp <> notRunning Then Timer1.Enabled = True
        picDirectX.Enabled = False
    
        chkRotate.Enabled = True
        chkScale.Enabled = True
        chkMove.Enabled = True
        
        chkX.Enabled = True
        chkY.Enabled = True
        chkZ.Enabled = True
    End If
        
End Sub
Private Sub chkRotate_Click()
    
    If chkRotate.Value = 1 Then
        chkX.Visible = True: chkY.Visible = True: chkZ.Visible = True
    Else
        chkX.Visible = False: chkY.Visible = False: chkZ.Visible = False
    End If

End Sub


Private Sub optFull_Click()
    
    chkDrag.Enabled = False

    If chkDrag.Value = 1 Then
        chkDrag.Value = 0
        Call chkDrag_Click
    End If

End Sub
Private Sub optWindow_Click()

    chkDrag.Enabled = True

End Sub

Private Sub picDirectX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim matPic As D3DMATRIX
        
    If Button <> 0 Then
        D3DXMatrixRotationYawPitchRoll matPic, _
            (X / picDirectX.ScaleWidth) * 2.5, _
            (Y / picDirectX.ScaleHeight) * 2.5, _
            0
        D3DDevice.SetTransform D3DTS_WORLD, matPic
    End If

End Sub


Private Sub Timer1_Timer()
    Dim matWorld As D3DMATRIX
    Dim matRotate As D3DMATRIX
    Dim matScale As D3DMATRIX
    Dim matTrans As D3DMATRIX
 
    If chkScale.Value = 1 Then
        D3DXMatrixScaling matScale, _
             Abs(Sin(Timer)), _
             Abs(Sin(Timer)), _
             Abs(Sin(Timer))
    Else
        D3DXMatrixIdentity matScale
    End If
    
    
    If chkRotate.Value = 1 Then
        D3DXMatrixRotationYawPitchRoll matRotate, _
            IIf(chkX.Value = 1, 2 * Sin(Timer), 0), _
            IIf(chkY.Value = 1, 0.8 * Timer, 0), _
            IIf(chkZ.Value = 1, Timer, 0)
    Else
        D3DXMatrixIdentity matRotate
    End If
    
    
    If chkMove.Value = 1 Then
        D3DXMatrixTranslation matTrans, _
            Cos(Timer), _
            Sin(Timer), _
            0
    Else
        D3DXMatrixIdentity matTrans
    End If
            
    D3DXMatrixMultiply matWorld, matRotate, matTrans
    D3DXMatrixMultiply matWorld, matScale, matWorld
         
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
       
End Sub

