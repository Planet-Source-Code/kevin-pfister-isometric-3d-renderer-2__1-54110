VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HeightMap Render 2"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   654
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   9960
      Picture         =   "FrmMain.frx":1F1DA
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
      Top             =   1320
      Width           =   465
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9960
      Picture         =   "FrmMain.frx":1F652
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   720
      Width           =   480
   End
   Begin VB.CommandButton CmdDemo 
      Caption         =   "Demo"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   7440
      Width           =   1335
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   9960
      Picture         =   "FrmMain.frx":1FB58
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   1
      Top             =   120
      Width           =   465
   End
   Begin VB.PictureBox PicRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7200
      Left            =   120
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Integer


Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const Deg = 3.141592653 / 180

Const BigS = 1000
Const RenderS = 40

Dim BigMap(BigS, BigS) As Integer
Dim RenderMap(RenderS, RenderS) As Integer

Const TW = 640 / 40
Const TH = 480 / 40
Const HW = TW / 2
Const HH = TH / 2

Dim Offset As Integer

Private Sub CmdDemo_Click()
    'Start the Demo
    Do
        'Render the map
        Call RenderHeightMap(1 + Offset, 1 + Offset)
        Offset = Offset + 2
        DoEvents
    Loop Until 40 + Offset + 1 > BigS
    Offset = 0
End Sub

Private Sub Form_Load()
    Randomize Timer
    
    'Default Settings, Altering these will produce different Landscapes
    MultX = 10
    MultY = 50
    YAlt = 0.8
    XAlt = 0.1



    For X = 1 To BigS Step 2
        For Y = 1 To BigS Step 2
            'Create the Landscape
            MathLUT = Sin(Deg * (X / XAlt)) * MultX + Sin(Deg * (Y / YAlt)) * MultY
            
            BigMap(X, Y) = MathLUT
            BigMap(X + 1, Y) = MathLUT
            BigMap(X + 1, Y + 1) = MathLUT
            BigMap(X, Y + 1) = MathLUT
        Next
    Next

    For X = 1 To BigS
        For Y = 1 To BigS
            If BigMap(X, Y) <= 0 Then
                'If its 0 or below it will be water and so make it flat
                BigMap(X, Y) = 0
            End If
        Next
    Next
    
    'Render the First Screen
    Call RenderHeightMap(1, 1)
End Sub

Sub RenderHeightMap(OffsetX As Integer, OffsetY As Integer)
    Dim Before As Long
    Before = Timer
    
    Dim BotTotal As Double
    Dim TopTotal As Double
    Dim X As Integer
    Dim Y As Integer
    
    'Copy the section to the Mini RenderMap for ease of use
    For X = 1 To RenderS
        For Y = 1 To RenderS
            If OffsetX + X > BigS Then
                X1 = OffsetX + X - BigS
            Else
                X1 = OffsetX + X
            End If
            If OffsetY + Y > BigS Then
                Y1 = OffsetY + Y - BigS
            Else
                Y1 = OffsetY + Y
            End If
            RenderMap(X, Y) = BigMap(X1, Y1)
        Next
    Next
    
    Dim RemHeight As Long
    
    'This is the dampener, stops the landscape going to far vertically
    RemHeight = HH * RenderMap(20, 20) / 2
    
    PicRender.Cls
    For Y = 1 To RenderS Step 2
        For X = 1 To RenderS Step 2
            XY = 320 + X * HW - Y * HW
            
            'Do some precalculations
            TL = Y * HH + X * HH - HH * RenderMap(X, Y) + RemHeight
            TR = Y * HH + (X + 2) * HH - HH * RenderMap(X + 1, Y) + RemHeight
            BL = (Y + 2) * HH + X * HH - HH * RenderMap(X, Y + 1) + RemHeight
            BR = (Y + 2) * HH + (X + 2) * HH - HH * RenderMap(X + 1, Y + 1) + RemHeight
            
            'Call The Rastering Function
            TopTotal = (RenderMap(X, Y) + RenderMap(X + 1, Y) + RenderMap(X, Y + 1) + RenderMap(X + 1, Y + 1)) / 4
            If TopTotal > 1 Then
                Call RasterSqr(XY, TL, 320 + (X + 2) * HW - Y * HW, TR, 320 + X * HW - (Y + 2) * HW, BL, 320 + (X + 2) * HW - (Y + 2) * HW, BR, PicRender, PicGrass)
            ElseIf TopTotal > 0 Then
                Call RasterSqr(XY, TL, 320 + (X + 2) * HW - Y * HW, TR, 320 + X * HW - (Y + 2) * HW, BL, 320 + (X + 2) * HW - (Y + 2) * HW, BR, PicRender, PicSand)
            Else
                Call RasterSqr(XY, TL, 320 + (X + 2) * HW - Y * HW, TR, 320 + X * HW - (Y + 2) * HW, BL, 320 + (X + 2) * HW - (Y + 2) * HW, BR, PicRender, PicWater)
            End If
            
            MoveToEx PicRender.hDC, XY, TL, ByVal 0&
            LineTo PicRender.hDC, 320 + (X + 2) * HW - Y * HW, TR
            LineTo PicRender.hDC, 320 + (X + 2) * HW - (Y + 2) * HW, BR
            LineTo PicRender.hDC, 320 + X * HW - (Y + 2) * HW, BL
            LineTo PicRender.hDC, XY, TL
        Next
    Next
    PicRender.Refresh
    
    Debug.Print "Rendertime: " & Timer - Before
End Sub

Sub RasterSqr(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer, ByVal X4 As Integer, ByVal Y4 As Integer, Output As PictureBox, Source As PictureBox)
    'Rastering, quick simple and could be optimized by loads if needed
    
    Dim SmallX As Double
    Dim SmallY As Double
    Dim BigX As Double
    Dim BigY As Double
    
    Dim X As Double
    Dim Y As Double
    
    Dim XCor1 As Double
    Dim XCor2 As Double
    Dim XCor3 As Double
    Dim XCor4 As Double
    
    Dim YCor1 As Double
    Dim YCor2 As Double
    Dim YCor3 As Double
    Dim YCor4 As Double
    
    SmallX = X1
    SmallY = Y1
    If X2 < SmallX Then
        SmallX = X2
    End If
    If Y2 < SmallY Then
        SmallY = Y2
    End If
    If X3 < SmallX Then
        SmallX = X3
    End If
    If Y3 < SmallY Then
        SmallY = Y3
    End If
    If X4 < SmallX Then
        SmallX = X4
    End If
    If Y4 < SmallY Then
        SmallY = Y4
    End If
    
    X1 = X1 - SmallX
    X2 = X2 - SmallX
    X3 = X3 - SmallX
    X4 = X4 - SmallX
    
    Y1 = Y1 - SmallY
    Y2 = Y2 - SmallY
    Y3 = Y3 - SmallY
    Y4 = Y4 - SmallY
    
    BigX = X1
    BigY = Y1
    If X2 > BigX Then
        BigX = X2
    End If
    If Y2 > BigY Then
        BigY = Y2
    End If
    If X3 > BigX Then
        BigX = X3
    End If
    If Y3 > BigY Then
        BigY = Y3
    End If
    If X4 > BigX Then
        BigX = X4
    End If
    If Y4 > BigY Then
        BigY = Y4
    End If
    ReDim RasterArray(BigX, BigY) As Integer
    
    XCor1 = X1
    XCor2 = X2
    YCor1 = Y1
    YCor2 = Y2
    
    YDiff = YCor2 - YCor1
    XDiff = XCor2 - XCor1
    If YDiff = 0 Then
        If XCor2 > XCor1 Then
            For X = XCor1 To XCor2
                RasterArray(X, YCor1) = 1
            Next
        Else
            For X = XCor1 To XCor2 Step -1
                RasterArray(X, YCor1) = 1
            Next
        End If
    ElseIf XDiff = 0 Then
        If YCor2 > YCor1 Then
            For Y = YCor1 To YCor2
                RasterArray(XCor1, Y) = 1
            Next
        Else
            For Y = YCor1 To YCor2 Step -1
                RasterArray(XCor1, Y) = 1
            Next
        End If
    Else
        If Abs(YDiff) > Abs(XDiff) Then
            XInt = XDiff / YDiff
            If YCor1 < YCor2 Then
                For Y = YCor1 To YCor2
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            Else
                For Y = YCor1 To YCor2 Step -1
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            End If
        ElseIf Abs(XDiff) > Abs(YDiff) Then
            YInt = YDiff / XDiff
            If XCor1 < XCor2 Then
                For X = XCor1 To XCor2
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            Else
                For X = XCor1 To XCor2 Step -1
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            End If
        Else
            YInt = YDiff / XDiff
            For X = XCor1 To XCor2
                RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
            Next
        End If
    End If

    XCor1 = X2
    XCor2 = X4
    YCor1 = Y2
    YCor2 = Y4
    
    YDiff = YCor2 - YCor1
    XDiff = XCor2 - XCor1
    If YDiff = 0 Then
        If XCor2 > XCor1 Then
            For X = XCor1 To XCor2
                RasterArray(X, YCor1) = 1
            Next
        Else
            For X = XCor1 To XCor2 Step -1
                RasterArray(X, YCor1) = 1
            Next
        End If
    ElseIf XDiff = 0 Then
        If YCor2 > YCor1 Then
            For Y = YCor1 To YCor2
                RasterArray(XCor1, Y) = 1
            Next
        Else
            For Y = YCor1 To YCor2 Step -1
                RasterArray(XCor1, Y) = 1
            Next
        End If
    Else
        If Abs(YDiff) > Abs(XDiff) Then
            XInt = XDiff / YDiff
            If YCor1 < YCor2 Then
                For Y = YCor1 To YCor2
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            Else
                For Y = YCor1 To YCor2 Step -1
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            End If
        ElseIf Abs(XDiff) > Abs(YDiff) Then
            YInt = YDiff / XDiff
            If XCor1 < XCor2 Then
                For X = XCor1 To XCor2
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            Else
                For X = XCor1 To XCor2 Step -1
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            End If
        Else
            YInt = YDiff / XDiff
            For X = XCor1 To XCor2
                RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
            Next
        End If
    End If

    XCor1 = X3
    XCor2 = X4
    YCor1 = Y3
    YCor2 = Y4
    
    YDiff = YCor2 - YCor1
    XDiff = XCor2 - XCor1
    If YDiff = 0 Then
        If XCor2 > XCor1 Then
            For X = XCor1 To XCor2
                RasterArray(X, YCor1) = 1
            Next
        Else
            For X = XCor1 To XCor2 Step -1
                RasterArray(X, YCor1) = 1
            Next
        End If
    ElseIf XDiff = 0 Then
        If YCor2 > YCor1 Then
            For Y = YCor1 To YCor2
                RasterArray(XCor1, Y) = 1
            Next
        Else
            For Y = YCor1 To YCor2 Step -1
                RasterArray(XCor1, Y) = 1
            Next
        End If
    Else
        If Abs(YDiff) > Abs(XDiff) Then
            XInt = XDiff / YDiff
            If YCor1 < YCor2 Then
                For Y = YCor1 To YCor2
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            Else
                For Y = YCor1 To YCor2 Step -1
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            End If
        ElseIf Abs(XDiff) > Abs(YDiff) Then
            YInt = YDiff / XDiff
            If XCor1 < XCor2 Then
                For X = XCor1 To XCor2
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            Else
                For X = XCor1 To XCor2 Step -1
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            End If
        Else
            YInt = YDiff / XDiff
            For X = XCor1 To XCor2
                RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
            Next
        End If
    End If
    
    XCor1 = X1
    XCor2 = X3
    YCor1 = Y1
    YCor2 = Y3
    
    YDiff = YCor2 - YCor1
    XDiff = XCor2 - XCor1
    If YDiff = 0 Then
        If XCor2 > XCor1 Then
            For X = XCor1 To XCor2
                RasterArray(X, YCor1) = 1
            Next
        Else
            For X = XCor1 To XCor2 Step -1
                RasterArray(X, YCor1) = 1
            Next
        End If
    ElseIf XDiff = 0 Then
        If YCor2 > YCor1 Then
            For Y = YCor1 To YCor2
                RasterArray(XCor1, Y) = 1
            Next
        Else
            For Y = YCor1 To YCor2 Step -1
                RasterArray(XCor1, Y) = 1
            Next
        End If
    Else
        If Abs(YDiff) > Abs(XDiff) Then
            XInt = XDiff / YDiff
            If YCor1 < YCor2 Then
                For Y = YCor1 To YCor2
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            Else
                For Y = YCor1 To YCor2 Step -1
                    RasterArray(XCor1 + XInt * (Y - YCor1), Y) = 1
                Next
            End If
        ElseIf Abs(XDiff) > Abs(YDiff) Then
            YInt = YDiff / XDiff
            If XCor1 < XCor2 Then
                For X = XCor1 To XCor2
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            Else
                For X = XCor1 To XCor2 Step -1
                    RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
                Next
            End If
        Else
            YInt = YDiff / XDiff
            For X = XCor1 To XCor2
                RasterArray(X, YCor1 + YInt * (X - XCor1)) = 1
            Next
        End If
    End If
    
    For Y = 0 To BigY
        If Y + SmallY >= 0 Then
            For X = 0 To BigX
                If RasterArray(X, Y) = 1 Then
                    Call SetPixelV(Output.hDC, X + SmallX, Y + SmallY, RGB(255, 140, 10))
                    For X1 = BigX To X + 1 Step -1
                        If RasterArray(X1, Y) = 1 Then
                            Call SetPixelV(Output.hDC, X1 + SmallX, Y + SmallY, GetPixel(Source.hDC, X1 Mod 30, Y Mod 30))
                            For X2 = X To X1
                                Call SetPixelV(Output.hDC, X2 + SmallX, Y + SmallY, GetPixel(Source.hDC, X2 Mod 30, Y Mod 30))
                            Next
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Make sure the Program is exited
    End
End Sub
