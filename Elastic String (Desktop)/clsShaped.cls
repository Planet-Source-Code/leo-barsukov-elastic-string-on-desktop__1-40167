VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShaped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Dim hRgn As Long

Public Sub Window(hwnd As Long, Pict As Picture, Optional MaskColor As Long = &HFFFFFF)
    hRgn = GetBitmapRegion(Pict, MaskColor)
    SetWindowRgn hwnd, hRgn, True
End Sub

Private Function GetBitmapRegion(Pict As Picture, MaskColor As Long)
    Dim hRgn As Long
    Dim tRgn As Long
    Dim X As Integer, Y As Integer, X0 As Integer
    Dim hDC As Long, BM As BITMAP
    
    hDC = CreateCompatibleDC(0)
    If hDC Then
        SelectObject hDC, Pict
        GetObject Pict, Len(BM), BM
        hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        For Y = 0 To BM.bmHeight
            For X = 0 To BM.bmWidth
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> MaskColor
                    X = X + 1
                Wend
                X0 = X
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) = MaskColor
                    X = X + 1
                Wend
                If X0 < X Then
                    tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                    CombineRgn hRgn, hRgn, tRgn, 4
                    DeleteObject tRgn
                End If
            Next X
        Next Y
        GetBitmapRegion = hRgn
        DeleteObject SelectObject(hDC, Pict)
    End If
    
    DeleteDC hDC
End Function

Private Sub Class_Terminate()
    If hRgn Then DeleteObject hRgn
End Sub


