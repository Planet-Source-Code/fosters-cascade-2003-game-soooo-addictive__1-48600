Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShowCursor Lib "user32" (ByVal fShow As Integer) As Integer
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

'***********************************************************************************
'**                                                                               **
'** If multiple people in the office are playing this, set the directory to a     **
'** common one                                                                    **
'**                                                                               **
          Public Const HiScoreFile As String = "c:\" '"P:\utilities\RichT\CHS.ini"
'**                                                                               **
'***********************************************************************************

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Public Type POINTAPI
    X As Long
    y As Long
End Type
Public Type UDTStar 'stars travel from offset 0,0 at an angle,
                    'speed is dependant on length!
    Speed As Long
    Angle As Single
    Len As Long
    Color As Long
End Type

Public Stars() As UDTStar

Public lCentreOfPicture(2) As Long
Public lMaxLength As Long 'the shortest distance of width or height of the screen
Public Const lNumStars As Long = 300 'set number of stars here
Public Const Pi As Single = 3.14159265358979
Public Function AdjustBrightness(ByRef RGB_In As Long, ByRef ShiftPercentage As Integer, Optional GotoWhite As Boolean = False) As Long
Dim lColor As Long
Dim r As Single, G As Single, B As Single

    lColor = RGB_In
    r = lColor Mod &H100
    lColor = lColor \ &H100
    G = lColor Mod &H100
    lColor = lColor \ &H100
    B = lColor Mod &H100

    r = r + ((r / 100) * ShiftPercentage)
    G = G + ((G / 100) * ShiftPercentage)
    B = B + ((B / 100) * ShiftPercentage)
    
    If r > 255 Or G > 255 Or B > 255 Then
        If GotoWhite Then
            If r > 255 Then r = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            AdjustBrightness = RGB(r, G, B)
        Else
            AdjustBrightness = RGB_In
        End If
    ElseIf r < 0 Or G < 0 Or B < 0 Then
        AdjustBrightness = RGB_In
    Else
        AdjustBrightness = RGB(r, G, B)
    End If
End Function

Function GimmeX(ByVal aIn As Single, lIn As Long) As Integer
    'from an angle and length, give the x axis co'ordinate
    GimmeX = sIn(aIn * (Pi / 180)) * lIn
End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Integer
    'from an angle and length, give the y axis co'ordinate
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function

Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function
Function writeINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Function FileExists(filename As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(filename) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
Function Sine(Degrees_Arg)
Sine = sIn(Degrees_Arg * Atn(1) / 45)
End Function

Function Cosine(Degrees_Arg)
Cosine = Cos(Degrees_Arg * Atn(1) / 45)
End Function


