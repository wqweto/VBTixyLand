VERSION 5.00
Begin VB.UserControl TixyLand 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   DrawStyle       =   2  'Dot
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
End
Attribute VB_Name = "TixyLand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VB6 TixyLand Control (c) 2020 by wqweto@gmail.com
'
' Based on the original idea of https://tixy.land
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "TixyLand"

#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)

'=========================================================================
' Public events
'=========================================================================

Event Click()
Event OwnerDraw(ByVal hGraphics As Long)
Event DblClick()
Event ContextMenu()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event ScriptError()

'=========================================================================
' API
'=========================================================================

'--- DIB Section constants
Private Const DIB_RGB_COLORS                As Long = 0 '  color table in RGBs
Private Const AC_SRC_ALPHA                  As Long = 1
'--- for GdipSetSmoothingMode
Private Const SmoothingModeAntiAlias        As Long = 4
'--- for Modern Subclassing Thunk (MST)
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'--- GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal lFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal sngX As Single, ByVal sngY As Single, ByVal sngWidth As Single, ByVal sngHeight As Single) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal lArgb As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lSmoothingMd As Long) As Long
'--- for Modern Subclassing Thunk (MST)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
'--- end MST

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_AUTOREDRAW        As Boolean = True
Private Const DEF_OPACITY           As Single = 1
Private Const DEF_MATRIXSIZE        As Long = 16
Private Const DEF_EXPRESSION        As String = vbNullString
Private Const STR_FUNC_TEMPLATE     As String = "function %1(t, i, x, y) { with (Math) { return(" & vbCrLf & "%2) } }"
Private Const STR_POLYFILL          As String = "function hypot(x, y) { return Math.sqrt(x*x + y*y) }"

Private m_bAutoRedraw           As Boolean
Private m_sngOpacity            As Single
Private m_lMatrixSize           As Long
Private m_sExpression           As String
'--- run-time
Private m_eContainerScaleMode   As ScaleModeConstants
Private m_bShown                As Boolean
Private m_hAttributes           As Long
Private m_hRedrawDib            As Long
Private m_nDownButton           As Integer
Private m_nDownShift            As Integer
Private m_sngDownX              As Single
Private m_sngDownY              As Single
Private m_sLastError            As String
Private m_aMatrix()             As Single
Private m_dblStartTime          As Double
Private m_pTimer                As IUnknown
Private m_uScript               As UcsActiveScriptData

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    m_sLastError = Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get AutoRedraw() As Boolean
    AutoRedraw = m_bAutoRedraw
End Property

Property Let AutoRedraw(ByVal bValue As Boolean)
    If m_bAutoRedraw <> bValue Then
        m_bAutoRedraw = bValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Opacity() As Single
    Opacity = m_sngOpacity
End Property

Property Let Opacity(ByVal sngValue As Single)
    If m_sngOpacity <> sngValue Then
        m_sngOpacity = IIf(sngValue > 1, 1, IIf(sngValue < 0, 0, sngValue))
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get MatrixSize() As Long
    MatrixSize = m_lMatrixSize
End Property

Property Let MatrixSize(ByVal lValue As Long)
    If m_lMatrixSize <> lValue Then
        m_lMatrixSize = IIf(lValue < 1, 1, lValue)
        pvResetMatrix
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Expression() As String
    Expression = m_sExpression
End Property

Property Let Expression(sValue As String)
    If m_sExpression <> sValue Then
        m_sLastError = vbNullString
        On Error GoTo EH
        ActiveScriptRunCode m_uScript, Replace(Replace(STR_FUNC_TEMPLATE, "%1", "__eval__"), "%2", sValue)
        On Error GoTo 0
        m_sExpression = sValue
        pvResetMatrix
        pvRefresh
        PropertyChanged
    End If
    Exit Property
EH:
    ActiveScriptRunCode m_uScript, Replace(Replace(STR_FUNC_TEMPLATE, "%1", "__eval__"), "%2", "0")
End Property

Property Get LastError() As String
     LastError = m_sLastError
End Property

Private Property Get pvAddressOfTimerProc() As TixyLand
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Refresh()
    Const FUNC_NAME     As String = "Refresh"
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(0)
        If hMemDC = 0 Then
            GoTo QH
        End If
        If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
            GoTo QH
        End If
        hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        pvPaintControl hMemDC
    End If
    UserControl.Refresh
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Repaint()
    Const FUNC_NAME     As String = "Repaint"
    
    On Error GoTo EH
    If m_bShown Then
        pvPrepareMatrix TimerEx - m_dblStartTime, m_aMatrix
        pvPrepareAttribs m_sngOpacity, m_hAttributes
        pvRefresh
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Repaint
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc(), Delay:=15)
End Function

Public Sub OnActiveScriptError(sDescription As String, sSourceLine As String, ByVal lLine As Long, ByVal lPos As Long)
Attribute OnActiveScriptError.VB_MemberFlags = "40"
    #If sSourceLine Then '--- touch args
    #End If
    m_sLastError = sDescription & " at position " & ((lLine * 45 + lPos) - 45)
    RaiseEvent ScriptError
End Sub

Public Function OnActiveScriptGetWindow() As Long
Attribute OnActiveScriptGetWindow.VB_MemberFlags = "40"
    OnActiveScriptGetWindow = ContainerHwnd
End Function

'= private ===============================================================

Private Sub pvResetMatrix()
    ReDim m_aMatrix(-1 To -1) As Single
    m_dblStartTime = TimerEx
    If Ambient.UserMode Then
        Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc(), Delay:=17)
    End If
End Sub

Private Function pvPrepareMatrix(ByVal sngElapsed As Single, aMatrix() As Single) As Boolean
    Const FUNC_NAME     As String = "pvPrepareMatrix"
    Dim lIdx            As Long
    
    On Error GoTo EH
    ReDim aMatrix(0 To m_lMatrixSize * m_lMatrixSize - 1) As Single
    For lIdx = 0 To m_lMatrixSize * m_lMatrixSize - 1
        m_aMatrix(lIdx) = pvEvalCellValue(sngElapsed, lIdx, lIdx Mod m_lMatrixSize, lIdx \ m_lMatrixSize)
    Next
    '--- success
    pvPrepareMatrix = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvEvalCellValue(ByVal sngT As Single, ByVal lIdx As Long, ByVal lX As Long, ByVal lY As Long) As Single
    Dim vResult         As Variant
    
    On Error GoTo QH
    vResult = ActiveScriptCallFunction(m_uScript, "__eval__", sngT, lIdx, lX, lY)
    If VarType(vResult) = vbBoolean Then
        pvEvalCellValue = -vResult
    Else
        pvEvalCellValue = vResult + 0
    End If
QH:
End Function

Private Sub pvRefresh()
    m_bShown = False
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    UserControl.Refresh
End Sub

Private Function pvPaintControl(ByVal hDC As Long) As Boolean
    Const FUNC_NAME     As String = "pvPaintControl"
    Dim hGraphics       As Long
    Dim lIdx            As Long
    Dim hWhiteBrush     As Long
    Dim hRedBrush       As Long
    Dim sngStepX        As Single
    Dim sngStepY        As Single
    Dim sngValue        As Single
    Dim sngOffsetX      As Single
    Dim sngOffsetY      As Single
    
    On Error GoTo EH
    If Not m_bShown Then
        m_bShown = True
        pvPrepareMatrix TimerEx - m_dblStartTime, m_aMatrix
        pvPrepareAttribs m_sngOpacity, m_hAttributes
    End If
    If CheckFailed(GdipCreateFromHDC(hDC, hGraphics)) Then
        GoTo QH
    End If
    If CheckFailed(GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)) Then
        GoTo QH
    End If
    If CheckFailed(GdipCreateSolidFill(&HFFFFFFFF, hWhiteBrush)) Then
        GoTo QH
    End If
    If CheckFailed(GdipCreateSolidFill(&HFFFF2030, hRedBrush)) Then
        GoTo QH
    End If
    sngStepX = ScaleWidth / m_lMatrixSize
    sngStepY = ScaleHeight / m_lMatrixSize
    For lIdx = 0 To m_lMatrixSize * m_lMatrixSize - 1
        sngValue = (1 - Abs(Clamp(m_aMatrix(lIdx), -1, 1))) * 0.95 + 0.05
        sngOffsetX = sngValue * sngStepX
        sngOffsetY = sngValue * sngStepY
        If CheckFailed(GdipFillEllipse(hGraphics, IIf(m_aMatrix(lIdx) >= 0, hWhiteBrush, hRedBrush), _
                sngStepX * (lIdx Mod m_lMatrixSize) + sngOffsetX / 2, _
                sngStepY * (lIdx \ m_lMatrixSize) + sngOffsetY / 2, _
                sngStepX - sngOffsetX, _
                sngStepY - sngOffsetY)) Then
            GoTo QH
        End If
    Next
    RaiseEvent OwnerDraw(hGraphics)
    '--- success
    pvPaintControl = True
QH:
    On Error Resume Next
    If hRedBrush <> 0 Then
        Call GdipDeleteBrush(hRedBrush)
        hRedBrush = 0
    End If
    If hWhiteBrush <> 0 Then
        Call GdipDeleteBrush(hWhiteBrush)
        hWhiteBrush = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareAttribs(ByVal sngAlpha As Single, hAttributes As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareAttribs"
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim hNewAttributes  As Long
    
    On Error GoTo EH
    If CheckFailed(GdipCreateImageAttributes(hNewAttributes)) Then
        GoTo QH
    End If
    clrMatrix(0, 0) = 1
    clrMatrix(1, 1) = 1
    clrMatrix(2, 2) = 1
    clrMatrix(3, 3) = sngAlpha
    clrMatrix(4, 4) = 1
    If CheckFailed(GdipSetImageAttributesColorMatrix(hNewAttributes, 0, 1, clrMatrix(0, 0), clrMatrix(0, 0), 0)) Then
        GoTo QH
    End If
    '--- commit
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
        hAttributes = 0
    End If
    hAttributes = hNewAttributes
    hNewAttributes = 0
    '--- success
    pvPrepareAttribs = True
QH:
    On Error Resume Next
    If hNewAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hNewAttributes)
        hNewAttributes = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

'= common ================================================================

Private Function pvCreateDib(ByVal hMemDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, hDib As Long, Optional lpBits As Long) As Boolean
    Const FUNC_NAME     As String = "pvCreateDib"
    Dim uHdr            As BITMAPINFOHEADER
    
    On Error GoTo EH
    With uHdr
        .biSize = Len(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lWidth
        .biHeight = -lHeight
        .biSizeImage = 4 * lWidth * lHeight
    End With
    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
    If hDib = 0 Then
        GoTo QH
    End If
    '--- success
    pvCreateDib = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function ToScaleMode(sScaleUnits As String) As ScaleModeConstants
    Select Case sScaleUnits
    Case "Twip"
        ToScaleMode = vbTwips
    Case "Point"
        ToScaleMode = vbPoints
    Case "Pixel"
        ToScaleMode = vbPixels
    Case "Character"
        ToScaleMode = vbCharacters
    Case "Centimeter"
        ToScaleMode = vbCentimeters
    Case "Millimeter"
        ToScaleMode = vbMillimeters
    Case "Inch"
        ToScaleMode = vbInches
    Case Else
        ToScaleMode = vbTwips
    End Select
End Function

Private Sub pvHandleMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_nDownButton = Button
    m_nDownShift = Shift
    m_sngDownX = x
    m_sngDownY = y
End Sub

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function Clamp(ByVal sngValue As Single, ByVal sngMin As Single, ByVal sngMax As Single) As Single
    Select Case True
    Case sngValue < sngMin
        Clamp = sngMin
    Case sngValue > sngMax
        Clamp = sngMax
    Case Else
        Clamp = sngValue
    End Select
End Function

Private Function CheckFailed(ByVal lResult As Long) As Boolean
    If lResult <> 0 Then
        CheckFailed = True
        m_sLastError = "GDI+ error " & lResult
    End If
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(x, ScaleMode, m_eContainerScaleMode), ScaleY(y, ScaleMode, m_eContainerScaleMode))
    pvHandleMouseDown Button, Shift, x, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(x, ScaleMode, m_eContainerScaleMode), ScaleY(y, ScaleMode, m_eContainerScaleMode))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseUp"
    
    On Error GoTo EH
    RaiseEvent MouseUp(Button, Shift, ScaleX(x, ScaleMode, m_eContainerScaleMode), ScaleY(y, ScaleMode, m_eContainerScaleMode))
    If Button = -1 Then
        GoTo QH
    End If
    If Button <> 0 And x >= 0 And x < ScaleWidth And y >= 0 And y < ScaleHeight Then
        If (m_nDownButton And Button And vbLeftButton) <> 0 Then
            RaiseEvent Click
        ElseIf (m_nDownButton And Button And vbRightButton) <> 0 Then
            RaiseEvent ContextMenu
        End If
    End If
    m_nDownButton = 0
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_DblClick()
    pvHandleMouseDown vbLeftButton, m_nDownShift, m_sngDownX, m_sngDownY
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Resize()
    pvRefresh
End Sub

Private Sub UserControl_Hide()
    m_bShown = False
End Sub

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"
    Const Opacity       As Long = &HFF
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(hDC)
        If hMemDC = 0 Then
            GoTo DefPaint
        End If
        If m_hRedrawDib = 0 Then
            If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
                GoTo DefPaint
            End If
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
            If Not pvPaintControl(hMemDC) Then
                GoTo DefPaint
            End If
        Else
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        End If
        If AlphaBlend(hDC, 0, 0, ScaleWidth, ScaleHeight, hMemDC, 0, 0, ScaleWidth, ScaleHeight, AC_SRC_ALPHA * &H1000000 + Opacity * &H10000) = 0 Then
            GoTo DefPaint
        End If
    Else
        If Not pvPaintControl(hDC) Then
            GoTo DefPaint
        End If
    End If
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
DefPaint:
    If m_hRedrawDib <> 0 Then
        '--- note: before deleting DIB try de-selecting from dc
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), vbBlack, B
    GoTo QH
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    m_bAutoRedraw = DEF_AUTOREDRAW
    m_sngOpacity = DEF_OPACITY
    m_lMatrixSize = DEF_MATRIXSIZE
    m_sExpression = DEF_EXPRESSION
    pvResetMatrix
    pvRefresh
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    With PropBag
        m_bAutoRedraw = .ReadProperty("AutoRedraw", DEF_AUTOREDRAW)
        m_sngOpacity = .ReadProperty("Opacity", DEF_OPACITY)
        m_lMatrixSize = .ReadProperty("MatrixSize", DEF_MATRIXSIZE)
        m_sExpression = .ReadProperty("Expression", DEF_EXPRESSION)
    End With
    pvResetMatrix
    pvRefresh
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "AutoRedraw", m_bAutoRedraw, DEF_AUTOREDRAW
        .WriteProperty "Opacity", m_sngOpacity, DEF_OPACITY
        .WriteProperty "MatrixSize", m_lMatrixSize, DEF_MATRIXSIZE
        .WriteProperty "Expression", m_sExpression, DEF_EXPRESSION
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'Private Sub UserControl_AmbientChanged(PropertyName As String)
'    If PropertyName = "ScaleUnits" Then
'        m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
'    End If
'End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Initialize()
    Dim aInput(0 To 3)  As Long
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    m_eContainerScaleMode = vbTwips
    On Error Resume Next
    If Not ActiveScriptInit(m_uScript, "JScript9", Me) Then
        ActiveScriptInit m_uScript, "JScript", Me
    End If
    On Error GoTo 0
    ActiveScriptRunCode m_uScript, STR_POLYFILL
End Sub

Private Sub UserControl_Terminate()
    If m_hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hAttributes)
        m_hAttributes = 0
    End If
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    ActiveScriptTerminate m_uScript
End Sub

'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgeogEQUAV1aLdCQUg8YIgz4AdCqL+oHHDBMFAIvCBSgSBQCri8IFZBIFAKuLwgV0EgUAqzPAq7kIAAAA86WBwgwTBQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItDBMFAAWgEgUAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g8EwUALSARBQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBItCEIXAdFuLCotBKIXAdCdS/9Bag/gBd0mLClL/USxahcB1PosKUmrw/3Eg/1EkWqkAAAAIdSuLClL/cghqAP9RHFr/QgQzwFBU/3IQ/1IUi1QkCMdCCAAAAABS6G////9YwhQADx8AjURAAQ==" ' 13.5.2020 18:59:12
    Const THUNK_SIZE    As Long = 5660
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
