Attribute VB_Name = "TixyScript"
'=========================================================================
'
' VB6 TixyLand Control (c) 2020 by wqweto@gmail.com
'
' Based on the original idea of https://tixy.land
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "TixyScript"

'=========================================================================
' API
'=========================================================================

Private Const E_NOTIMPL                 As Long = &H80004001
Private Const E_NOINTERFACE             As Long = &H80004002
Private Const TYPE_E_ELEMENTNOTFOUND    As Long = &H8002802B

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CLSIDFromProgID Lib "ole32" (ByVal lpszProgID As Long, ByRef lpclsid As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (rClsID As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, rIID As Any, ppv As Any) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function VariantCopy Lib "oleaut32" (pvarDest As Any, pvargSrc As Any) As Long

Private Type EXCEPINFO
    wCode               As Integer
    wReserved           As Integer
    Source              As String
    Description         As String
    HelpFile            As String
    dwHelpContext       As Long
    pvReserved          As Long
    pfnDeferredFillIn   As Long
    dwSCode             As Long
End Type

Private Type GUID
    Data1               As Currency
    Data2               As Currency
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lVTable(0 To 10)          As Long
Private m_lVTableWindow(0 To 4)     As Long
Private IID_IUnknown                As GUID
Private IID_IActiveScript           As GUID
Private IID_IActiveScriptParse      As GUID
Private IID_IActiveScriptSite       As GUID
Private IID_IActiveScriptSiteWindow As GUID

Public Type UcsActiveScriptData
    pVTable             As Long
    pVTableWindow       As Long
    pSite               As IUnknown ' IActiveScriptSite
    pCallback           As Object
    cObjects            As Collection
    pScript             As IUnknown ' IActiveScript
    pParse              As IUnknown ' IActiveScriptParse
End Type

Private Type UcsActiveScriptSiteWindowData
    pVTableWindow       As Long
    pSite               As IUnknown ' IActiveScriptSite
    pCallback           As Object
End Type

Private Enum UcsInterfaceVTableIndexEnum
    '--- IActiveScript
    IDX_QueryInterface = 0
    IDX_SetScriptSite = 3
    IDX_GetScriptSite = 4
    IDX_SetScriptState = 5
    IDX_GetScriptState = 6
    IDX_Close = 7
    IDX_AddNamedItem = 8
    IDX_AddTypeLib = 9
    IDX_GetScriptDispatch = 10
    '--- IActiveScriptParse
    IDX_InitNew = 3
    IDX_AddScriptlet = 4
    IDX_ParseScriptText = 5
    '--- IActiveScriptError
    IDX_GetExceptionInfo = 3
    IDX_GetSourcePosition = 4
    IDX_GetSourceLineText = 5
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Function ActiveScriptInit(uData As UcsActiveScriptData, sLang As String, pCallback As Object, Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "ActiveScriptInit"
    Const CLSCTX_INPROC_SERVER As Long = 1
    Dim hResult         As Long
    Dim aCLSID          As GUID
    Dim pUnk            As IUnknown
    
    On Error GoTo EH
    '--- perform one-time initializations
    If IID_IUnknown.Data2 = 0 Then
        IID_IUnknown.Data2 = 504403158265495.5712@
        IID_IActiveScript.Data1 = 128342581131674.6977@
        IID_IActiveScript.Data2 = 726435498762961.7295@
        IID_IActiveScriptParse.Data1 = 128342581131674.6978@
        IID_IActiveScriptParse.Data2 = 726435498762961.7295@
        IID_IActiveScriptSite.Data1 = 128342492708874.6979@
        IID_IActiveScriptSite.Data2 = 726435498762961.7295@
        IID_IActiveScriptSiteWindow.Data1 = 128338945908194.6977@
        IID_IActiveScriptSiteWindow.Data2 = 726435498762961.7295@
    End If
    If m_lVTable(0) = 0 Then
        m_lVTable(0) = VBA.CLng(AddressOf IActiveScriptSite_QueryInterface)
        m_lVTable(1) = VBA.CLng(AddressOf IActiveScriptSite_AddRef)
        m_lVTable(2) = VBA.CLng(AddressOf IActiveScriptSite_Release)
        m_lVTable(3) = VBA.CLng(AddressOf IActiveScriptSite_GetLCID)
        m_lVTable(4) = VBA.CLng(AddressOf IActiveScriptSite_GetItemInfo)
        m_lVTable(5) = VBA.CLng(AddressOf IActiveScriptSite_GetDocVersionString)
        m_lVTable(6) = VBA.CLng(AddressOf IActiveScriptSite_OnScriptTerminate)
        m_lVTable(7) = VBA.CLng(AddressOf IActiveScriptSite_OnStateChange)
        m_lVTable(8) = VBA.CLng(AddressOf IActiveScriptSite_OnScriptError)
        m_lVTable(9) = VBA.CLng(AddressOf IActiveScriptSite_OnEnterScript)
        m_lVTable(10) = VBA.CLng(AddressOf IActiveScriptSite_OnLeaveScript)
        m_lVTableWindow(0) = VBA.CLng(AddressOf IActiveScriptSiteWindow_QueryInterface)
        m_lVTableWindow(1) = VBA.CLng(AddressOf IActiveScriptSiteWindow_AddRef)
        m_lVTableWindow(2) = VBA.CLng(AddressOf IActiveScriptSiteWindow_Release)
        m_lVTableWindow(3) = VBA.CLng(AddressOf IActiveScriptSiteWindow_GetWindow)
        m_lVTableWindow(4) = VBA.CLng(AddressOf IActiveScriptSiteWindow_EnableModeless)
    End If
    '--- instantiate scripting engine
    If LCase$(sLang) = "jscript9" Then
        aCLSID.Data1 = 551568143666833.5481@
        aCLSID.Data2 = 618998863008730.077@
    Else
        Call CLSIDFromProgID(StrPtr(sLang), aCLSID)
    End If
    hResult = CoCreateInstance(aCLSID, 0, CLSCTX_INPROC_SERVER, IID_IUnknown, pUnk)
    If hResult < 0 Then
        Err.Raise hResult, "CoCreateInstance(" & sLang & ")"
    End If
    '--- get IActiveScript and IActiveScriptParse interfaces
    Set uData.pScript = Nothing
    hResult = DispCallByVtbl(ObjPtr(pUnk), IDX_QueryInterface, VarPtr(IID_IActiveScript), VarPtr(uData.pScript))
    If hResult < 0 Then
        Err.Raise hResult, "IUnknown.QueryInterface"
    End If
    Set uData.pParse = Nothing
    hResult = DispCallByVtbl(ObjPtr(pUnk), IDX_QueryInterface, VarPtr(IID_IActiveScriptParse), VarPtr(uData.pParse))
    If hResult < 0 Then
        Err.Raise hResult, "IUnknown.QueryInterface"
    End If
    '--- prepare light-weight object for IActiveScriptSite interface
    uData.pVTable = VarPtr(m_lVTable(0))
    uData.pVTableWindow = VarPtr(m_lVTableWindow(0))
    Call CopyMemory(uData.pSite, VarPtr(uData), 4)
    Set uData.pCallback = pCallback
    Set uData.cObjects = New Collection
    '--- init scripting engine
    hResult = DispCallByVtbl(ObjPtr(uData.pScript), IDX_SetScriptSite, ObjPtr(uData.pSite))
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScript.SetScriptSite"
    End If
    hResult = DispCallByVtbl(ObjPtr(uData.pParse), IDX_InitNew)
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScriptParse.InitNew"
    End If
    '--- success
    ActiveScriptInit = True
    Exit Function
EH:
    Error = Err.Description & " in " & Err.Source
    PrintError FUNC_NAME
End Function

Public Sub ActiveScriptTerminate(uData As UcsActiveScriptData)
    Const FUNC_NAME     As String = "ActiveScriptTerminate"
    Dim hResult         As Long
    
    On Error GoTo EH
    If Not uData.pScript Is Nothing Then
        hResult = DispCallByVtbl(ObjPtr(uData.pScript), IDX_Close)
        If hResult < 0 Then
            Err.Raise hResult, "IActiveScript.Close"
        End If
    End If
    Set uData.pParse = Nothing
    Set uData.pScript = Nothing
    Set uData.cObjects = Nothing
    Set uData.pCallback = Nothing
    Set uData.pSite = Nothing
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Sub ActiveScriptReset(uData As UcsActiveScriptData)
    Dim hResult         As Long
    
    If uData.pScript Is Nothing Then
        Exit Sub
    End If
    Set uData.cObjects = New Collection
    hResult = DispCallByVtbl(ObjPtr(uData.pScript), IDX_Close)
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScript.Close"
    End If
    hResult = DispCallByVtbl(ObjPtr(uData.pScript), IDX_SetScriptSite, ObjPtr(uData.pSite))
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScript.SetScriptSite"
    End If
    hResult = DispCallByVtbl(ObjPtr(uData.pParse), IDX_InitNew)
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScriptParse.InitNew"
    End If
End Sub

Public Function ActiveScriptRunCode(uData As UcsActiveScriptData, sCode As String) As Variant
    Dim hResult         As Long
    Dim uException      As EXCEPINFO
    
    If uData.pParse Is Nothing Then
        Exit Function
    End If
    hResult = DispCallByVtbl(ObjPtr(uData.pParse), IDX_ParseScriptText, StrPtr(sCode), 0&, 0&, 0&, 0&, 0&, 0&, VarPtr(ActiveScriptRunCode), VarPtr(uException))
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScriptParse.ParseScriptText"
    End If
End Function

Public Function ActiveScriptCallFunction(uData As UcsActiveScriptData, ProcName As String, ParamArray Args() As Variant) As Variant
    Dim hResult         As Long
    Dim pDisp           As Object
    Dim vArgs           As Variant
    
    If uData.pScript Is Nothing Then
        Exit Function
    End If
    hResult = DispCallByVtbl(ObjPtr(uData.pScript), IDX_GetScriptDispatch, 0&, VarPtr(pDisp))
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScript.GetScriptDispatch"
    End If
    vArgs = Args
    ActiveScriptCallFunction = DispInvoke(pDisp, ProcName, VbMethod Or VbGet, vArgs)
End Function

Public Sub ActiveScriptAddObject(uData As UcsActiveScriptData, sName As String, oObj As Object)
    Const SCRIPTITEM_ISVISIBLE As Long = 2
    Const SCRIPTITEM_GLOBALMEMBERS As Long = 8
    Dim hResult         As Long
    
    uData.cObjects.Add oObj, sName
    hResult = DispCallByVtbl(ObjPtr(uData.pScript), IDX_AddNamedItem, StrPtr(sName), SCRIPTITEM_ISVISIBLE Or SCRIPTITEM_GLOBALMEMBERS)
    If hResult < 0 Then
        Err.Raise hResult, "IActiveScript.AddNamedItem"
    End If
End Sub

'= private ===============================================================

Private Function DispInvoke( _
            ByVal pDisp As Object, _
            ProcName As Variant, _
            ByVal CallType As VbCallType, _
            Optional Args As Variant) As Variant
    Const DISP_E_MEMBERNOTFOUND As Long = &H80020003
    Const DISP_E_PARAMNOTOPTIONAL As Long = &H8002000F
    Const DISPID_PROPERTYPUT    As Long = -3
    Const IDX_GetIDsOfNames     As Long = 5
    Const IDX_Invoke            As Long = 6
    Dim IID_NULL(0 To 3) As Long
    Dim lDispID         As Long
    Dim vRevArgs        As Variant
    Dim lIdx            As Long
    Dim aParams(0 To 3) As Long
    Dim lPropPutDispID  As Long
    Dim lResultPtr      As Long
    Dim hResult         As Long

    If pDisp Is Nothing Then
        hResult = DISP_E_PARAMNOTOPTIONAL
        GoTo QH
    End If
    '--- figure out procedure DispID
    If IsNumeric(ProcName) Then
        lDispID = ProcName
    Else
        hResult = DispCallByVtbl(ObjPtr(pDisp), IDX_GetIDsOfNames, VarPtr(IID_NULL(0)), VarPtr(StrPtr(ProcName)), 1&, 0&, VarPtr(lDispID))
        If hResult < 0 Then
            GoTo QH
        End If
    End If
    '--- reverse arguments
    If IsArray(Args) Then
        If UBound(Args) >= 0 Then
            ReDim vRevArgs(0 To UBound(Args) - LBound(Args)) As Variant
            For lIdx = 0 To UBound(vRevArgs)
                '--- have to keep VT_BYREF so cannot use simple assignment here
                Call VariantCopy(vRevArgs(lIdx), Args(UBound(Args) - lIdx))
            Next
            aParams(0) = VarPtr(vRevArgs(0))        ' .rgPointerToVariantArray
            aParams(2) = UBound(vRevArgs) + 1       ' .cArgs
        End If
    End If
    If (CallType And (VbLet Or VbSet)) <> 0 Then
        lPropPutDispID = DISPID_PROPERTYPUT
        aParams(1) = VarPtr(lPropPutDispID)     ' .rgPointerToLongNamedArgs
        aParams(3) = 1                          ' .cNamedArgs
    End If
    If (CallType And (VbGet Or VbMethod)) <> 0 Then
        lResultPtr = VarPtr(DispInvoke)
    End If
    hResult = DispCallByVtbl(ObjPtr(pDisp), IDX_Invoke, lDispID, VarPtr(IID_NULL(0)), 0&, CallType, VarPtr(aParams(0)), lResultPtr, 0&, 0&)
    '--- take care of subs (some do not accept result pointer)
    If hResult = DISP_E_MEMBERNOTFOUND And (CallType And VbMethod) <> 0 Then
        hResult = DispCallByVtbl(ObjPtr(pDisp), IDX_Invoke, lDispID, VarPtr(IID_NULL(0)), 0&, CallType, VarPtr(aParams(0)), 0&, 0&, 0&)
    End If
QH:
    If hResult < 0 Then
        IID_NULL(0) = vbError
        IID_NULL(2) = hResult
        Call VariantCopy(DispInvoke, IID_NULL(0))
    End If
End Function

Private Function DispCallByVtbl(ByVal pUnk As Long, ByVal lIndex As Long, ParamArray Args() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim vParams         As Variant
    Dim lIdx            As Long
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long

    vParams = Args
    For lIdx = 0 To UBound(vParams)
        vType(lIdx) = VarType(vParams(lIdx))
        vPtr(lIdx) = VarPtr(vParams(lIdx))
    Next
    hResult = DispCallFunc(pUnk, lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

'=========================================================================
' IActiveScriptSite interface
'=========================================================================

Private Function IActiveScriptSite_QueryInterface(This As UcsActiveScriptData, rIID As GUID, pvObj As Long) As Long
    pvObj = 0
    Select Case rIID.Data1
    Case IID_IUnknown.Data1
        If rIID.Data2 = IID_IUnknown.Data2 Then
            pvObj = VarPtr(This)
        End If
    Case IID_IActiveScriptSite.Data1
        If rIID.Data2 = IID_IActiveScriptSite.Data2 Then
            pvObj = VarPtr(This)
        End If
    Case IID_IActiveScriptSiteWindow.Data1
        If rIID.Data2 = IID_IActiveScriptSiteWindow.Data2 Then
            pvObj = VarPtr(This) + 4
        End If
    End Select
    If pvObj = 0 Then
        IActiveScriptSite_QueryInterface = E_NOINTERFACE
    End If
End Function

Private Function IActiveScriptSite_AddRef(This As UcsActiveScriptData) As Long
    '--- do nothing
End Function

Private Function IActiveScriptSite_Release(This As UcsActiveScriptData) As Long
    '--- do nothing
End Function

Private Function IActiveScriptSite_GetLCID(This As UcsActiveScriptData, plcid As Long) As Long
    IActiveScriptSite_GetLCID = E_NOTIMPL
End Function

Private Function IActiveScriptSite_GetItemInfo(This As UcsActiveScriptData, ByVal pstrName As String, ByVal dwReturnMask As Long, ppiunkItem As Long, ppti As Long) As Long
    Const SCRIPTINFO_IUNKNOWN As Long = 1
    Dim pUnk       As IUnknown
     
    On Error Resume Next
    Set pUnk = This.cObjects.Item(pstrName)
    If Not pUnk Is Nothing Then
        If dwReturnMask = SCRIPTINFO_IUNKNOWN Then
            ppiunkItem = ObjPtr(pUnk)
            Call CopyMemory(pUnk, 0&, 4)
        Else
            IActiveScriptSite_GetItemInfo = TYPE_E_ELEMENTNOTFOUND
        End If
    Else
        IActiveScriptSite_GetItemInfo = TYPE_E_ELEMENTNOTFOUND
    End If
End Function

Private Function IActiveScriptSite_GetDocVersionString(This As UcsActiveScriptData, ByVal pbstrVersion As Long) As Long
    IActiveScriptSite_GetDocVersionString = E_NOTIMPL
End Function

Private Function IActiveScriptSite_OnScriptTerminate(This As UcsActiveScriptData, pvarResult As Variant, ByVal pExcepinfo As Long) As Long
    IActiveScriptSite_OnScriptTerminate = E_NOTIMPL
End Function

Private Function IActiveScriptSite_OnStateChange(This As UcsActiveScriptData, ByVal ssScriptState As Long) As Long
    IActiveScriptSite_OnStateChange = E_NOTIMPL
End Function

Private Function IActiveScriptSite_OnScriptError(This As UcsActiveScriptData, ByVal pScriptError As IUnknown) As Long
    Dim uException      As EXCEPINFO
    Dim sSourceLine     As String
    Dim lCtx            As Long
    Dim lLine           As Long
    Dim lPos            As Long
    
    If Not This.pCallback Is Nothing Then
        Call DispCallByVtbl(ObjPtr(pScriptError), IDX_GetExceptionInfo, VarPtr(uException))
        Call DispCallByVtbl(ObjPtr(pScriptError), IDX_GetSourceLineText, VarPtr(sSourceLine))
        Call DispCallByVtbl(ObjPtr(pScriptError), IDX_GetSourcePosition, VarPtr(lCtx), VarPtr(lLine), VarPtr(lPos))
        Call This.pCallback.OnActiveScriptError(uException.Description, sSourceLine, lLine, lPos)
    End If
End Function

Private Function IActiveScriptSite_OnEnterScript(This As UcsActiveScriptData) As Long
    IActiveScriptSite_OnEnterScript = E_NOTIMPL
End Function

Private Function IActiveScriptSite_OnLeaveScript(This As UcsActiveScriptData) As Long
    IActiveScriptSite_OnLeaveScript = E_NOTIMPL
End Function

'=========================================================================
' IActiveScriptSiteWindow interface
'=========================================================================

Private Function IActiveScriptSiteWindow_QueryInterface(This As UcsActiveScriptSiteWindowData, rIID As GUID, pvObj As Long) As Long
    pvObj = 0
    Select Case rIID.Data1
    Case IID_IUnknown.Data1
        If rIID.Data2 = IID_IUnknown.Data2 Then
            pvObj = VarPtr(This) - 4
        End If
    Case IID_IActiveScriptSite.Data1
        If rIID.Data2 = IID_IActiveScriptSite.Data2 Then
            pvObj = VarPtr(This) - 4
        End If
    Case IID_IActiveScriptSiteWindow.Data1
        If rIID.Data2 = IID_IActiveScriptSiteWindow.Data2 Then
            pvObj = VarPtr(This)
        End If
    End Select
    If pvObj = 0 Then
        IActiveScriptSiteWindow_QueryInterface = E_NOINTERFACE
    End If
End Function

Private Function IActiveScriptSiteWindow_AddRef(This As UcsActiveScriptSiteWindowData) As Long
    '--- do nothing
End Function

Private Function IActiveScriptSiteWindow_Release(This As UcsActiveScriptSiteWindowData) As Long
    '--- do nothing
End Function

Private Function IActiveScriptSiteWindow_GetWindow(This As UcsActiveScriptSiteWindowData, phwnd As Long) As Long
    If Not This.pCallback Is Nothing Then
        phwnd = This.pCallback.OnActiveScriptGetWindow()
    End If
End Function

Private Function IActiveScriptSiteWindow_EnableModeless(This As UcsActiveScriptSiteWindowData, ByVal fEnable As Long) As Long
    IActiveScriptSiteWindow_EnableModeless = E_NOTIMPL
End Function
