Attribute VB_Name = "MStdPic"
Option Explicit
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As LongPtr) As LongPtr
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&

Private Declare Function AlphaBlend Lib "msimg32" (ByVal Dst_hDC As LongPtr, ByVal Dst_x As Long, ByVal Dst_y As Long, ByVal Dst_W As Long, ByVal Dst_H As Long, _
                                                   ByVal Src_hDC As LongPtr, ByVal Src_x As Long, ByVal Src_y As Long, ByVal Src_W As Long, ByVal Src_H As Long, ByVal Blendfunc As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr

'A Lightweight object wrapping a StdPicture-objekt and implementing an AlphaBlend-Render-function
Private Const IID_StdPicture   As String = "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
'GUID of IPicture:                         "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Private Const IID_IPicture     As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
'GUID of IPictureDisp:                     "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
Private Type VBGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data5(0 To 7) As Byte
End Type

'A VTable contains pointers to the functions of a class
Private Type TIPictureVTable
    '0 to 2 IUnknown
    '3 to 6 IDispatch
    '7 to 20 IPictureDisp
    Funcs(0 To 21) As LongPtr
End Type

Private m_IPictureVTable  As TIPictureVTable
Private m_pIPictureVTable As LongPtr

Private m_IPictureDispVTable  As TIPictureVTable
Private m_pIPictureDispVTable As LongPtr

Public Type TIPicture
    pVTable As LongPtr    ' First element in an object always is a pointer to it's VTable
    refCnt  As Long       ' the reference counter
    Picture As IPicture   ' the StdPicture-Variable holding all the bitmap-Data
    FncAlp  As Long       ' the AlphaBlend-function-type
    'CurhDC  As LongPtr    ' the current handle-device-context
    'hBmp    As LongPtr    ' the handle to the bitmap
End Type

Private Const S_OK    As Long = &H0&
Private Const S_FALSE As Long = &H1&

Public Sub InitIPictureVTable()
    m_pIPictureVTable = InitVTable(m_IPictureVTable)
    m_pIPictureDispVTable = InitVTable(m_IPictureDispVTable, True)
End Sub

Private Function InitVTable(vtb As TIPictureVTable, Optional ByVal bAddDispatch As Boolean = False) As LongPtr
    Dim i As Long
    With vtb
        
        'IUnkown
        .Funcs(i) = FncPtr(AddressOf IUnknown_FncQueryInterface):      i = i + 1
        .Funcs(i) = FncPtr(AddressOf IUnknown_SubAddRef):              i = i + 1
        .Funcs(i) = FncPtr(AddressOf IUnknown_SubRelease):             i = i + 1
        
        'IDispatch
        If bAddDispatch Then
            .Funcs(i) = FncPtr(AddressOf IDispatch_get_TypeInfoCount): i = i + 1
            .Funcs(i) = FncPtr(AddressOf IDispatch_get_TypeInfo):      i = i + 1
            .Funcs(i) = FncPtr(AddressOf IDispatch_get_IDsOfNames):    i = i + 1
            .Funcs(i) = FncPtr(AddressOf IDispatch_FncInvoke):         i = i + 1
        End If
        
        'IPicture
'        HRESULT ( STDMETHODCALLTYPE *get_Handle )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_Handle):             i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *get_hPal )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_hPal):               i = i + 1

'        HRESULT ( STDMETHODCALLTYPE *get_Type )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_Type):               i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *get_Width )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_Width):              i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *get_Height )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_Height):             i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *Render )(
        .Funcs(i) = FncPtr(AddressOf IPicture_SubRender):              i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *set_hPal )(
        .Funcs(i) = FncPtr(AddressOf IPicture_set_hPal):               i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *get_CurDC )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_CurDC):              i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *SelectPicture )(
        .Funcs(i) = FncPtr(AddressOf IPicture_SelectPicture):          i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *get_KeepOriginalFormat )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_KeepOriginalFormat): i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *put_KeepOriginalFormat )(
        .Funcs(i) = FncPtr(AddressOf IPicture_put_KeepOriginalFormat): i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *PictureChanged )(
        .Funcs(i) = FncPtr(AddressOf IPicture_PictureChanged):         i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *SaveAsFile )(
        .Funcs(i) = FncPtr(AddressOf IPicture_SaveAsFile):             i = i + 1
        
'        HRESULT ( STDMETHODCALLTYPE *get_Attributes )(
        .Funcs(i) = FncPtr(AddressOf IPicture_get_Attributes):         i = i + 1
        
        .Funcs(i) = FncPtr(AddressOf IPicture_SetHdc):                 i = i + 1
    End With
    InitVTable = VarPtr(vtb)
    Dim hr As Long: hr = VirtualProtect(InitVTable, i * SizeOf_LongPtr, PAGE_EXECUTE_READWRITE, 0&)
End Function
Private Function FncPtr(ByVal pFnc As Long) As Long
    FncPtr = pFnc
End Function

Public Function New_IPicture(this As TIPicture, aPicture As IPicture, Optional ByVal AlphaFunc As Long = &H1FF0000) As IPicture
    If m_pIPictureVTable = 0 Then m_pIPictureVTable = InitVTable(m_IPictureVTable)
    With this
        .pVTable = m_pIPictureVTable
        Set .Picture = aPicture
        .FncAlp = AlphaFunc
        .refCnt = 2
        '.CurhDC = CreateCompatibleDC(0&)
        '.hBmp = SelectObject(.CurhDC, aPicture.Handle)
    End With
    'bring the object to life
    RtlMoveMemory New_IPicture, VarPtr(this), SizeOf_LongPtr
End Function

Public Function New_IPictureDisp(this As TIPicture, aPicture As IPicture, Optional ByVal AlphaFunc As Long = &H1FF0000) As IPictureDisp
    If m_pIPictureDispVTable = 0 Then m_pIPictureDispVTable = InitVTable(m_IPictureDispVTable, True)
    With this
        .pVTable = m_pIPictureDispVTable
        Set .Picture = aPicture
        .FncAlp = AlphaFunc
        .refCnt = 2
    End With
    'bring the object to life
    RtlMoveMemory New_IPictureDisp, VarPtr(this), SizeOf_LongPtr
End Function

' v ############################## v '    IUnkown    ' v ############################## v '
Private Function IUnknown_FncQueryInterface(this As TIPicture, riid As VBGuid, pvObj As Long) As Long
    '7BF80981-BF32-101A-8BBB-00AA00300CAB
    With riid
        'do you want IPicture or IPictureDisp?
        If .Data2 = &HBF32 And .Data3 = &H101A And _
           .Data5(0) = &H8B And .Data5(1) = &HBB And .Data5(2) = &H0 And .Data5(3) = &HAA And _
           .Data5(4) = &H0 And .Data5(5) = &H30 And .Data5(6) = &HC And .Data5(7) = &HAB Then
           Const d1 As Long = &H7BF80980
            If .Data1 = d1 Then
                'OK you want IPicture
                Debug.Print "IUnknown_FncQueryInterface IPicture"
                pvObj = VarPtr(this) '<-- important!
                IUnknown_FncQueryInterface = S_OK ' yes we have this Interface
                Exit Function
            End If
            If .Data1 = d1 + 1 Then
                'OK you want IPictureDisp
                Debug.Print "IUnknown_FncQueryInterface IPictureDisp"
                pvObj = VarPtr(this) '<-- important!
                IUnknown_FncQueryInterface = S_OK ' yes we have this Interface
                Exit Function
            End If
        End If
    End With
End Function

Private Function IUnknown_SubAddRef(this As TIPicture) As Long
    ' now we add one reference
    With this
        .refCnt = .refCnt + 1
    End With
End Function

Private Function IUnknown_SubRelease(this As TIPicture) As Long
    ' now we subtract one reference
    With this
        .refCnt = .refCnt - 1
    End With
    'If this.refCnt = 0 Then 'cleanup
End Function
' ^ ############################## ^ '    IUnkown    ' ^ ############################## ^ '

' v ############################## v '   IDispatch   ' v ############################## v '
Private Function IDispatch_get_TypeInfoCount(this As TIPicture) As Long
    '
End Function
Private Function IDispatch_get_TypeInfo(this As TIPicture) As Long
    '
End Function
Private Function IDispatch_get_IDsOfNames(this As TIPicture) As Long
    '
End Function
Private Function IDispatch_FncInvoke(this As TIPicture) As Long
    '
End Function
' ^ ############################## ^ '   IDispatch   ' ^ ############################## ^ '

' v ############################## v '    IPicture   ' v ############################## v '
'HRESULT ( STDMETHODCALLTYPE *get_Handle )( __RPC__in IPicture * This, /* [out] */ __RPC__out OLE_HANDLE *pHandle);
Private Function IPicture_get_Handle(this As TIPicture, pHandle_out As LongPtr) As Long
    pHandle_out = this.Picture.Handle
End Function

'HRESULT ( STDMETHODCALLTYPE *get_hPal )( __RPC__in IPicture * This, /* [out] */ __RPC__out OLE_HANDLE *phPal);
Private Function IPicture_get_hPal(this As TIPicture, phPal_out As LongPtr) As Long
    phPal_out = this.Picture.hPal
End Function


'HRESULT ( STDMETHODCALLTYPE *get_Type )( __RPC__in IPicture * This, /* [out] */ __RPC__out SHORT *pType);
Private Function IPicture_get_Type(this As TIPicture, pType_out As Integer) As Long
    pType_out = this.Picture.Type
End Function

'HRESULT ( STDMETHODCALLTYPE *get_Width )( __RPC__in IPicture * This, /* [out] */ __RPC__out OLE_XSIZE_HIMETRIC *pWidth);
Private Function IPicture_get_Width(this As TIPicture, pWidth_out As Long) As Long
    pWidth_out = this.Picture.Width
End Function

'HRESULT ( STDMETHODCALLTYPE *get_Height )( __RPC__in IPicture * This, /* [out] */ __RPC__out OLE_YSIZE_HIMETRIC *pHeight);
Private Function IPicture_get_Height(this As TIPicture, pHeight_out As Long) As Long
    pHeight_out = this.Picture.Height
End Function

'HRESULT ( STDMETHODCALLTYPE *Render )( __RPC__in IPicture * This,
'            /* [in] */ __RPC__in HDC hDC,
'            /* [in] */ LONG x,
'            /* [in] */ LONG y,
'            /* [in] */ LONG cx,
'            /* [in] */ LONG cy,
'            /* [in] */ OLE_XPOS_HIMETRIC xSrc,
'            /* [in] */ OLE_YPOS_HIMETRIC ySrc,
'            /* [in] */ OLE_XSIZE_HIMETRIC cxSrc,
'            /* [in] */ OLE_YSIZE_HIMETRIC cySrc,
'            /* [in] */ __RPC__in LPCRECT pRcWBounds);
Private Function IPicture_SubRender(this As TIPicture, ByVal Dst_hDC As LongPtr, _
                                    ByVal Dst_x As Long, ByVal Dst_y As Long, ByVal Dst_cx As Long, ByVal Dst_cy As Long, _
                                    ByVal Src_x As Long, ByVal Src_y As Long, ByVal Src_cx As Long, ByVal Src_cy As Long) As Long
    'This.CurhDC = Dst_hDC '???
    
    Dim Src_hDC As LongPtr: Src_hDC = CreateCompatibleDC(0&)
    Dim hBmp    As LongPtr:    hBmp = SelectObject(Src_hDC, this.StdPic.Handle)
    
    AlphaBlend Dst_hDC, Dst_x, Dst_y, Dst_cx, Dst_cy, Src_hDC, 0, 0, Src_cx, Src_cy, this.FncAlp  ' &H1FF0000
    
    SelectObject Src_hDC, hBmp
    DeleteDC Src_hDC
End Function
'the function of LaVolpe
'Private Function NewRender(ByVal This As Long, ByVal hdc As Long, _
'                           ByVal X As Long, ByVal Y As Long, _
'                           ByVal cx As Long, ByVal cy As Long, _
'                           ByVal xSrc As Long, ByVal ySrc As Long, _
'                           ByVal cxSrc As Long, ByVal cySrc As Long, _
'                           ByVal pRcBounds As Long) As Long
'
'    Debug.Print This, hdc; X; Y; cx; cy; xSrc; ySrc; cxSrc; cySrc
'    Dim IPic As IPicture, tObj As Object, srcDC As Long, hBmp As Long
'
'    CopyMemory tObj, This, 4&
'    Set IPic = tObj
'    CopyMemory tObj, 0&, 4&
'
'    srcDC = CreateCompatibleDC(0&)
'    hBmp = SelectObject(srcDC, IPic.Handle)
'
'    AlphaBlend hdc, X, Y, cx, cy, srcDC, 0, 0, cx, cy, &H1FF0000
'
'    SelectObject srcDC, hBmp
'    DeleteDC srcDC
'
'End Function

'
'        HRESULT ( STDMETHODCALLTYPE *set_hPal )(
'            __RPC__in IPicture * This,
'            /* [in] */ OLE_HANDLE hPal);
Private Function IPicture_set_hPal(this As TIPicture, ByVal hPal_in As LongPtr)
    Debug.Print "IPicture_set_hPal " & hPal_in
    this.Picture.hPal = hPal_in
End Function

'        HRESULT ( STDMETHODCALLTYPE *get_CurDC )(
'            __RPC__in IPicture * This,
'            /* [out] */ __RPC__deref_out_opt HDC *phDC);
Private Function IPicture_get_CurDC(this As TIPicture, phDC_out As LongPtr)
    Debug.Print "IPicture_get_CurDC " & phDC_out
    phDC_out = this.Picture.CurDC ' CurhDC
End Function

'        HRESULT ( STDMETHODCALLTYPE *SelectPicture )(
'            __RPC__in IPicture * This,
'            /* [in] */ __RPC__in HDC hDCIn,
'            /* [out] */ __RPC__deref_out_opt HDC *phDCOut,
'            /* [out] */ __RPC__out OLE_HANDLE *phBmpOut);
Private Function IPicture_SelectPicture(this As TIPicture, ByVal hDC_in As LongPtr, phDC_out As LongPtr, phBmp_out As LongPtr)
    Debug.Print "IPicture_SelectPicture " & hDC_in & " " & phDC_out & " " & phBmp_out
    'phDC_out = this.Picture.CurDC 'CurhDC
    'phBmp_out = this.Picture.SelectPicture 'hBmp ' this.StdPic.Handle
    this.Picture.SelectPicture hDC_in, phDC_out, phBmp_out
End Function

'        HRESULT ( STDMETHODCALLTYPE *get_KeepOriginalFormat )(
'            __RPC__in IPicture * This,
'            /* [out] */ __RPC__out BOOL *pKeep);
Private Function IPicture_get_KeepOriginalFormat(this As TIPicture, pKeep_out As Long)
    Debug.Print "IPicture_get_KeepOriginalFormat " & pKeep_out
    pKeep_out = this.Picture.KeepOriginalFormat
    'phDC_out = this.CurhDC
    'phBmp_out = this.StdPic.Handle
    'pKeep_out = ???
End Function

'        HRESULT ( STDMETHODCALLTYPE *put_KeepOriginalFormat )(
'            __RPC__in IPicture * This,
'            /* [in] */ BOOL keep);
Private Function IPicture_put_KeepOriginalFormat(this As TIPicture, ByVal bKeep As Boolean)
    Debug.Print "IPicture_put_KeepOriginalFormat " & bKeep
    this.Picture.KeepOriginalFormat = bKeep
    'todo
End Function

'        HRESULT ( STDMETHODCALLTYPE *PictureChanged )(
'            __RPC__in IPicture * This);
Private Function IPicture_PictureChanged(this As TIPicture)
    Debug.Print "IPicture_PictureChanged"
    this.Picture.PictureChanged
End Function

'        HRESULT ( STDMETHODCALLTYPE *SaveAsFile )(
'            __RPC__in IPicture * This,
'            /* [in] */ __RPC__in_opt LPSTREAM pStream,
'            /* [in] */ BOOL fSaveMemCopy,
'            /* [out] */ __RPC__out LONG *pCbSize);
Private Function IPicture_SaveAsFile(this As TIPicture, ByVal pStream_in As LongPtr, ByVal fSaveMemCopy As Boolean, pCbSize_out As Long)
    Debug.Print "IPicture_SaveAsFile"
    this.Picture.SaveAsFile ByVal pStream_in, fSaveMemCopy, pCbSize_out
End Function

'        HRESULT ( STDMETHODCALLTYPE *get_Attributes )(
'            __RPC__in IPicture * This,
'            /* [out] */ __RPC__out DWORD *pDwAttr);
Private Function IPicture_get_Attributes(this As TIPicture, pDwAttr_out As Long)
    'todo
    pDwAttr_out = this.Picture.Attributes
End Function

Private Function IPicture_SetHdc(this As TIPicture, ByVal Value As Long)
    this.Picture.SetHdc Value
End Function

' ^ ############################## ^ '    IPicture   ' ^ ############################## ^ '

