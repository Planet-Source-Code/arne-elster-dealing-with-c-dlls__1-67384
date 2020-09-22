Attribute VB_Name = "modCallCDECL"
Option Explicit

Private Declare Function LoadLibraryA Lib "kernel32" ( _
    ByVal strLib As String _
) As Long

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hLib As Long, ByVal strProc As String _
) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" ( _
    ByVal strMod As String _
) As Long

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long, _
    ByVal flAllocationType As Long, ByVal flProtect As Long _
) As Long

Private Declare Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Private Declare Function VirtualLock Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long _
) As Long

Private Declare Function VirtualUnlock Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long _
) As Long

Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal pFnc As Long, ByVal arg1 As Long, _
    ByVal arg2 As Long, ByVal arg3 As Long, _
    ByVal arg4 As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cBytes As Long _
)

Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_DECOMMIT              As Long = &H4000

Public Function GetProcAddressEx(ByVal strLib As String, ByVal strFnc As String) As Long
    Dim hMod    As Long
    
    hMod = GetModuleHandleA(strLib)
    If hMod = 0 Then hMod = LoadLibraryA(strLib)
    If hMod = 0 Then Exit Function
    
    GetProcAddressEx = GetProcAddress(hMod, strFnc)
End Function

Public Function CreateCdeclCbWrap(ByVal pFnc As Long, ByVal ArgCount As Long) As Long
    Dim i       As Long
    Dim pAsm    As Long
    Dim pInstr  As Long
    
    pAsm = VirtualAlloc(0, 256, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock pAsm, 256
    
    pInstr = pAsm
    
    ' stdcall callee removes all parameters from the stack,
    ' but cdecl caller does this, too.
    ' all arguments have to be duplicated.
    For i = 1 To ArgCount
        AddByte pInstr, &HFF        ' PUSH [ESP+ARGCOUNT*4]
        AddByte pInstr, &H74
        AddByte pInstr, &H24
        AddByte pInstr, ArgCount * 4
    Next
    
    AddCall pInstr, pFnc            ' CALL pFnc
    AddByte pInstr, &HC3            ' RET
    AddByte pInstr, 0
    
    CreateCdeclCbWrap = pAsm
End Function

Public Sub DestroyDeclCbWrap(ByVal hCb As Long)
    VirtualUnlock hCb, 256
    VirtualFree hCb, 256, MEM_DECOMMIT
End Sub

Public Function CallCdecl(ByVal pFnc As Long, ParamArray args()) As Long
    Dim i       As Long
    Dim pAsm    As Long
    Dim pInstr  As Long
    
    pAsm = VirtualAlloc(0, 256, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock pAsm, 256
    
    pInstr = pAsm
    
    ' remove the CallWindowProc arguments from the stack
    AddByte pInstr, &H58                    ' POP EAX
    AddByte pInstr, &H59                    ' POP ECX
    AddByte pInstr, &H59                    ' POP ECX
    AddByte pInstr, &H59                    ' POP ECX
    AddByte pInstr, &H59                    ' POP ECX
    AddByte pInstr, &H50                    ' PUSH EAX
    
    For i = UBound(args) To 0 Step -1
        AddPush pInstr, CLng(args(i))       ' PUSH arg(i)
    Next
    
    AddCall pInstr, pFnc                    ' CALL pFnc
    AddByte pInstr, &H83                    ' ADD ESP, ArgCount*4
    AddByte pInstr, &HC4
    AddByte pInstr, 4 * (UBound(args) + 1)
    AddByte pInstr, &HC3                    ' RET
    AddByte pInstr, &H0
    
    CallCdecl = CallWindowProcA(pAsm, 0, 0, 0, 0)
    
    VirtualUnlock pAsm, 256
    VirtualFree pAsm, 256, MEM_DECOMMIT
End Function

Private Sub AddPush(pAsm As Long, lng As Long)
    AddByte pAsm, &H68
    AddLong pAsm, lng
End Sub

Private Sub AddCall(pAsm As Long, addr As Long)
    AddByte pAsm, &HE8
    AddLong pAsm, addr - pAsm - 4
End Sub

Private Sub AddLong(pAsm As Long, lng As Long)
    CpyMem ByVal pAsm, lng, 4
    pAsm = pAsm + 4
End Sub

Private Sub AddByte(pAsm As Long, Bt As Byte)
    CpyMem ByVal pAsm, Bt, 1
    pAsm = pAsm + 1
End Sub
