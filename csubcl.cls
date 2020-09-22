VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "csubcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'SubCL class - one-file, IDE safe, machine code subclasser
'-------------------------------------------------------------------
'This code is based on Paul Caton's (Paul_Caton@hotmail.com)
'cSubclass and Karl E. Peterson's (http://vb.mvps.org) HookXP
'
'Winand, 2010 (http://sourceforge.net/projects/audica/)


Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function GetProcAddressByOrd Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpOrdinal As Long) As Long

Private Const WM_NCDESTROY As Long = &H82
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function GetWindowSubclass Lib "comctl32" Alias "#411" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, pdwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private hlib_comctl32 As Long, RemoveWindowSubclass_ As Long, DefSubclassProc_ As Long ', SetWindowSubclass_ As Long, GetWindowSubclass_ As Long
Private nAddrSubclass As Long

Const PATCH_01 As Long = 15
Const PATCH_03 As Long = 55
Const PATCH_02 As Long = 63
Const PATCH_04 As Long = 82
Const PATCH_05 As Long = 117
Const CODE_STR As String = "C808000031C08945FC8945F8EB0EE8xxxxx01x83F802740F85C07417E839000000837DF8007505E8190000008B45FCC9C21800FF751868xxxxx03xFF7508E8xxxxx02xEBE7FF7514FF7510FF750CFF7508E8xxxxx04x8945FCC3FF751CFF7518FF7514FF7510FF750CFF75088D45FC508D45F850B8xxxxx05x508B00FF501CC3"

'THIS PROCEDURE SHOULD BE FIRST IN CLASS - don't move it anywhere
Public Sub proc(ByRef bHandled As Boolean, ByRef lReturn As Long, _
                    ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
                    ByVal lParam As Long, ByVal uIdSubclass As Object, ByVal dwRefData As Long)
    ' http://msdn.microsoft.com/en-us/library/bb776774(VS.85).aspx
    lReturn = uIdSubclass.WindowProc(bHandled, hWnd, uMsg, wParam, lParam, dwRefData)
    ' This should *never* be necessary, but just in case client fails to...
    If uMsg = WM_NCDESTROY Then
        Debug.Print "Window " & hWnd & " is being destroyed. Removing hook automatically..."
        Call HookClear(uIdSubclass, hWnd)
    End If
End Sub

Private Sub Class_Initialize()
    Dim i As Long
    Dim nLen As Long 'String lengths
    Dim sCode As String 'Binary code string
    Dim EbMode As Long 'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    
    nLen = Len(CODE_STR) 'Length of hex pair string
    'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
    For i = 1 To nLen Step 2
      sCode = sCode & ChrB$(Val("&H" & Mid$(CODE_STR, i, 2)))
    Next i
    nLen = LenB(sCode) 'Get the machine code length
    nAddrSubclass = GlobalAlloc(0, nLen) 'Allocate fixed memory for machine code buffer
    Call CopyMemory(ByVal nAddrSubclass, ByVal StrPtr(sCode), nLen) 'Copy the code to allocated memory
    Call PatchVal(PATCH_03, nAddrSubclass) 'addr. of subclass proc
    Debug.Print "Subclassing procedure address is " & nAddrSubclass
    
    hlib_comctl32 = LoadLibrary("comctl32")
    If hlib_comctl32 Then
'        SetWindowSubclass_ = GetProcAddress(hlib_comctl32, "SetWindowSubclass")
'        GetWindowSubclass_ = GetProcAddressByOrd(hlib_comctl32, 411)
        If b(RemoveWindowSubclass_, GetProcAddress(hlib_comctl32, "RemoveWindowSubclass")) Then _
            Call PatchRel(PATCH_02, RemoveWindowSubclass_) _
        Else Err.Raise 1, , "Cannot get RemoveWindowSubclass procedure address"
        If b(DefSubclassProc_, GetProcAddress(hlib_comctl32, "DefSubclassProc")) Then _
            Call PatchRel(PATCH_04, DefSubclassProc_) _
        Else Err.Raise 1, , "Cannot get DefSubclassProc procedure address"
    Else
        Err.Raise 1, , "Cannot load comctl32 library"
    End If
    
    If inIDE Then
        'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
        Call CopyMemory(ByVal nAddrSubclass + 12, &H9090, 2)
        'Patch the relative address to the EbMode api function
        If b(EbMode, GetProcAddress(GetModuleHandle("vba6"), "EbMode")) Then _
            Call PatchRel(PATCH_01, EbMode) _
        Else Err.Raise 1, , "Cannot get EbMode procedure address" 'Get the address of EbMode in vba6.dll
    End If
    Call PatchVal(PATCH_05, ObjPtr(Me)) 'Address of the owner object
End Sub

Private Sub Class_Terminate()
    Call GlobalFree(nAddrSubclass) 'Release the allocated memory
    If hlib_comctl32 Then Call FreeLibrary(hlib_comctl32)
End Sub

Private Function b(p1 As Long, p2 As Long) As Boolean
    p1 = p2
    b = CBool(p1)
End Function

Private Function inIDE() As Boolean
On Error GoTo 1:
    Debug.Print 1 / 0
Exit Function
1:  inIDE = True
End Function

'Patch the machine code buffer offset with the relative address to the target address
Private Sub PatchRel(ByVal nOffset As Long, _
                     ByVal nTargetAddr As Long)
  Call CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
End Sub

'Patch the machine code buffer offset with the passed value
Private Sub PatchVal(ByVal nOffset As Long, _
                     ByVal nValue As Long)
  Call CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
End Sub

Public Function HookSet(ByVal Thing As Object, Optional ByVal hWnd As Long = 0, Optional dwRefData As Long) As Boolean
    ' http://msdn.microsoft.com/en-us/library/bb762102(VS.85).aspx
    If hWnd Then _
        HookSet = CBool(SetWindowSubclass(hWnd, nAddrSubclass, ObjPtr(Thing), dwRefData)) _
    Else HookSet = CBool(SetWindowSubclass(Thing.hWnd, nAddrSubclass, ObjPtr(Thing), dwRefData))
    Debug.Print "HookSet returns " & HookSet & ". Subclass objest is " & ObjPtr(Thing) & ", subclassed window is " & hWnd & ", user data is " & dwRefData
End Function

Public Function HookGetData(ByVal Thing As Object, Optional ByVal hWnd As Long, Optional ByRef dwRefData As Long) As Long
   ' http://msdn.microsoft.com/en-us/library/bb776430(VS.85).aspx
    If hWnd Then _
        HookGetData = GetWindowSubclass(hWnd, nAddrSubclass, ObjPtr(Thing), dwRefData) _
    Else HookGetData = GetWindowSubclass(Thing.hWnd, nAddrSubclass, ObjPtr(Thing), dwRefData)
End Function

Public Function HookClear(ByVal Thing As Object, Optional ByVal hWnd As Long = 0) As Boolean
    ' http://msdn.microsoft.com/en-us/library/bb762094(VS.85).aspx
    If hWnd Then _
        HookClear = CBool(RemoveWindowSubclass(hWnd, nAddrSubclass, ObjPtr(Thing))) _
    Else HookClear = CBool(RemoveWindowSubclass(Thing.hWnd, nAddrSubclass, ObjPtr(Thing)))
    Debug.Print "HookClear returns " & HookClear & ". Subclass objest is " & ObjPtr(Thing) & ", subclassed window is " & hWnd
End Function

Public Function HookDefault(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' http://msdn.microsoft.com/en-us/library/bb776403(VS.85).aspx
   HookDefault = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

