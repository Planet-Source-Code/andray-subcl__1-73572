;cSubclass.cls model assembler source

;Runtime patch markers
%define _patch1_    01BCCAABh ;Relative address of the EbMode function
%define _patch2_    02BCCAABh ;Rel. addr. of BOOL RemoveWindowSubclass(in HWND hWnd, in SUBCLASSPROC pfnSubclass, in UINT_PTR uIdSubclass)
%define _patch3_    03BCCAABh ;addr. of subclass proc (this module)
%define _patch4_    04BCCAABh ;Relative address of DefSubclassProc(in HWND hWnd, in UINT uMsg, in WPARAM WPARAM, in LPARAM LPARAM)
%define _patch5_    05BCCAABh ;Address of the owner object

;Stack frame parameters and local variables. After "enter 8, 0" the ebp stack frame will look like this...
%define dwRefData   [ebp+28]
%define uIdSubclass [ebp+24]
%define lParam      [ebp+20]            ;lParam parameter
%define wParam      [ebp+16]            ;wParam parameter
%define uMsg        [ebp+12]            ;Message number parameter
%define hWnd        [ebp +8]            ;Window handle parameter
       ;Information [ebp +4] return address to the caller
       ;Information [ebp +0] previous value of ebp pushed here, (implicitly restored with the leave statement)
%define lReturn     [ebp -4]            ;Local variable lReturn
%define bHandled    [ebp -8]            ;Local variable bHandled
                   
%define GWL_WNDPROC -4                  ;SetWindowsLong WndProc offset

[bits 32]   ;32bit code

;Entry point, setup stack frame
	enter 8, 0 ;WA/ push ebp / mov ebp, esp / sub esp, 8
    
;Initialize locals
    xor     eax, eax                    ;Clear eax
    mov     lReturn, eax                ;Clear lReturn
    mov     bHandled, eax               ;Clear bHandled

    jmp     _no_ide                     ;Patched with two nop's if running in the the IDE

;Check to see if the IDE is on a breakpoint or has stopped
    db 0E8h 	;Far call
    dd _patch1_ ;EbMode
    cmp     eax, 2                      ;If 2 
    je      _break                      ;  IDE is on a breakpoint, just call the original WndProc
    test    eax, eax                    ;If 0
    je      _unsub                      ;  IDE has stopped, unsubclass the window
    
_no_ide:
	call 	_callback ;WA/ user proc
    cmp     bHandled, dword 0           ;Has message been handled?
    jne     _return                     ;  Yep, return
_break:     ;The IDE is on a breakpoint, call the original WndProc and return
    call    _original                   ;Call the original 

_return:    ;Cleanup and exit
    mov     eax, lReturn                ;Function return value
    leave                               ;Restore Stack for Procedure Exit /WA/(MOV ESP, EBP; POP EBP)
    ret     24                          ;Return and adjust esp (6 variables * 4)

_unsub:     ;IDE has stopped, unsubclass the window
    push    dword uIdSubclass
    push    dword _patch3_
    push    dword hWnd                  ;Push the window handle
    db 0E8h 	;Far call
    dd _patch2_ ;RemoveWindowSubclass
    jmp     _return                     ;Return
    
_original:  ;Call original WndProc
    push    dword lParam                ;ByVal lParam
    push    dword wParam                ;ByVal wParam
    push    dword uMsg                  ;ByVal uMsg
    push    dword hWnd                  ;ByVal hWnd
    db 0E8h 	;Far call
    dd _patch4_ ;DefSubclassProc
	mov     lReturn, eax                ;Preserve the return value
    ret
    
_callback:
	push dword dwRefData
	push dword uIdSubclass
	push dword lParam
	push dword wParam
	push dword uMsg
	push dword hWnd
    lea     eax, lReturn                ;Address of lReturn into eax
    push    eax                         ;Push ByRef lReturn
    lea     eax, bHandled               ;Address of bHandled into eax
    push    eax                         ;Push ByRef bHandled

    mov     eax, _patch5_               ;Address of the owner object, patched at runtime
    push    eax                         ;Push address of the owner object
    mov     eax, [eax]                  ;Get the address of the vTable into eax
    call    dword [eax+1Ch]             ;Call subclass proc, vTable offset 1Ch (1st proc in csubcl.cls)
    ret
    