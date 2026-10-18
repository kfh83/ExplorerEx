option casemap:none
EXTERN ResolveDxgiExport:PROC

.code
; Preserve integer, floating/vector arguments, and untouched stack arguments.
; Resolve once in C++, then tail-call the real system export with its exact ABI.
DXGI_FORWARD MACRO name, index
PUBLIC name
name PROC FRAME
    sub rsp, 0A8h
    .allocstack 0A8h
    .endprolog
    mov [rsp+20h], rcx
    mov [rsp+28h], rdx
    mov [rsp+30h], r8
    mov [rsp+38h], r9
    movdqu [rsp+40h], xmm0
    movdqu [rsp+50h], xmm1
    movdqu [rsp+60h], xmm2
    movdqu [rsp+70h], xmm3
    movdqu [rsp+80h], xmm4
    movdqu [rsp+90h], xmm5
    mov ecx, index
    call ResolveDxgiExport
    mov r11, rax
    mov rcx, [rsp+20h]
    mov rdx, [rsp+28h]
    mov r8, [rsp+30h]
    mov r9, [rsp+38h]
    movdqu xmm0, [rsp+40h]
    movdqu xmm1, [rsp+50h]
    movdqu xmm2, [rsp+60h]
    movdqu xmm3, [rsp+70h]
    movdqu xmm4, [rsp+80h]
    movdqu xmm5, [rsp+90h]
    add rsp, 0A8h
    jmp r11
name ENDP
ENDM

DXGI_FORWARD ProxyApplyCompatResolutionQuirking, 0
DXGI_FORWARD ProxyCompatString, 1
DXGI_FORWARD ProxyCompatValue, 2
DXGI_FORWARD ProxyCreateDXGIFactory, 3
DXGI_FORWARD ProxyCreateDXGIFactory1, 4
DXGI_FORWARD ProxyCreateDXGIFactory2, 5
DXGI_FORWARD ProxyDXGID3D10CreateDevice, 6
DXGI_FORWARD ProxyDXGID3D10CreateLayeredDevice, 7
DXGI_FORWARD ProxyDXGID3D10GetLayeredDeviceSize, 8
DXGI_FORWARD ProxyDXGID3D10RegisterLayers, 9
DXGI_FORWARD ProxyDXGIDumpJournal, 11
DXGI_FORWARD ProxyDXGIGetDebugInterface1, 12
DXGI_FORWARD ProxyDXGIReportAdapterConfiguration, 13
DXGI_FORWARD ProxyPIXBeginCapture, 14
DXGI_FORWARD ProxyPIXEndCapture, 15
DXGI_FORWARD ProxyPIXGetCaptureState, 16
DXGI_FORWARD ProxySetAppCompatStringPointer, 17
DXGI_FORWARD ProxyUpdateHMDEmulationStatus, 18
END
