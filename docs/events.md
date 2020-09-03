# OneNote events &mdash; very buggy!

This document applies to OneNote 16.0.13127.20266 64-bit for Windows.

## Problem 1: `IConnectionPoint.Advise` fails with `E_FAIL`

The symptom is that the [`events.cc`](../examples/events.cc) will get `E_FAIL` upon calling `IConnectionPoint.Advise` method of the connection point found by `FindConnectionPoint` on the `OneNote.Application` object.

By breaking (the debugger!) into OneNote, I found that `CStdIdentity::CInternalUnk::QueryInterface` to `IID_IOneNoteEvents` fails with `E_FAIL`, and on my side, the call to `QueryInterface` returned `S_OK`. By recalling [Raymond's](https://devblogs.microsoft.com/oldnewthing/20040220-00/?p=40533) [blog](https://devblogs.microsoft.com/oldnewthing/20041213-00/?p=37043), I turned to the registry. But the interfaces are registered.

What about VBA? OneNote does not have built-in VBA, so I used another Office app (Excel in this case) and cross-referenced OneNote. With the `WithEvents` construct, I still got no luck. I verified that my constructs were correct by working with PowerPoint instead (so VBA and C++ event sinks for `PowerPoint.Application`). Everything worked perfectly for PowerPoint, so it must be a bug of OneNote!

Wait a second, I found that PowerPoint didn't even bother querying `EApplication` interface (it should have done so to avoid unintended invocations via `IDispatch`). It used `IDispatch.Invoke` to invoke the events. In contrast, OneNote really is querying the `IOneNoteEvents` interface.

By comparing the type libraries of OneNote and PowerPoint, I figured that OneNote type library declared `IOneNoteEvents` as `dispinterface` (that's also why one version of this repository defined `DIID_IOneNoteEvents`), while PowerPoint declared `EApplication` as dual `interface`. **This resolves the mystery.** The standard marshaler needs type information to marshal interfaces, it finds the interface registered to the OneNote type library, so it happily tries looking for its definition, except there isn't such a definition since it is only declared as `dispinterface`! The unmarshaling fails with `E_FAIL`.

Knowing the reason, we can solve this problem by editing the IDL, generating a new TLB, and set the registry to use the new TLB.

```diff
 [
+  odl,
   uuid(E2E1511D-502D-4BD0-8B3A-8A89A05CDCAE),
   helpstring("IOneNoteEvents Interface"),
+  dual,
+  oleautomation,
   nonextensible
 ]
-dispinterface IOneNoteEvents {
+interface IOneNoteEvents : IDispatch
+{
-properties:
-methods:
-  [id(0x00000001)]
-  void OnNavigate();
-  [id(0x00000002)]
-  void OnHierarchyChange([in] BSTR bstrActivePageID);
+  [id(0x00000001)]
+  HRESULT OnNavigate();
+  [id(0x00000002)]
+  HRESULT OnHierarchyChange([in] BSTR bstrActivePageID);
 };
```

I took the liberty to make them actual methods instead of just being `IDispatch` methods. **NOTE that this change is NOT OFFICIAL by Microsoft, so it could be incompatible with ANY version of OneNote.**

Compile it using `midl.exe` into `.tlb`, then set `HKCU\SOFTWARE\Classes\TypeLib\{...}` to point to the updated library. The registry can be set one a per-user basis, see [`Repair-OneNoteTypeLib.ps1`](Repair-OneNoteTypeLib.ps1).

## Problem 2: OneNote does not fire the events!

It seems that OneNote never calls into the sink, even when we have resolved the type library problem so `Advise` returns `S_OK`.

If we set breakpoints on `combase.dll!NdrProxyForwardingFunction` 3-5 and `combase.dll!ObjectSublessClient` 7-8 on the main thread, we can see that they are never called on the sink. (Those functions are the stubs used by the standard marshaler.)

## Code of `Advise`

When I was debugging, at one point I read the code of `Advise` method, so here it is. The `onmain.dll!Advise` is found by the following procedure:

1. Set a conditional breakpoint at `CStdIdentity::CInternalUnk::QueryInterface` with the following conditions:
  - it happens on the main thread; and
  - that `riid.Data1 == 0xe2e1511d` (querying `IID_IOneNoteEvents`).
2. `onmain.dll!Advise` = return address of this call minus 0x64.

```cpp
STDMETHODIMP /* virtual HRESULT __stdcall */
Advise(IUnknown *pUnkSink, DWORD *pdwCookie)
{
  HRESULT hr, hr2;
  /*** Prologue, setting up call security cookie. ***/
  // onmain.dll!Advise        push       rbx
  // onmain.dll!Advise+0x2    push       rsi
  // onmain.dll!Advise+0x3    push       rdi
  // onmain.dll!Advise+0x4    sub        rsp, 40h
  /*** The global variable name (__cookie) is randomly chosen. ***/
  // onmain.dll!Advise+0x8    mov        rax, qword ptr [onmain.dll!__cookie]
  // onmain.dll!Advise+0xf    xor        rax, rsp
  // onmain.dll!Advise+0x12   mov        qword ptr [rsp+38h], rax
  // onmain.dll!Advise+0x17   and        qword ptr [rsp+20h], 0
  // onmain.dll!Advise+0x1d   mov        rdi, r8
  // onmain.dll!Advise+0x20   mov        rbx, rdx
  // onmain.dll!Advise+0x23   mov        rsi, rcx
  // onmain.dll!Advise+0x26   test       r8, r8
  if (pdwCookie)
  // onmain.dll!Advise+0x29   je         onmain.dll!Advise+0x2f
  // onmain.dll!Advise+0x2b   and        dword ptr [r8], 0
    *pdwCookie = 0;
  // onmain.dll!Advise+0x2f   test       rbx, rbx
  if (!pUnkSink)
  // onmain.dll!Advise+0x32   je         onmain.dll!Advise+0xb0
    return E_POINTER; /*** goto return_POINTER; ***/
  // onmain.dll!Advise+0x34   test       rdi, rdi
  if (!pdwCookie)
  // onmain.dll!Advise+0x37   je         onmain.dll!Advise+0xb0
    return E_POINTER; /*** goto return_POINTER; ***/
  // onmain.dll!Advise+0x39   mov        rax, qword ptr [rcx]
  // onmain.dll!Advise+0x3c   lea        rdx, [rsp+28h]
  IID iid;
  // onmain.dll!Advise+0x41   mov        rax, qword ptr [rax+18h]
  // onmain.dll!Advise+0x45   call       ntdll.dll!LdrpDispatchUserCallTarget
  /*** iid = IID_IOneNoteEvents; ***/
  this->GetConnectionInterface(&iid);
  // onmain.dll!Advise+0x4b   mov        rax, qword ptr [rbx]
  // onmain.dll!Advise+0x4e   lea        r8, [rsp+20h]
  IOneNoteEvents *ppEvents;
  // onmain.dll!Advise+0x53   lea        rdx, [rsp+28h]
  // onmain.dll!Advise+0x58   mov        rcx, rbx
  // onmain.dll!Advise+0x5b   mov        rax, qword ptr [rax]
  // onmain.dll!Advise+0x5e   call       ntdll.dll!LdrpDispatchUserCallTarget
  /*** This will call into CStdIdentity::CInternalUnk::QueryInterface. ***/
  hr = pUnkSink->QueryInterface(iid, (void **)&ppEvents);
  // onmain.dll!Advise+0x64   mov        ebx, eax
  // onmain.dll!Advise+0x66   test       eax, eax
  // onmain.dll!Advise+0x68   js         onmain.dll!Advise+0x9b
  if (!SUCCEEDED(hr))
    /*** hr2 = hr; goto return_NOINTERFACE_FAIL_set_pdwCookie; ***/
    return (hr == E_NOINTERFACE ? E_NOINTERFACE : E_FAIL);
  // onmain.dll!Advise+0x6a   mov        rdx, qword ptr [rsp+20h]
  // onmain.dll!Advise+0x6f   lea        rcx, [rsi+8]
  /*** This line calls a global (helper) function.
  /*** It is chosen randomly (TakeSinkOwnership). ***/
  // onmain.dll!Advise+0x73   call       onmain.dll!TakeSinkOwnership
  DWORD dwCookie = TakeSinkOwnership(&this->SomeMember, ppEvents);
  // onmain.dll!Advise+0x78   mov        dword ptr [rdi], eax
  *pdwCookie = dwCookie;
  // onmain.dll!Advise+0x7a   test       eax, eax
  if (!dwCookie)
  // onmain.dll!Advise+0x7c   je         onmain.dll!Advise+0x82
    goto release_return_ERROR;
  // onmain.dll!Advise+0x7e   xor        ebx, ebx
  hr2 = S_OK;
  // onmain.dll!Advise+0x80   jmp        onmain.dll!Advise+0xac
  return hr2; /*** goto return_hr2; ***/

release_return_ERROR:
  // onmain.dll!Advise+0x82   mov        rcx, qword ptr [rsp+20h]
  // onmain.dll!Advise+0x87   mov        ebx, 80040201h
  hr2 = ONENOTE_ERROR_HR;
  // onmain.dll!Advise+0x8c   mov        rax, qword ptr [rcx]
  // onmain.dll!Advise+0x8f   mov        rax, qword ptr [rax+10h]
  // onmain.dll!Advise+0x93   call       ntdll.dll!LdrpDispatchUserCallTarget
  ppEvents->Release();
  // onmain.dll!Advise+0x99   jmp        onmain.dll!Advise+0xa9
  /*** *pdwCookie = 0; return hr2; ***/
  goto set_pdwCookie;

return_NOINTERFACE_FAIL_set_pdwCookie:
  // onmain.dll!Advise+0x9b   cmp        ebx, 80004002h
  // onmain.dll!Advise+0xa1   mov        eax, 80040202h
  // onmain.dll!Advise+0xa6   cmove      ebx, eax
  hr2 = (hr2 == E_NOINTERFACE ? E_NOINTERFACE : E_FAIL);

set_pdwCookie:
  // onmain.dll!Advise+0xa9   and        dword ptr [rdi], 0
  *pdwCookie = 0;

return_hr2:
  // onmain.dll!Advise+0xac   mov        eax, ebx
  hr = hr2;
  // onmain.dll!Advise+0xae   jmp        onmain.dll!Advise+0xb5
  /*** goto return_hr; ***/
  return hr;

return_POINTER:
  // onmain.dll!Advise+0xb0   mov        eax, 80004003h
  hr = E_POINTER;

return_hr:
  /*** Epilogue, check call security cookie. ***/
  // onmain.dll!Advise+0xb5   mov        rcx, qword ptr [rsp+38h]
  // onmain.dll!Advise+0xba   xor        rcx, rsp
  /*** The function name (__cookie_check) is randomly chosen.
  /*** It checks the global __cookie. ***/
  // onmain.dll!Advise+0xbd   call       onmain.dll!__cookie_check
  // onmain.dll!Advise+0xc2   add        rsp, 40h
  // onmain.dll!Advise+0xc6   pop        rdi
  // onmain.dll!Advise+0xc7   pop        rsi
  // onmain.dll!Advise+0xc8   pop        rbx
  // onmain.dll!Advise+0xc8   ret
  return hr;
}
```

In normal C++ code:

```cpp
STDMETHODIMP Advise(IUnknown *pUnkSink, DWORD *pdwCookie)
{
  if (!pdwCookie) return E_POINTER;
  if (!pUnkSink)
  {
    *pdwCookie = 0;
    return E_POINTER;
  }
  IID iid;
  this->GetConnectionInterface(&iid);
  IOneNoteEvents *ppEvents;
  /*** This step fails because of incomplete type library provided
  /*** by OneNote. It goes through once we fix the type library. ***/
  HRESULT hr = pUnkSink->QueryInterface(iid, (void **)&ppEvents);
  if (!SUCCEEDED(hr))
  {
    *pdwCookie = 0;
    return (hr == E_NOINTERFACE ? E_NOINTERFACE : E_FAIL);
  }
  /*** From the usage pattern, it can be inferred that
  /*** TakeSinkOwnership does not call AddRef on ppEvents
  /*** even if it decides to connect to the sink. ***/
  if (!(*pdwCookie = TakeSinkOwnership(&this->SomeMember, ppEvents)))
  {
    ppEvents->Release();
    return ONENOTE_ERROR_HR;
  }
  return S_OK;
}
```
