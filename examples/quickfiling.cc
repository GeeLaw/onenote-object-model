#include"utils.hpp"
using namespace OneNote15;

struct DialogCallback : public IQuickFilingDialogCallback
{
  /*** IUnknown ***/
  STDMETHODIMP QueryInterface(REFIID riid, void **ppv)
  {
    wprintf(L"QueryInterface(IID = "
      "%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX"
      ")\n",
      riid.Data1, riid.Data2, riid.Data3, 
      riid.Data4[0], riid.Data4[1], riid.Data4[2], riid.Data4[3],
      riid.Data4[4], riid.Data4[5], riid.Data4[6], riid.Data4[7]);
    if (!ppv)
    {
      return E_POINTER;
    }
    if (riid == IID_IQuickFilingDialogCallback)
    {
      *ppv = static_cast<IQuickFilingDialogCallback *>(this);
      AddRef();
      return S_OK;
    }
    if (riid == IID_IDispatch)
    {
      *ppv = static_cast<IDispatch *>(this);
      AddRef();
      return S_OK;
    }
    if (riid == IID_IUnknown)
    {
      *ppv = static_cast<IUnknown *>(static_cast<IDispatch *>(this));
      AddRef();
      return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
  }
  STDMETHODIMP_(ULONG) AddRef() { wprintf(L"AddRef()\n"); return ++refcount; }
  STDMETHODIMP_(ULONG) Release() { wprintf(L"Release()\n"); return --refcount; }
  /*** IDispatch ***/
  STDMETHODIMP GetTypeInfoCount(UINT *pctinfo)
  {
    wprintf(L"GetTypeInfoCount()\n");
    *pctinfo = 0;
    return E_NOTIMPL;
  }
  STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
  {
    wprintf(L"GetTypeInfo()\n");
    *ppTInfo = NULL;
    return E_NOTIMPL;
  }
  STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames,
    LCID lcid, DISPID *rgDispId)
  {
    wprintf(L"GetIDsOfNames()\n");
    if (riid != IID_NULL)
    {
      return E_INVALIDARG;
    }
    bool unk = false;
    for (UINT i = 0; i != cNames; ++i)
    {
      if (StringCompareMixTwo(rgszNames[i],
        L"ONDIALOGCLOSED", L"ondialogclosed"))
      {
        rgDispId[i] = DISPID_OnDialogClosed;
      }
      else
      {
        rgDispId[i] = DISPID_UNKNOWN;
        unk = true;
      }
    }
    return unk ? DISP_E_UNKNOWNNAME : S_OK;
  }
  STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
    WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult,
    EXCEPINFO *pExcepInfo, UINT *puArgErr)
  {
    wprintf(L"Invoke(DispId = %d)\n", (int)dispIdMember);
    if (pVarResult)
    {
      VariantInit(pVarResult);
    }
    if (dispIdMember != DISPID_OnDialogClosed || wFlags != DISPATCH_METHOD)
    {
      return DISP_E_MEMBERNOTFOUND;
    }
    if (pDispParams->cNamedArgs != 0)
    {
      return DISP_E_NONAMEDARGS;
    }
    if (pDispParams->cArgs != 1)
    {
      return DISP_E_BADPARAMCOUNT;
    }
    VARIANT &arg0 = pDispParams->rgvarg[0];
    IUnknown *punk = NULL;
    if (arg0.vt == VT_UNKNOWN)
    {
      punk = arg0.punkVal;
    }
    else if (arg0.vt == VT_DISPATCH)
    {
      punk = static_cast<IUnknown *>(arg0.pdispVal);
    }
    HRESULT hr = DISP_E_TYPEMISMATCH;
    ComPtr<IQuickFilingDialog> qfd;
    if (!punk || !SUCCEEDED(hr = punk->QueryInterface(
      IID_IQuickFilingDialog, qfd.ppvOut())))
    {
      return hr;
    }
    return OnDialogClosed(&*qfd);
  }
  /*** IQuickFilingDialogCallback ***/
  STDMETHODIMP OnDialogClosed(IQuickFilingDialog *qfd)
  {
    wprintf(L"OnDialogClosed()\n");
    PostQuitMessage(0);
    HRESULT hr;
    BSTR item, link;
    if (!SUCCEEDED(hr = qfd->get_SelectedItem(&item)))
    {
      wprintf(L"get_SelectedItem returned %08x.\n", (int)hr);
      return hr;
    }
    BStrSysFreeString freeItem(item);
    if (item == NULL || SysStringLen(item) == 0)
    {
      wprintf(L"--- You cancelled the operation. ---\n");
      return S_OK;
    }
    if (!SUCCEEDED(hr = app->GetWebHyperlinkToObject(item, NULL, &link)))
    {
      wprintf(L"GetWebHyperlinkToObject returned %08x.\n", (int)hr);
      return hr;
    }
    BStrSysFreeString freeLink(link);
    wprintf(L"--- %ls\n", link);
    return S_OK;
  }
  /*** C++ ***/
  DialogCallback(ComPtr<IApplication> const &onenote)
    : refcount(1), app(onenote)
  {
  }
  DialogCallback(DialogCallback &&) = delete;
  DialogCallback &operator =(DialogCallback &&) = delete;
  DialogCallback(DialogCallback const &) = delete;
  DialogCallback &operator =(DialogCallback const &) = delete;
private:
  ULONG refcount;
  ComPtr<IApplication> app;
};

int main()
{
  HRESULT hr;
  InitializeOLE init;
  if (!SUCCEEDED(hr = (HRESULT)init))
  {
    wprintf(L"OleInitialize returned %08x.\n", (int)hr);
    return (int)hr;
  }
  ComPtr<IApplication> onenote;
  if (!SUCCEEDED(hr = CreateApplication(onenote.ppiOut())))
  {
    wprintf(L"CoCreateInstance returned %08x.\n", (int)hr);
    return (int)hr;
  }
  ComPtr<IQuickFilingDialog> dialog;
  if (!SUCCEEDED(hr = onenote->QuickFiling(dialog.ppiOut())))
  {
    wprintf(L"QuickFiling returned %08x.\n", (int)hr);
    return (int)hr;
  }
  BSTR title = SysAllocString(L"Choose a location");
  if (!title)
  {
    wprintf(L"SysAllocString failed with HRESULT %08x.\n",
      (int)(hr = HRESULT_FROM_WIN32(GetLastError())));
    return hr;
  }
  BStrSysFreeString freeTitle(title);
  if (!SUCCEEDED(hr = dialog->put_Title(title)))
  {
    wprintf(L"put_Title returned %08x.\n", (int)hr);
    return (int)hr;
  }
  BSTR desc = SysAllocString(L"Choose a location to obtain its link:");
  if (!desc)
  {
    wprintf(L"SysAllocString failed with HRESULT %08x.\n",
      (int)(hr = HRESULT_FROM_WIN32(GetLastError())));
    return hr;
  }
  BStrSysFreeString freeDesc(desc);
  if (!SUCCEEDED(hr = dialog->put_Description(desc)))
  {
    wprintf(L"put_Description returned %08x.\n", (int)hr);
    return (int)hr;
  }
  BSTR btn = SysAllocString(L"OK");
  if (!btn)
  {
    wprintf(L"SysAllocString failed with HRESULT %08x.\n",
      (int)(hr = HRESULT_FROM_WIN32(GetLastError())));
    return hr;
  }
  BStrSysFreeString freeBtn(btn);
  if (!SUCCEEDED(hr = dialog->AddButton(btn,
    (HierarchyElement)(heNotebooks | heSectionGroups | heSections | hePages),
    (HierarchyElement)(heNotebooks | heSectionGroups | heSections | hePages),
    0xffff)))
  {
    wprintf(L"AddButton returned %08x.\n", (int)hr);
    return (int)hr;
  }
  DialogCallback callback(onenote);
  if (!SUCCEEDED(hr = dialog->Run(&callback)))
  {
    wprintf(L"Run returned %08x.\n", (int)hr);
    return (int)hr;
  }
  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0))
  {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return 0;
}
