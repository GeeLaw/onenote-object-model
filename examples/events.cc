#include"utils.hpp"
using namespace OneNote15;

struct MyEvents : public IOneNoteEvents
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
    if (riid == IID_IOneNoteEvents)
    {
      *ppv = static_cast<IOneNoteEvents *>(this);
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
      *ppv = static_cast<IUnknown *>(this);
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
  STDMETHODIMP GetTypeInfo(UINT uTInfo, LCID lcid, ITypeInfo **ppTInfo)
  {
    wprintf(L"GetTypeInfo()\n");
    *ppTInfo = NULL;
    return E_NOTIMPL;
  }
  STDMETHODIMP GetIDsOfNames(REFIID riid,
    LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
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
        L"ONNAVIGATE", L"onnavigate"))
      {
        rgDispId[i] = DISPID_OnNavigate;
      }
      else if (StringCompareMixTwo(rgszNames[i],
        L"ONHIERARCHYCHANGE", L"onhierarchychange"))
      {
        rgDispId[i] = DISPID_OnHierarchyChange;
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
    if (pExcepInfo)
    {
      memset(pExcepInfo, sizeof(EXCEPINFO), 0);
    }
    return S_OK;
  }
  /*** IOneNoteEvents ***/
  STDMETHODIMP OnNavigate()
  {
    wprintf(L"OnNavigate()\n");
    return S_OK;
  }
  STDMETHODIMP OnHierarchyChange(BSTR bstrActivePageID)
  {
    wprintf(L"OnHierarchyChange()\n");
    return S_OK;
  }
  /*** C++ ***/
  MyEvents(ComPtr<IApplication> const &onenote)
    : refcount(1), app(onenote)
  {
  }
  MyEvents(MyEvents &&) = delete;
  MyEvents &operator =(MyEvents &&) = delete;
  MyEvents(MyEvents const &) = delete;
  MyEvents &operator =(MyEvents const &) = delete;
  ~MyEvents() = default;
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
  ThreadTimer timer;
  if (!SUCCEEDED(hr = (HRESULT)timer))
  {
    wprintf(L"SetTimer failed with HRESULT %08x.\n", (int)hr);
    return (int)hr;
  }
  ComPtr<IApplication> onenote;
  if (!SUCCEEDED(hr = CreateApplication(onenote.ppiOut())))
  {
    wprintf(L"CoCreateInstance returned %08x.\n", (int)hr);
    return (int)hr;
  }
  ComPtr<IConnectionPointContainer> app;
  if (!SUCCEEDED(hr = onenote->QueryInterface(IID_IConnectionPointContainer,
    app.ppvOut())))
  {
    wprintf(L"QueryInterface for IConnectionPointContainer returned %08x.\n",
      (int)hr);
    return (int)hr;
  }
  ComPtr<IConnectionPoint> connpoint;
  if (!SUCCEEDED(hr = app->FindConnectionPoint(IID_IOneNoteEvents,
    connpoint.ppiOut())))
  {
    wprintf(L"FindConnectionPoint returned %08x.\n", (int)hr);
    return (int)hr;
  }
  MyEvents events(onenote);
  AdvisoryConenction advisory(connpoint, &events);
  /* The type library provided by OneNote is incomplete,
  /* and this call will fail. Applying the fixed library
  /* makes the call succeed, but the events are never fired. */
  if (!SUCCEEDED(hr = (HRESULT)advisory))
  {
    wprintf(L"Advise returned %08x.\n", (int)hr);
    return (int)hr;
  }
  bool hit;
  MSG msg;
  while (!(hit = (bool)_kbhit()) && GetMessage(&msg, NULL, 0, 0))
  {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  if (hit)
  {
    _getch();
  }
  return 0;
}
