#pragma warning(disable: 4100)
#pragma comment(lib, "user32")
#pragma comment(lib, "ole32")
#pragma comment(lib, "oleaut32")
#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include<windows.h>
#include<objbase.h>
#include<oaidl.h>
#include<ocidl.h>
#include<oleauto.h>
#include"../hpp/onenote.hpp"
#include<stdio.h>
#include<conio.h>

template <typename TIntf>
struct ComPtr
{
  ComPtr() : target(NULL) { }
  ComPtr(TIntf *&&ptr) : target(ptr) { ptr = NULL; }
  explicit ComPtr(TIntf * const &ptr) : target(ptr)
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->AddRef();
    }
  }
  ComPtr(ComPtr &&other) : target(other.Detach()) { }
  ComPtr(ComPtr const &other) : target(other.target)
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->AddRef();
    }
  }
  ~ComPtr()
  {
    Clear();
  }
  ComPtr &operator =(ComPtr const &other)
  {
    CopyFrom(other.target);
    return *this;
  }
  ComPtr &operator =(ComPtr &&other)
  {
    TIntf *tmp = other.target;
    other.target = target;
    target = tmp;
    return *this;
  }
  void Attach(TIntf *ptr)
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->Release();
    }
    target = ptr;
  }
  TIntf *Detach()
  {
    TIntf *ret = target;
    target = NULL;
    return ret;
  }
  void CopyFrom(TIntf *ptr)
  {
    if (ptr)
    {
      static_cast<IUnknown *>(ptr)->AddRef();
    }
    if (target)
    {
      static_cast<IUnknown *>(target)->Release();
    }
    target = ptr;
  }
  TIntf *Copy()
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->AddRef();
    }
    return target;
  }
  void Clear()
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->Release();
      target = NULL;
    }
  }
  void *pvIn() { return target; }
  TIntf *piIn() { return target; }
  void **ppvInOut() { return reinterpret_cast<void **>(&target); }
  TIntf **ppiInOut() { return &target; }
  void **ppvOut()
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->Release();
    }
    return reinterpret_cast<void **>(&target);
  }
  TIntf **ppiOut()
  {
    if (target)
    {
      static_cast<IUnknown *>(target)->Release();
    }
    return &target;
  }
  TIntf * const operator ->() const { return target; }
  TIntf &operator *() const { return *target; }
  operator bool() const { return target != NULL; }
  bool operator !() const { return target == NULL; }
private:
  TIntf *target;
};

struct InitializeOLE
{
  InitializeOLE() : hr(OleInitialize(NULL)) { }
  InitializeOLE(InitializeOLE &&) = delete;
  InitializeOLE &operator =(InitializeOLE &&) = delete;
  InitializeOLE(InitializeOLE const &) = delete;
  InitializeOLE &operator =(InitializeOLE const &) = delete;
  ~InitializeOLE() { if (SUCCEEDED(hr)) { OleUninitialize(); } }
  operator HRESULT() const { return hr; }
private:
  HRESULT const hr;
};

struct ThreadTimer
{
  ThreadTimer() : timer(SetTimer(NULL, NULL, 200, NULL)),
    hr(timer ? S_OK : HRESULT_FROM_WIN32(GetLastError()))
  {
  }
  ThreadTimer(ThreadTimer &&) = delete;
  ThreadTimer &operator =(ThreadTimer &&) = delete;
  ThreadTimer(ThreadTimer const &) = delete;
  ThreadTimer &operator =(ThreadTimer const &) = delete;
  ~ThreadTimer() { if (SUCCEEDED(hr)) { KillTimer(NULL, timer); } }
  operator HRESULT() const { return hr; }
private:
  UINT_PTR timer;
  HRESULT hr;
};

struct BStrSysFreeString
{
  BStrSysFreeString(BSTR str) : bstr(str) { }
  BStrSysFreeString() = delete;
  BStrSysFreeString(BStrSysFreeString &&) = delete;
  BStrSysFreeString &operator =(BStrSysFreeString &&) = delete;
  BStrSysFreeString(BStrSysFreeString const &) = delete;
  BStrSysFreeString &operator =(BStrSysFreeString const &) = delete;
  ~BStrSysFreeString() { SysFreeString(bstr); }
private:
  BSTR bstr;
};

bool StringCompareMixTwo(wchar_t const *a, wchar_t const *b, wchar_t const *c)
{
  while (*b && (*a == *b || *a == *c))
  {
    ++a;
    ++b;
    ++c;
  }
  return *a == *b || *a == *c;
}

struct AdvisoryConenction
{
  AdvisoryConenction(ComPtr<IConnectionPoint> const &connpoint, IUnknown *punk)
    : target(connpoint), cookie(0), hr(connpoint->Advise(punk, &cookie))
  {
  }
  AdvisoryConenction(AdvisoryConenction &&) = delete;
  AdvisoryConenction &operator =(AdvisoryConenction &&) = delete;
  AdvisoryConenction(AdvisoryConenction const &) = delete;
  AdvisoryConenction &operator =(AdvisoryConenction const &) = delete;
  operator HRESULT() const { return hr; }
  ~AdvisoryConenction()
  {
    if (SUCCEEDED(hr))
    {
      target->Unadvise(cookie);
    }
  }
private:
  ComPtr<IConnectionPoint> const target;
  DWORD cookie;
  HRESULT const hr;
};
