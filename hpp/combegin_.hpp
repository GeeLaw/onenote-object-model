#include<objbase.h>
#include<oaidl.h>

#ifndef INITGUID
#define INITGUID
#include<guiddef.h>
#undef INITGUID
#else
#include<guiddef.h>
#endif

#define COM_DECLARE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
  static constexpr type const name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#define COM_DECLARE_ENUM(name, guid) \
  typedef enum __declspec(uuid(guid)) name

#define COM_DECLARE_HRESULT(name, value) \
  static constexpr ::HRESULT const hr##name = value

#define COM_FWD_DECLARE_INTERFACE(name) struct name

/* All the interfaces used by OneNote directly derives from IDispatch. */
#define COM_DECLARE_INTERFACE(name, iid, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8, ...) \
  COM_DECLARE_GUID(IID, IID_##name, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8); \
  typedef struct __declspec(uuid(iid)) __declspec(novtable) name : public ::IDispatch

/* All methods in OneNote return HRESULT and support IDispatch. */
#define COM_DECLARE_METHOD(name, dispid, ...) \
  public: \
  static constexpr ::DISPID const DISPID_##name = dispid; \
  virtual ::HRESULT __stdcall name(__VA_ARGS__) = 0

#define COM_DEFINE_EXT_METHOD(name, ...) \
  public: inline ::HRESULT name(__VA_ARGS__)

#define COM_DECLARE_CLASS(cls, intf, iid, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
  COM_DECLARE_GUID(CLSID, CLSID_##cls, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8); \
  static inline ::HRESULT Create##cls(intf **ppi, \
    ::DWORD dwClsContext = 0x5 \
    /* CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER */) \
  { \
    return ::CoCreateInstance(CLSID_##cls, NULL, \
      dwClsContext, iid, (void **)ppi); \
  }
