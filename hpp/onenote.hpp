#ifndef ONENOTE_HPP
#define ONENOTE_HPP

#include<cstdint>
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

#define COM_DECLARE_DISP_INTERFACE(name, iid, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8, ...) \
  COM_DECLARE_GUID(IID, DIID_##name, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8); \
  typedef struct __declspec(uuid(iid)) __declspec(novtable) name : public ::IDispatch

/* All methods (except for dispinterface methods) return HRESULT. */
#define COM_DECLARE_METHOD(name, dispid, ...) \
  COM_DECLARE_DISP_METHOD(name, dispid); \
  public: virtual ::HRESULT __stdcall name(__VA_ARGS__) = 0

#define COM_DEFINE_EXT_METHOD(name, ...) \
  public: inline ::HRESULT name(__VA_ARGS__)

#define COM_DECLARE_DISP_METHOD(name, dispid, ...) \
  public: static constexpr ::DISPID const DISPID_##name = dispid

#define COM_DECLARE_CLASS(cls, intf, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
  COM_DECLARE_GUID(CLSID, CLSID_##cls, l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8); \
  static inline ::HRESULT Create##cls(intf **p##intf, \
    ::DWORD dwClsContext = 0x17 /* CLSCTX_ALL */) \
  { \
    return ::CoCreateInstance(CLSID_##cls, NULL, \
      dwClsContext, IID_##intf, (void **)p##intf); \
  }

namespace OneNote12
{
#include"onenote12_.hpp"
}

namespace OneNote15
{
#include"onenote15_.hpp"
}

namespace OneNote15Ex
{
#include"onenote15ex_.hpp"
}

#undef COM_DECLARE_GUID
#undef COM_DECLARE_ENUM
#undef COM_DECLARE_HRESULT
#undef COM_FWD_DECLARE_INTERFACE
#undef COM_DECLARE_INTERFACE
#undef COM_DECLARE_DISP_INTERFACE
#undef COM_DECLARE_METHOD
#undef COM_DECLARE_DISP_INTERFACE
#undef COM_DECLARE_CLASS

#endif
