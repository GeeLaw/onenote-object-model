# C++ header library

The headers in this folder are created manually from the IDL. They are a lightweight header-only library to automating OneNote from C++.

The benefit of this library over the usual MIDL-to-C++ preprocessing technique is that name collisions (e.g., `IID_IApplication` and `CLSID_Application`) are avoided using namespaces.
