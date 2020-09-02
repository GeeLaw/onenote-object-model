# Interfaces and co-classes

Note that OneNote interfaces and co-classes have not changed since OneNote 2013, so do not freak out if we are talking 7-year-old things in 2020 (when you have OneNote... 2016)!

## OneNote 2007&ndash;2013

One thing that bothers me is that OneNote 2013 does not provide OneNote type library 14 (i.e., for OneNote 2010). It turns out that OneNote 2013 is a simple extension of OneNote 2010, so they share most interface definitions. This can be reassured by the fact that version 15 shares the same IDs with [version 14](https://docs.microsoft.com/en-us/archive/blogs/descapa/cc-add-ins-for-onenote-2010) (IDs for COM elements are frozen once published).

It also turns out that nothing in OneNote type library 12, except numerical values of enumerations and HRESULTs, is compatible with those in versions 14 and 15. This is checked by calling `QueryInterface` on new/old objects for new/old interfaces. So `IApplication` for version 12 can only be used with `OneNote.Application.12`.

Since the interfaces inherit `IDispatch` and extended methods only gain optional parameters in the new version, code in scripting languages largely remains source-code compatible.

The content below applies to version 14 and 15. (Remember that there is no Office version 13.)

- The `Application` object (with CLSID `D7FA...337E`) is for OneNote 2010.
  - It implements `IApplication` (with IID `452A...7949`) and `IOneNoteEvents`.
- The `Application2` object (with CLSID `DC67...6C8E`) is for OneNote 2013.
  - It implements `IApplication`, `IOneNoteEvents`, and `IApplicationEx`.
- OneNote 2013 also implements the objects for OneNote 2007 and 2010.
  - In general, it is possible to create lower-versioned objects with higher-versioned software, so you can write the program targeting OneNote 2007 (with CLSID `0039...FC27`), and it will run fine with OneNote 2010 and 2013.
  - On the other hand, you can detect the presence of OneNote 2013 by trying creating `Application2` object. (Note that doing so will run OneNote, so do not use it to just check the version &mdash; use it when you need to access a feature only available in OneNote 2013.)

A gotcha: OneNote 2010 comes before 2013, so if you try passing `XMLSchema` value `xs2013` to `IApplication` interface to OneNote 2010 object, the call will fail. This can be again confirmed by the fact that `xs2013` did not exist in the [header file](https://msdnshared.blob.core.windows.net/media/MSDNBlogsFS/prod.evol.blogs.msdn.com/CommunityServer.Components.PostAttachments/00/10/19/99/73/onenote14-x86.h) for OneNote 2010 and that `xsCurrent` **referred** to `xs2010`.

Another gotcha regarding header files and scripting: If we set a default value in the header files and compiled with the header file (as of the writing of this documentation) with default value for `XMLSchema` being `xsCurrent`, the program would always use 2013 format instead of the latest available from the object. This is very undesirable because it will never work with OneNote 2010. In contrast, scripting language uses `IDispatch` with the latest type library, so the default value is always correct. For this reason, the default arguments are completely stripped off the version 15 header (since it is shared with version 14).

Additional resources from Microsoft:

- OneNote object model
  - [OneNote 2007](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/ms788684(v=office.12))
  - [OneNote 2010](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff700506(v=office.14))
  - [OneNote (Win32)](https://docs.microsoft.com/en-us/office/client-developer/onenote/onenote-home) has the latest documentation for OneNote object model.
- XML schemas
  - [OneNote 2007](https://docs.microsoft.com/en-us/archive/blogs/descapa/final-onenote-2007-xml-schema-posted), all the links broken. Other posts: [final namespace](https://docs.microsoft.com/en-us/archive/blogs/descapa/final-onenote-2007-schema-namespace), [beta 2 namespace](https://docs.microsoft.com/en-us/archive/blogs/descapa/onenote-2007-beta2-xml-schema-reference).
  - [OneNote 2010](https://docs.microsoft.com/en-us/archive/blogs/descapa/onenote-2010-xml-schema)
  - [OneNote 2013](https://docs.microsoft.com/en-us/archive/blogs/descapa/onenote-2013-com-api-xml-schema-onenote-2013-xsd)
- [Header file](https://docs.microsoft.com/en-us/archive/blogs/descapa/cc-add-ins-for-onenote-2010) for OneNote 2010

## OneNote 2003 SP1 (pre-history)

There was a `CSimpleImporterClass` object that allowed adding content to OneNote 2003 SP1. However, the only documentation I could find are these:

- [Importing Content into OneNote 2003 SP1](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa168020(v=office.11)) introduces the object.
- [OneNote 2003 XML Schema](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa203882(v=office.11))
