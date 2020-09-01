COM_DECLARE_ENUM(OneNoteAddIn_Event, "AAE363E2-3D91-4B0C-9021-EFDA0ACBD858")
{
  evtAddInNavigation = 0,
  evtAddInHierarchyChange = 1
} OneNoteAddIn_Event;

COM_DECLARE_ENUM(HierarchyScope, "97CB5BF9-BF0C-47E5-A9BB-6B189BCA3C25")
{
  hsSelf = 0,
  hsChildren = 1,
  hsNotebooks = 2,
  hsSections = 3,
  hsPages = 4
} HierarchyScope;

COM_DECLARE_ENUM(CreateFileType, "65D0EDAB-9D2F-479C-81C2-E0B481734320")
{
  cftNone = 0,
  cftNotebook = 1,
  cftFolder = 2,
  cftSection = 3
} CreateFileType;

COM_DECLARE_ENUM(NewPageStyle, "47656582-CDF9-4933-86A5-4D192D6AC8CC")
{
  npsDefault = 0,
  npsBlankPageWithTitle = 1,
  npsBlankPageNoTitle = 2
} NewPageStyle;

COM_DECLARE_ENUM(PageInfo, "8E4BA554-9AC4-4E7B-B6E6-39705192F8D1")
{
  piBasic = 0,
  piBinaryData = 1,
  piSelection = 2,
  piBinaryDataSelection = 3,
  piAll = 3
} PageInfo;

COM_DECLARE_ENUM(PublishFormat, "B6876F11-4F18-4913-BF40-7698D08C791D")
{
  pfOneNote = 0,
  pfOneNotePackage = 1,
  pfMHTML = 2,
  pfPDF = 3,
  pfXPS = 4,
  pfWord = 5,
  pfEMF = 6
} PublishFormat;

COM_DECLARE_ENUM(SpecialLocation, "2A0B42F4-2F24-4392-A800-F0A979424A57")
{
  slBackUpFolder = 0,
  slUnfiledNotesSection = 1,
  slDefaultNotebookFolder = 2
} SpecialLocation;

namespace Error
{
  COM_DECLARE_HRESULT(MalformedXML, 0x80042000);
  COM_DECLARE_HRESULT(InvalidXML, 0x80042001);
  COM_DECLARE_HRESULT(CreatingSection, 0x80042002);
  COM_DECLARE_HRESULT(OpeningSection, 0x80042003);
  COM_DECLARE_HRESULT(SectionDoesNotExist, 0x80042004);
  COM_DECLARE_HRESULT(PageDoesNotExist, 0x80042005);
  COM_DECLARE_HRESULT(FileDoesNotExist, 0x80042006);
  COM_DECLARE_HRESULT(InsertingImage, 0x80042007);
  COM_DECLARE_HRESULT(InsertingInk, 0x80042008);
  COM_DECLARE_HRESULT(InsertingHtml, 0x80042009);
  COM_DECLARE_HRESULT(NavigatingToPage, 0x8004200a);
  COM_DECLARE_HRESULT(SectionReadOnly, 0x8004200b);
  COM_DECLARE_HRESULT(PageReadOnly, 0x8004200c);
  COM_DECLARE_HRESULT(InsertingOutlineText, 0x8004200d);
  COM_DECLARE_HRESULT(PageObjectDoesNotExist, 0x8004200e);
  COM_DECLARE_HRESULT(BinaryObjectDoesNotExist, 0x8004200f);
  COM_DECLARE_HRESULT(LastModifiedDateDidNotMatch, 0x80042010);
  COM_DECLARE_HRESULT(GroupDoesNotExist, 0x80042011);
  COM_DECLARE_HRESULT(PageDoesNotExistInGroup, 0x80042012);
  COM_DECLARE_HRESULT(NoActiveSelection, 0x80042013);
  COM_DECLARE_HRESULT(ObjectDoesNotExist, 0x80042014);
  COM_DECLARE_HRESULT(NotebookDoesNotExist, 0x80042015);
  COM_DECLARE_HRESULT(InsertingFile, 0x80042016);
  COM_DECLARE_HRESULT(InvalidName, 0x80042017);
  COM_DECLARE_HRESULT(FolderDoesNotExist, 0x80042018);
  COM_DECLARE_HRESULT(InvalidQuery, 0x80042019);
  COM_DECLARE_HRESULT(FileAlreadyExists, 0x8004201a);
  COM_DECLARE_HRESULT(SectionEncryptedAndLocked, 0x8004201b);
  COM_DECLARE_HRESULT(DisabledByPolicy, 0x8004201c);
  COM_DECLARE_HRESULT(NotYetSynchronized, 0x8004201d);
  COM_DECLARE_HRESULT(LegacySection, 0x8004201e);
}

COM_DECLARE_INTERFACE(IOneNoteAddIn,
  "C9590FA7-2132-47FB-9A78-AF0BF19AF4E6",
  0xC9590FA7,0x2132,0x47FB,0x9A,0x78,0xAF,0x0B,0xF1,0x9A,0xF4,0xE6)
{
  COM_DECLARE_METHOD(OnClick, 0x60020000,
    /* [in] */ ::BSTR bstrActivePageID,
    /* [out, retval] */ ::VARIANT_BOOL *pfEnabled);
  COM_DECLARE_METHOD(OnEvent, 0x60020001,
    /* [in] */ OneNoteAddIn_Event evt,
    /* [in] */ ::BSTR bstrParameter,
    /* [out, retval] */ ::VARIANT_BOOL *pfEnabled);
} IOneNoteAddIn;

COM_DECLARE_INTERFACE(IApplication,
  "2DA16203-3F58-404F-839D-E4CDE7DD0DED",
  0x2DA16203,0x3F58,0x404F,0x83,0x9D,0xE4,0xCD,0xE7,0xDD,0x0D,0xED)
{
  COM_DECLARE_METHOD(GetHierarchy, 0x60020000,
    /* [in] */ ::BSTR bstrStartNodeID,
    /* [in] */ HierarchyScope hsScope,
    /* [out] */ ::BSTR *pbstrHierarchyXmlOut);
  COM_DECLARE_METHOD(UpdateHierarchy, 0x60020001,
    /* [in] */ ::BSTR bstrChangesXmlIn);
  COM_DECLARE_METHOD(OpenHierarchy, 0x60020002,
    /* [in] */ ::BSTR bstrPath,
    /* [in] */ ::BSTR bstrRelativeToObjectID,
    /* [out] */ ::BSTR *pbstrObjectID,
    /* [in, optional, defaultvalue(0)] */ CreateFileType cftIfNotExist = cftNone);
  COM_DECLARE_METHOD(DeleteHierarchy, 0x60020003,
    /* [in] */ ::BSTR bstrObjectID,
    /* [in, optional, defaultvalue(12:00:00 AM)] */ ::DATE dateExpectedLastModified = 0);
  COM_DECLARE_METHOD(CreateNewPage, 0x60020004,
    /* [in] */ ::BSTR bstrSectionID,
    /* [out] */ ::BSTR *pbstrPageID,
    /* [in, optional, defaultvalue(0)] */ NewPageStyle npsNewPageStyle = npsDefault);
  COM_DECLARE_METHOD(CloseNotebook, 0x60020005,
    /* [in] */ ::BSTR bstrNotebookID);
  COM_DECLARE_METHOD(GetHierarchyParent, 0x60020006,
    /* [in] */ ::BSTR bstrObjectID,
    /* [out] */ ::BSTR *pbstrParentID);
  COM_DECLARE_METHOD(GetPageContent, 0x60020007,
    /* [in] */ ::BSTR bstrPageID,
    /* [out] */ ::BSTR *pbstrPageXmlOut,
    /* [in, optional, defaultvalue(0)] */ PageInfo pageInfoToExport = piBasic);
  COM_DECLARE_METHOD(UpdatePageContent, 0x60020008,
    /* [in] */ ::BSTR bstrPageChangesXmlIn,
    /* [in, optional, defaultvalue(12:00:00 AM)] */ ::DATE dateExpectedLastModified = 0);
  COM_DECLARE_METHOD(GetBinaryPageContent, 0x60020009,
    /* [in] */ ::BSTR bstrPageID,
    /* [in] */ ::BSTR bstrCallbackID,
    /* [out] */ ::BSTR *pbstrBinaryObjectB64Out);
  COM_DECLARE_METHOD(DeletePageContent, 0x6002000a,
    /* [in] */ ::BSTR bstrPageID,
    /* [in] */ ::BSTR bstrObjectID,
    /* [in, optional, defaultvalue(12:00:00 AM)] */ ::DATE dateExpectedLastModified = 0);
  COM_DECLARE_METHOD(NavigateTo, 0x6002000b,
    /* [in] */ ::BSTR bstrHierarchyObjectID,
    /* [in, optional, defaultvalue("0")] */ ::BSTR bstrObjectID,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fNewWindow = 0);
  COM_DECLARE_METHOD(Publish, 0x6002000c,
    /* [in] */ ::BSTR bstrHierarchyID,
    /* [in] */ ::BSTR bstrTargetFilePath,
    /* [in, optional, defaultvalue(0)] */ PublishFormat pfPublishFormat,
    /* [in, optional, defaultvalue("0")] */ ::BSTR bstrCLSIDofExporter);
  COM_DECLARE_METHOD(OpenPackage, 0x6002000d,
    /* [in] */ ::BSTR bstrPathPackage,
    /* [in] */ ::BSTR bstrPathDest,
    /* [out] */ ::BSTR *pbstrPathOut);
  COM_DECLARE_METHOD(GetHyperlinkToObject, 0x6002000e,
    /* [in] */ ::BSTR bstrHierarchyID,
    /* [in] */ ::BSTR bstrPageContentObjectID,
    /* [out] */ ::BSTR *pbstrHyperlinkOut);
  COM_DECLARE_METHOD(FindPages, 0x6002000f,
    /* [in] */ ::BSTR bstrStartNodeID,
    /* [in] */ ::BSTR bstrSearchString,
    /* [out] */ ::BSTR *pbstrHierarchyXmlOut,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fIncludeUnindexedPages = 0,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fDisplay = 0);
  COM_DECLARE_METHOD(FindMeta, 0x60020010,
    /* [in] */ ::BSTR bstrStartNodeID,
    /* [in] */ ::BSTR bstrSearchStringName,
    /* [out] */ ::BSTR *pbstrHierarchyXmlOut,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fIncludeUnindexedPages = 0);
  COM_DECLARE_METHOD(GetSpecialLocation, 0x60020011,
    /* [in] */ SpecialLocation slToGet,
    /* [out] */ ::BSTR *pbstrSpecialLocationPath);

  /*** Handle default arguments. ***/
  COM_DEFINE_EXT_METHOD(NavigateTo,
    /* [in] */ ::BSTR bstrHierarchyObjectID)
  {
    ::BSTR bstr0;
    if (!(bstr0 = ::SysAllocString(L"0")))
    {
      return 0x8007000E /* E_OUTOFMEMORY */;
    }
    ::HRESULT hr = NavigateTo(bstrHierarchyObjectID, bstr0, 0);
    ::SysFreeString(bstr0);
    return hr;
  }
  COM_DEFINE_EXT_METHOD(Publish,
    /* [in] */ ::BSTR bstrHierarchyID,
    /* [in] */ ::BSTR bstrTargetFilePath,
    /* [in, optional, defaultvalue(0)] */ PublishFormat pfPublishFormat = pfOneNote)
  {
    ::BSTR bstr0;
    if (!(bstr0 = ::SysAllocString(L"0")))
    {
      return 0x8007000E /* E_OUTOFMEMORY */;
    }
    ::HRESULT hr = Publish(bstrHierarchyID, bstrTargetFilePath, pfPublishFormat, bstr0);
    ::SysFreeString(bstr0);
    return hr;
  }
} IApplication;

COM_DECLARE_CLASS(Application, IApplication,
  0x0039FFEC,0xA022,0x4232,0x82,0x74,0x6B,0x34,0x78,0x7B,0xFC,0x27)
