COM_DECLARE_ENUM(HierarchyScope, "552E0E02-B287-4EC6-9CC0-4BA019EE5EA1")
{
  hsSelf = 0,
  hsChildren = 1,
  hsNotebooks = 2,
  hsSections = 3,
  hsPages = 4
} HierarchyScope;

COM_DECLARE_ENUM(HierarchyElement, "41C8F6EA-0AF0-4A4F-99E9-5EB01EBFC9A3")
{
  heNone = 0,
  heNotebooks = 1,
  heSectionGroups = 2,
  heSections = 4,
  hePages = 8
} HierarchyElement;

COM_DECLARE_ENUM(RecentResultType, "4DB67B4F-CC7D-45B5-88FE-569AE5798FF2")
{
  rrtNone = 0,
  rrtFiling = 1,
  rrtSearch = 2,
  rrtLinks = 3
} RecentResultType;

COM_DECLARE_ENUM(CreateFileType, "B5EB9D34-5278-4D8A-AE1F-2F88EA56BBCE")
{
  cftNone = 0,
  cftNotebook = 1,
  cftFolder = 2,
  cftSection = 3
} CreateFileType;

COM_DECLARE_ENUM(PageInfo, "D6E78E55-7EE7-4A31-BF3E-B01E819599BA")
{
  piBasic = 0,
  piBinaryData = 1,
  piSelection = 2,
  piFileType = 4,
  piBinaryDataSelection = 3,
  piBinaryDataFileType = 5,
  piSelectionFileType = 6,
  piAll = 7
} PageInfo;

COM_DECLARE_ENUM(PublishFormat, "D6166973-3665-4EDB-94B0-77C65C34B51C")
{
  pfOneNote = 0,
  pfOneNotePackage = 1,
  pfMHTML = 2,
  pfPDF = 3,
  pfXPS = 4,
  pfWord = 5,
  pfEMF = 6,
  pfHTML = 7,
  pfOneNote2007 = 8
} PublishFormat;

COM_DECLARE_ENUM(SpecialLocation, "E195F3E3-8EC3-4A67-81FE-DDBEC2B42D3F")
{
  slBackUpFolder = 0,
  slUnfiledNotesSection = 1,
  slDefaultNotebookFolder = 2
} SpecialLocation;

COM_DECLARE_ENUM(FilingLocation, "452D048E-7F61-4258-94B9-A39E19C290DA")
{
  flEMail = 0,
  flContacts = 1,
  flTasks = 2,
  flMeetings = 3,
  flWebContent = 4,
  flPrintOuts = 5
} FilingLocation;

COM_DECLARE_ENUM(FilingLocationType, "82FC5A95-FEB7-4242-95E1-369C5DFE3F49")
{
  fltNamedSectionNewPage = 0,
  fltCurrentSectionNewPage = 1,
  fltCurrentPage = 2,
  fltNamedPage = 4
} FilingLocationType;

COM_DECLARE_ENUM(NewPageStyle, "226CC8E6-1ED0-4770-A7F1-A80BB4DDF07B")
{
  npsDefault = 0,
  npsBlankPageWithTitle = 1,
  npsBlankPageNoTitle = 2
} NewPageStyle;

COM_DECLARE_ENUM(DockLocation, "B67BC7E1-91B9-4F50-8471-77C76F8D63D6")
{
  dlDefault = 0xffffffff,
  dlNone = 0,
  dlLeft = 1,
  dlRight = 2,
  dlTop = 3,
  dlBottom = 4
} DockLocation;

COM_DECLARE_ENUM(XMLSchema, "68555133-B62F-4490-9D66-9E9BFC68F6C6")
{
  xs2007 = 0,
  xs2010 = 1,
  xs2013 = 2,
  xsCurrent = 2
} XMLSchema;

COM_DECLARE_ENUM(TreeCollapsedStateType, "1ECC88B3-6D2B-4EDD-8DD5-BB11E5D34C09")
{
  tcsExpanded = 0,
  tcsCollapsed = 1
} TreeCollapsedStateType;

COM_DECLARE_ENUM(NotebookFilterOutType, "13F18B04-E76F-42E0-97E6-8B6ABF38B484")
{
  nfoLocal = 1,
  nfoNetwork = 2,
  nfoWeb = 4,
  nfoNoWacUrl = 8
} NotebookFilterOutType;

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
  COM_DECLARE_HRESULT(MergeFailed, 0x8004201f);
  COM_DECLARE_HRESULT(InvalidXMLSchema, 0x80042020);
  COM_DECLARE_HRESULT(FutureContentLoss, 0x80042022);
  COM_DECLARE_HRESULT(TimeOut, 0x80042023);
  COM_DECLARE_HRESULT(RecordingInProgress, 0x80042024);
  COM_DECLARE_HRESULT(UnknownLinkedNoteState, 0x80042025);
  COM_DECLARE_HRESULT(NoShortNameForLinkedNote, 0x80042026);
  COM_DECLARE_HRESULT(NoFriendlyNameForLinkedNote, 0x80042027);
  COM_DECLARE_HRESULT(InvalidLinkedNoteUri, 0x80042028);
  COM_DECLARE_HRESULT(InvalidLinkedNoteThumbnail, 0x80042029);
  COM_DECLARE_HRESULT(ImportLNTThumbnailFailed, 0x8004202a);
  COM_DECLARE_HRESULT(UnreadDisabledForNotebook, 0x8004202b);
  COM_DECLARE_HRESULT(InvalidSelection, 0x8004202c);
  COM_DECLARE_HRESULT(ConvertFailed, 0x8004202d);
  COM_DECLARE_HRESULT(RecycleBinEditFailed, 0x8004202e);
  COM_DECLARE_HRESULT(AppInModalUI, 0x80042030);
}

typedef struct Point
{
  long x;
  long y;
} Point;

typedef uint64_t uint64;

COM_FWD_DECLARE_INTERFACE(IQuickFilingDialog);
COM_FWD_DECLARE_INTERFACE(IQuickFilingDialogCallback);
COM_FWD_DECLARE_INTERFACE(IApplication);
COM_FWD_DECLARE_INTERFACE(IWindows);
COM_FWD_DECLARE_INTERFACE(IWindow);
COM_FWD_DECLARE_INTERFACE(IOneNoteEvents);

COM_DECLARE_INTERFACE(IQuickFilingDialog,
  "1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9",
  0x1D12BD3F,0x89B6,0x4077,0xAA,0x2C,0xC9,0xDC,0x2B,0xCA,0x42,0xF9)
{
  COM_DECLARE_METHOD(get_Title, 0,
    /* [out, retval] */ ::BSTR *bstrTitle);
  COM_DECLARE_METHOD(put_Title, 0,
    /* [in] */ ::BSTR bstrTitle);
  COM_DECLARE_METHOD(get_Description, 0x00000001,
    /* [out, retval] */ ::BSTR *bstrDescription);
  COM_DECLARE_METHOD(put_Description, 0x00000001,
    /* [in] */ ::BSTR bstrDescription);
  COM_DECLARE_METHOD(get_CheckboxText, 0x00000002,
    /* [out, retval] */ ::BSTR *bstrText);
  COM_DECLARE_METHOD(put_CheckboxText, 0x00000002,
    /* [in] */ ::BSTR bstrText);
  COM_DECLARE_METHOD(get_CheckboxState, 0x00000003,
    /* [out, retval] */ ::VARIANT_BOOL *pfChecked);
  COM_DECLARE_METHOD(put_CheckboxState, 0x00000003,
    /* [in] */ ::VARIANT_BOOL pfChecked);
  COM_DECLARE_METHOD(get_WindowHandle, 0x00000004,
    /* [out, retval] */ uint64 *pHWNDWindow);
  COM_DECLARE_METHOD(get_TreeDepth, 0x00000005,
    /* [out, retval] */ HierarchyElement *pTreeDepth);
  COM_DECLARE_METHOD(put_TreeDepth, 0x00000005,
    /* [in] */ HierarchyElement pTreeDepth);
  COM_DECLARE_METHOD(get_ParentWindowHandle, 0x00000006,
    /* [out, retval] */ uint64 *pHWNDParentWindow);
  COM_DECLARE_METHOD(put_ParentWindowHandle, 0x00000006,
    /* [in] */ uint64 pHWNDParentWindow);
  COM_DECLARE_METHOD(get_Position, 0x00000007,
    /* [out, retval] */ Point *pPoint);
  COM_DECLARE_METHOD(put_Position, 0x00000007,
    /* [in] */ Point pPoint);
  COM_DECLARE_METHOD(SetRecentResults, 0x00000008,
    /* [in] */ RecentResultType recentResults,
    /* [in] */ ::VARIANT_BOOL fShowCurrentSection,
    /* [in] */ ::VARIANT_BOOL fShowCurrentPage,
    /* [in] */ ::VARIANT_BOOL fShowUnfiledNotes);
  COM_DECLARE_METHOD(AddButton, 0x0000000a,
    /* [in] */ ::BSTR bstrText,
    /* [in] */ HierarchyElement allowedElements,
    /* [in] */ HierarchyElement allowedReadOnlyElements,
    /* [in] */ ::VARIANT_BOOL fDefault);
  COM_DECLARE_METHOD(Run, 0x0000000b,
    /* [in] */ IQuickFilingDialogCallback *piCallback);
  COM_DECLARE_METHOD(get_SelectedItem, 0x0000000c,
    /* [out, retval] */ ::BSTR *pbstrSelectedNodeID);
  COM_DECLARE_METHOD(get_PressedButton, 0x0000000d,
    /* [out, retval] */ unsigned long *pButtonIndex);
  COM_DECLARE_METHOD(put_TreeCollapsedState, 0x0000000e,
    /* [in] */ TreeCollapsedStateType rhs);
  COM_DECLARE_METHOD(put_NotebookFilterOut, 0x0000000f,
    /* [in] */ NotebookFilterOutType rhs);
  COM_DECLARE_METHOD(ShowCreateNewNotebook, 0x00000010);
  COM_DECLARE_METHOD(AddInitialEditor, 0x00000011,
    ::BSTR initialEditor);
  COM_DECLARE_METHOD(ClearInitialEditors, 0x00000012);
  COM_DECLARE_METHOD(ShowSharingHyperlink, 0x00000013);
} IQuickFilingDialog;

COM_DECLARE_INTERFACE(IQuickFilingDialogCallback,
  "627EA7B4-95B5-4980-84C1-9D20DA4460B1",
  0x627EA7B4,0x95B5,0x4980,0x84,0xC1,0x9D,0x20,0xDA,0x44,0x60,0xB1)
{
  COM_DECLARE_METHOD(OnDialogClosed, 0x60020000,
    /* [in] */ IQuickFilingDialog *dialog);
} IQuickFilingDialogCallback;

namespace Impl_
{
  static constexpr ::BSTR const bstrEmpty = (::BSTR)NULL;
}

COM_DECLARE_INTERFACE(IApplication,
  "452AC71A-B655-4967-A208-A4CC39DD7949",
  0x452AC71A,0xB655,0x4967,0xA2,0x08,0xA4,0xCC,0x39,0xDD,0x79,0x49)
{
  COM_DECLARE_METHOD(GetHierarchy, 0x60020000,
    /* [in] */ ::BSTR bstrStartNodeID,
    /* [in] */ HierarchyScope hsScope,
    /* [out] */ ::BSTR *pbstrHierarchyXmlOut,
    /* [in, optional, defaultvalue(2)] */ XMLSchema xsSchema = xsCurrent);
  COM_DECLARE_METHOD(UpdateHierarchy, 0x60020001,
    /* [in] */ ::BSTR bstrChangesXmlIn,
    /* [in, optional, defaultvalue(2)] */ XMLSchema xsSchema = xsCurrent);
  COM_DECLARE_METHOD(OpenHierarchy, 0x60020002,
    /* [in] */ ::BSTR bstrPath,
    /* [in] */ ::BSTR bstrRelativeToObjectID,
    /* [out] */ ::BSTR *pbstrObjectID,
    /* [in, optional, defaultvalue(0)] */ CreateFileType cftIfNotExist = cftNone);
  COM_DECLARE_METHOD(DeleteHierarchy, 0x60020003,
    /* [in] */ ::BSTR bstrObjectID,
    /* [in, optional, defaultvalue(12:00:00 AM)] */ ::DATE dateExpectedLastModified = 0,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL deletePermanently = 0);
  COM_DECLARE_METHOD(CreateNewPage, 0x60020004,
    /* [in] */ ::BSTR bstrSectionID,
    /* [out] */ ::BSTR *pbstrPageID,
    /* [in, optional, defaultvalue(0)] */ NewPageStyle npsNewPageStyle = npsDefault);
  COM_DECLARE_METHOD(CloseNotebook, 0x60020005,
    /* [in] */ ::BSTR bstrNotebookID,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL force = 0);
  COM_DECLARE_METHOD(GetHierarchyParent, 0x60020006,
    /* [in] */ ::BSTR bstrObjectID,
    /* [out] */ ::BSTR *pbstrParentID);
  COM_DECLARE_METHOD(GetPageContent, 0x60020007,
    /* [in] */ ::BSTR bstrPageID,
    /* [out] */ ::BSTR *pbstrPageXmlOut,
    /* [in, optional, defaultvalue(0)] */ PageInfo pageInfoToExport = piBasic,
    /* [in, optional, defaultvalue(2)] */ XMLSchema xsSchema = xsCurrent);
  COM_DECLARE_METHOD(UpdatePageContent, 0x60020008,
    /* [in] */ ::BSTR bstrPageChangesXmlIn,
    /* [in, optional, defaultvalue(12:00:00 AM)] */ ::DATE dateExpectedLastModified = 0,
    /* [in, optional, defaultvalue(2)] */ XMLSchema xsSchema = xsCurrent,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL force = 0);
  COM_DECLARE_METHOD(GetBinaryPageContent, 0x60020009,
    /* [in] */ ::BSTR bstrPageID,
    /* [in] */ ::BSTR bstrCallbackID,
    /* [out] */ ::BSTR *pbstrBinaryObjectB64Out);
  COM_DECLARE_METHOD(DeletePageContent, 0x6002000a,
    /* [in] */ ::BSTR bstrPageID,
    /* [in] */ ::BSTR bstrObjectID,
    /* [in, optional, defaultvalue(12:00:00 AM)] */ ::DATE dateExpectedLastModified = 0,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL force = 0);
  COM_DECLARE_METHOD(NavigateTo, 0x6002000b,
    /* [in] */ ::BSTR bstrHierarchyObjectID,
    /* [in, optional, defaultvalue("")] */ ::BSTR bstrObjectID = Impl_::bstrEmpty,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fNewWindow = 0);
  COM_DECLARE_METHOD(NavigateToUrl, 0x6002000c,
    /* [in] */ ::BSTR bstrUrl,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fNewWindow = 0);
  COM_DECLARE_METHOD(Publish, 0x6002000d,
    /* [in] */ ::BSTR bstrHierarchyID,
    /* [in] */ ::BSTR bstrTargetFilePath,
    /* [in, optional, defaultvalue(0)] */ PublishFormat pfPublishFormat = pfOneNote,
    /* [in, optional, defaultvalue("")] */ ::BSTR bstrCLSIDofExporter = Impl_::bstrEmpty);
  COM_DECLARE_METHOD(OpenPackage, 0x6002000e,
    /* [in] */ ::BSTR bstrPathPackage,
    /* [in] */ ::BSTR bstrPathDest,
    /* [out] */ ::BSTR *pbstrPathOut);
  COM_DECLARE_METHOD(GetHyperlinkToObject, 0x6002000f,
    /* [in] */ ::BSTR bstrHierarchyID,
    /* [in] */ ::BSTR bstrPageContentObjectID,
    /* [out] */ ::BSTR *pbstrHyperlinkOut);
  COM_DECLARE_METHOD(FindPages, 0x60020010,
    /* [in] */ ::BSTR bstrStartNodeID,
    /* [in] */ ::BSTR bstrSearchString,
    /* [out] */ ::BSTR *pbstrHierarchyXmlOut,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fIncludeUnindexedPages = 0,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fDisplay = 0,
    /* [in, optional, defaultvalue(2)] */ XMLSchema xsSchema = xsCurrent);
  COM_DECLARE_METHOD(FindMeta, 0x60020011,
    /* [in] */ ::BSTR bstrStartNodeID,
    /* [in] */ ::BSTR bstrSearchStringName,
    /* [out] */ ::BSTR *pbstrHierarchyXmlOut,
    /* [in, optional, defaultvalue(0)] */ ::VARIANT_BOOL fIncludeUnindexedPages = 0,
    /* [in, optional, defaultvalue(2)] */ XMLSchema xsSchema = xsCurrent);
  COM_DECLARE_METHOD(GetSpecialLocation, 0x60020012,
    /* [in] */ SpecialLocation slToGet,
    /* [out] */ ::BSTR *pbstrSpecialLocationPath);
  COM_DECLARE_METHOD(MergeFiles, 0x60020013,
    /* [in] */ ::BSTR bstrBaseFile,
    /* [in] */ ::BSTR bstrClientFile,
    /* [in] */ ::BSTR bstrServerFile,
    /* [in] */ ::BSTR bstrTargetFile);
  COM_DECLARE_METHOD(QuickFiling, 0x60020014,
    /* [out, retval] */ IQuickFilingDialog **ppiDialog);
  COM_DECLARE_METHOD(SyncHierarchy, 0x60020015,
    /* [in] */ ::BSTR bstrHierarchyID);
  COM_DECLARE_METHOD(SetFilingLocation, 0x60020016,
    /* [in] */ FilingLocation flToSet,
    /* [in] */ FilingLocationType fltToSet,
    /* [in] */ ::BSTR bstrFilingSectionID);
  COM_DECLARE_METHOD(get_Windows, 0x00000064,
    /* [out, retval] */ IWindows **ppONWindows);
  COM_DECLARE_METHOD(get_Dummy1, 0x00000066,
    /* [out, retval] */ ::VARIANT_BOOL *pBool);
  COM_DECLARE_METHOD(MergeSections, 0x60020019,
    /* [in] */ ::BSTR bstrSectionSourceId,
    /* [in] */ ::BSTR bstrSectionDestinationId);
  COM_DECLARE_METHOD(get_COMAddIns, 0x00000068,
    /* [out, retval] */ IDispatch **ppiComAddins);
  COM_DECLARE_METHOD(get_LanguageSettings, 0x00000069,
    /* [out, retval] */ IDispatch **ppiLanguageSettings);
  COM_DECLARE_METHOD(GetWebHyperlinkToObject, 0x6002001c,
    /* [in] */ ::BSTR bstrHierarchyID,
    /* [in] */ ::BSTR bstrPageContentObjectID,
    /* [out] */ ::BSTR *pbstrHyperlinkOut);
} IApplication;

COM_DECLARE_INTERFACE(IWindows,
  "6D4B9C3E-CC05-493F-85E2-43D1006DF96A",
  0x6D4B9C3E,0xCC05,0x493F,0x85,0xE2,0x43,0xD1,0x00,0x6D,0xF9,0x6A)
{
  COM_DECLARE_METHOD(get_Item, 0,
    /* [in] */ unsigned long Index,
    /* [out, retval] */ IWindow **Item);
  COM_DECLARE_METHOD(get_Count, 0x00000001,
    /* [out, retval] */ unsigned long *Count);
  COM_DECLARE_METHOD(get__NewEnum, 0xfffffffc,
    /* [out, retval] */ IUnknown **_NewEnum);
  COM_DECLARE_METHOD(get_CurrentWindow, 0x00000003,
    /* [out, retval] */ IWindow **ppCurrentWindow);
} Windows;

COM_DECLARE_INTERFACE(IWindow,
  "8E8304B8-CBD1-44F8-B0E8-89C625B2002E",
  0x8E8304B8,0xCBD1,0x44F8,0xB0,0xE8,0x89,0xC6,0x25,0xB2,0x00,0x2E)
{
  COM_DECLARE_METHOD(get_WindowHandle, 0,
    /* [out, retval] */ uint64 *pHWNDWindow);
  COM_DECLARE_METHOD(get_CurrentPageId, 0x00000001,
    /* [out, retval] */ ::BSTR *pbstrPageObjectId);
  COM_DECLARE_METHOD(get_CurrentSectionId, 0x00000002,
    /* [out, retval] */ ::BSTR *pbstrSectionObjectId);
  COM_DECLARE_METHOD(get_CurrentSectionGroupId, 0x00000003,
    /* [out, retval] */ ::BSTR *pbstrSectionObjectId);
  COM_DECLARE_METHOD(get_CurrentNotebookId, 0x00000004,
    /* [out, retval] */ ::BSTR *pbstrNotebookObjectId);
  COM_DECLARE_METHOD(NavigateTo, 0x00000009,
    /* [in] */ ::BSTR bstrHierarchyObjectID,
    /* [in, optional, defaultvalue("")] */ ::BSTR bstrObjectID = Impl_::bstrEmpty);
  COM_DECLARE_METHOD(get_FullPageView, 0x0000000a,
    /* [out, retval] */ ::VARIANT_BOOL *pIsFullPageView);
  COM_DECLARE_METHOD(put_FullPageView, 0x0000000a,
    ::VARIANT_BOOL pIsFullPageView);
  COM_DECLARE_METHOD(get_Active, 0x0000000b,
    /* [out, retval] */ ::VARIANT_BOOL *pIsActive);
  COM_DECLARE_METHOD(put_Active, 0x0000000b,
    ::VARIANT_BOOL pIsActive);
  COM_DECLARE_METHOD(get_DockedLocation, 0x0000000d,
    /* [out, retval] */ DockLocation *pDockLocation);
  COM_DECLARE_METHOD(put_DockedLocation, 0x0000000d,
    DockLocation pDockLocation);
  COM_DECLARE_METHOD(get_Application, 0x0000000e,
    /* [out, retval] */ IApplication **ppiApp);
  COM_DECLARE_METHOD(get_SideNote, 0x0000000f,
    /* [out, retval] */ ::VARIANT_BOOL *pIsSideNote);
  COM_DECLARE_METHOD(NavigateToUrl, 0x00000010,
    /* [in] */ ::BSTR bstrUrl);
  COM_DECLARE_METHOD(SetDockedLocation, 0x00000011,
    /* [in] */ DockLocation DockLocation,
    /* [in] */ Point ptMonitor);
} Window;

COM_DECLARE_DISP_INTERFACE(IOneNoteEvents,
  "E2E1511D-502D-4BD0-8B3A-8A89A05CDCAE",
  0xE2E1511D,0x502D,0x4BD0,0x8B,0x3A,0x8A,0x89,0xA0,0x5C,0xDC,0xAE)
{
  /* These two methods return void, not HRESULT. */
  COM_DECLARE_DISP_METHOD(OnNavigate, 0x00000001);
  COM_DECLARE_DISP_METHOD(OnHierarchyChange, 0x00000002,
    /* [in] */ BSTR bstrActivePageID);
} IOneNoteEvents;

COM_DECLARE_CLASS(Application, IApplication,
  0xD7FAC39E,0x7FF1,0x49AA,0x98,0xCF,0xA1,0xDD,0xD3,0x16,0x33,0x7E)

COM_DECLARE_CLASS(Application2, IApplication,
  0xDC67E480,0xC3CB,0x49F8,0x82,0x32,0x60,0xB0,0xC2,0x05,0x6C,0x8E)
