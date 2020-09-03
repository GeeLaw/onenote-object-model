COM_DECLARE_ENUM(HtmlStatus, "4C5A2A84-A20C-4CFE-8A54-0C5D3BF3CA2C")
{
  hsHtmlOk = 0x80042050,
  hsHtmlNoDocument = 0x80042051,
  hsHtmlDocumentNotFocused = 0x80042052,
  hsHtmlNoSelection = 0x80042053,
  hsHtmlWacSite = 0x80042054,
  hsHtmlPasswordField = 0x80042055
} HtmlStatus;

COM_DECLARE_INTERFACE(IApplicationEx,
  "8BF94B48-1E76-4AA3-AB1D-463F49B3E681",
  0x8BF94B48,0x1E76,0x4AA3,0xAB,0x1D,0x46,0x3F,0x49,0xB3,0xE6,0x81)
{
  COM_DECLARE_METHOD(IsFirstRunComplete, 0x60020000,
    /* [out] */ ::VARIANT_BOOL *pfFirstRunCompleteOut,
    /* [out] */ ::BSTR *pbstrErrorMessage);
  COM_DECLARE_METHOD(IsValidWebContent, 0x60020001,
    /* [in] */ ::BSTR bstrHtml,
    /* [in] */ HtmlStatus hsHtmlStatus,
    /* [out] */ ::VARIANT_BOOL *pfValid,
    /* [out] */ ::BSTR *pbstrErrorMessage);
  COM_DECLARE_METHOD(SendWebContentToOneNote, 0x60020002,
    /* [in] */ ::BSTR bstrUrl,
    /* [in] */ ::BSTR bstrTitle,
    /* [in] */ ::BSTR bstrHtml,
    /* [in] */ ::VARIANT_BOOL fRtl,
    /* [in] */ ::UINT64 hwndParent);
} IApplicationEx;
