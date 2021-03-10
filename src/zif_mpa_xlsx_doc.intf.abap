interface ZIF_MPA_XLSX_DOC
  public .


  methods SAVE
    returning
      value(RV_FILE_DATA) type XSTRING
    raising
      CX_OPENXML_FORMAT
      CX_OPENXML_NOT_FOUND
      CX_OPENXML_NOT_ALLOWED
      CX_DYNAMIC_CHECK .
  methods GET_SHEETS
    returning
      value(RT_SHEET_INFO) type ZCL_mpa_XLSX=>GTY_TH_SHEET_INFO
    raising
      CX_DYNAMIC_CHECK .
  methods GET_SHEET_BY_ID
    importing
      !IV_SHEET_ID type I
    returning
      value(RO_SHEET) type ref to zif_mpa_XLSX_SHEET
    raising
      CX_OPENXML_FORMAT
      CX_OPENXML_NOT_FOUND
      CX_DYNAMIC_CHECK .
  methods GET_SHEET_BY_NAME
    importing
      !IV_SHEET_NAME type STRING
    returning
      value(RO_SHEET) type ref to zif_mpa_XLSX_SHEET
    raising
      CX_OPENXML_FORMAT
      CX_OPENXML_NOT_FOUND
      CX_DYNAMIC_CHECK .
  methods ADD_NEW_SHEET
    importing
      !IV_SHEET_NAME type STRING
    returning
      value(RO_SHEET) type ref to zif_mpa_XLSX_SHEET
    raising
      CX_OPENXML_FORMAT
      CX_OPENXML_NOT_ALLOWED
      CX_DYNAMIC_CHECK .
endinterface.
