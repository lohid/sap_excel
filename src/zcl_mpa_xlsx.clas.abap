class ZCL_MPA_XLSX definition
  public
  create private .

public section.

  types:
    BEGIN OF gty_s_sheet_info,
      name     TYPE string,
      sheet_id TYPE i,
      rid      TYPE string,
    END OF gty_s_sheet_info .
  types:
    gty_th_sheet_info TYPE TABLE OF gty_s_sheet_info
                          WITH KEY name .

  class-methods GET_INSTANCE
    returning
      value(RO_XLSX) type ref to ZCL_MPA_XLSX .
  methods CREATE_DOC
    returning
      value(RO_DOC) type ref to zif_mpa_XLSX_DOC .
  methods LOAD_DOC
    importing
      !IV_FILE_DATA type XSTRING
    returning
      value(RO_DOC) type ref to zif_mpa_XLSX_DOC
    raising
      CX_OPENXML_FORMAT
      CX_OPENXML_NOT_ALLOWED
      CX_DYNAMIC_CHECK .
  methods GET_INPUT_TYPE_FOR_ABAP_TYPE
    importing
      !IO_ABAP_TYPE_DESC type ref to CL_ABAP_TYPEDESCR
    returning
      value(RV_TYPE) type zif_mpa_XLSX_SHEET=>CELLTYPE .
  PROTECTED SECTION.
  PRIVATE SECTION.
    CLASS-DATA: mo_instance TYPE REF TO zcl_mpa_xlsx.
ENDCLASS.



CLASS ZCL_MPA_XLSX IMPLEMENTATION.


  METHOD create_doc.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lo_doc TYPE REF TO lcl_xlsx_doc.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    CREATE OBJECT lo_doc.
    lo_doc->mo_xlsx_document = cl_xlsx_document=>create_document(
                           ).
    TRY.
        lo_doc->initialize( ).
      CATCH cx_openxml_not_found
            cx_openxml_format
            cx_openxml_not_allowed.
*    There is an implementation error - assert
      ASSERT 1 = 2.
    ENDTRY.
    ro_doc = lo_doc.
  ENDMETHOD.


  METHOD get_input_type_for_abap_type.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: ls_ddic_header TYPE x030l.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Assume charlike as Fallback
    rv_type = zif_mpa_xlsx_sheet=>gc_celltype-charlike.

    CASE io_abap_type_desc->type_kind.    "#EC CI_INT8_OK
      WHEN cl_abap_typedescr=>typekind_date.
*       Set Date Type for Date
        rv_type = zif_mpa_xlsx_sheet=>gc_celltype-date.
      WHEN cl_abap_typedescr=>typekind_time.
*       Set Time Type for Time
        rv_type = zif_mpa_xlsx_sheet=>gc_celltype-time.
      WHEN cl_abap_typedescr=>typekind_num.
*       Set Numeric Char Type for Nums
        rv_type = zif_mpa_xlsx_sheet=>gc_celltype-numericchar.
      WHEN cl_abap_typedescr=>typekind_decfloat OR
           cl_abap_typedescr=>typekind_decfloat16 OR
           cl_abap_typedescr=>typekind_decfloat34 OR
           cl_abap_typedescr=>typekind_float.
*       Assign Float Type to all Floating point numbers
        rv_type = zif_mpa_xlsx_sheet=>gc_celltype-float.
      WHEN cl_abap_typedescr=>typekind_int OR
           cl_abap_typedescr=>typekind_int1 OR
           cl_abap_typedescr=>typekind_int2.
*       Assign Integer type to all kinds of integers
        rv_type = zif_mpa_xlsx_sheet=>gc_celltype-integer.
      WHEN cl_abap_typedescr=>typekind_packed.
        " code for other data types including INT8
*       If we have a packed type
*       Check if it is a timestamp
        ls_ddic_header = io_abap_type_desc->get_ddic_header( ).
        IF sy-subrc = 0 AND cl_ehfnd_exp_utilities=>is_time_stamp_domain( ls_ddic_header-refname ) EQ abap_true.
          rv_type = zif_mpa_xlsx_sheet=>gc_celltype-timestamp.
        ENDIF.
    ENDCASE.
  ENDMETHOD.


  METHOD get_instance.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    IF mo_instance IS INITIAL.
      CREATE OBJECT mo_instance.
    ENDIF.
    ro_xlsx = mo_instance.
  ENDMETHOD.


  METHOD load_doc.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lo_doc TYPE REF TO lcl_xlsx_doc.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    CREATE OBJECT lo_doc.
    lo_doc->mo_xlsx_document = cl_xlsx_document=>load_document( iv_file_data ).
    TRY.
        lo_doc->initialize( ).
      CATCH cx_openxml_not_found.

*    There is an implementation error - assert
      ASSERT 1 = 2.
    ENDTRY.
    ro_doc = lo_doc.
  ENDMETHOD.
ENDCLASS.
