class ZCL_XLSX_STRING_UTIL definition
  public
  create public .

public section.

  methods INIT_FROM_XML
    importing
      !IV_SHAREDSTRINGS_XML type XSTRING .
  methods GET_INDEX_OF_STRING
    importing
      !IV_STRING type CLIKE
    returning
      value(RV_INDEX) type I .
  methods TO_XML
    returning
      value(RV_SHAREDSTRINGS_XML) type XSTRING .
  type-pools ABAP .
  methods HAS_STRINGS
    returning
      value(RV_HAS_STRINGS) type ABAP_BOOL .
  methods GET_STRING_AT_INDEX
    importing
      !IV_STRING_INDEX type I
    returning
      value(RV_STRING) type STRING .
protected section.


  types:
    BEGIN OF gty_s_string_r,
           string TYPE string,
         END OF gty_s_string_r .
  TYPES gty_t_string_r TYPE STANDARD TABLE OF gty_s_string_r WITH NON-UNIQUE KEY string.
  types:
    BEGIN OF gty_s_string,
           string TYPE string,
           r TYPE gty_t_string_r,
         END OF gty_s_string .
  types:
    BEGIN OF gty_s_string_index,
           string TYPE string,
           index  TYPE i,
         END OF gty_s_string_index .
  types:
    gty_th_string_index TYPE HASHED TABLE OF gty_s_string_index
                            WITH UNIQUE KEY string .
  types:
    gty_t_strings TYPE TABLE OF gty_s_string .

  types:
    BEGIN OF gty_s_shared_strings,
           string_tab TYPE TABLE OF gty_s_string WITH DEFAULT KEY,
           string_count TYPE i,
           string_ucount TYPE i,
         END OF gty_s_shared_strings .

  data MT_STRINGS type GTY_T_STRINGS .
  data MTH_STRING_INDEX type GTY_TH_STRING_INDEX .
private section.

ENDCLASS.



CLASS ZCL_XLSX_STRING_UTIL IMPLEMENTATION.


METHOD get_index_of_string.
**********************************************************************
* DATA DEFINITION
**********************************************************************
  DATA: lr_s_string_index         TYPE REF TO gty_s_string_index.
  DATA: lr_s_string               TYPE REF TO gty_s_string.

**********************************************************************
* FUNCTIONAL BODY
**********************************************************************

* check whether the string is already present in the string table
  READ TABLE me->mth_string_index WITH TABLE KEY string = iv_string REFERENCE INTO lr_s_string_index.
  IF sy-subrc = 0.
*   the string is already included in the shared strings table. set the index
    rv_index = lr_s_string_index->index.
  ELSE.
*   the string is currently not in the shared string table. include it
    APPEND INITIAL LINE TO me->mt_strings REFERENCE INTO lr_s_string.
    lr_s_string->string = iv_string.
*   set the index starting with 0
    rv_index = lines( me->mt_strings ) - 1.
*   now insert the new string into the hash table (register for later access)
    CREATE DATA lr_s_string_index.
    lr_s_string_index->index = rv_index.
    lr_s_string_index->string = iv_string.
    INSERT lr_s_string_index->* INTO TABLE me->mth_string_index.
  ENDIF.


ENDMETHOD.


METHOD get_string_at_index.
**********************************************************************
* DATA DEFINITION
**********************************************************************
  DATA: lv_index    TYPE i.
  DATA: lr_s_string TYPE REF TO gty_s_string.

**********************************************************************
* FUNCTIONAL BODY
**********************************************************************

* If index is out of range, empty string will be returned
  CLEAR rv_string.

  IF iv_string_index LT 0.
    RETURN.
  ENDIF.

* String ids are 0 based whereas abap indexes are 1 based
  lv_index = iv_string_index + 1.

* Check if index is in bounds of internal table and return it if found
  IF lv_index LE lines( mt_strings ).
    READ TABLE mt_strings INDEX lv_index REFERENCE INTO lr_s_string.
    rv_string = lr_s_string->string.
  ENDIF.

ENDMETHOD.


method HAS_STRINGS.
  rv_has_strings = abap_false.
  IF lines( mt_strings ) > 0.
    rv_has_strings = abap_true.
  ENDIF.
endmethod.


METHOD init_from_xml.
**********************************************************************
* DATA DEFINITION
**********************************************************************
 DATA: ls_string_index LIKE LINE OF mth_string_index.
 DATA: lr_s_string TYPE REF TO gty_s_string.
**********************************************************************
* FUNCTIONAL BODY
**********************************************************************
  CLEAR: mt_strings, mth_string_index.

  " Read the strings from the shared strings part
  CALL TRANSFORMATION xl_mpa_get_strings SOURCE XML iv_sharedstrings_xml RESULT strings = mt_strings.

  " Create a lookup table for the indexes of the strings in the list so we can
  " determine a strings index later on.
  " Note: We must not sort the original table of strings because otherwise we would have to
  " adjust the corresponding indexes in all sheets when we update the shared strings part.
  LOOP AT mt_strings REFERENCE INTO lr_s_string.
    IF lr_s_string->r IS INITIAL.
      ls_string_index-string = lr_s_string->string.
    ELSE.
*     For multi-formatted strings, create a guuid
      TRY.
        ls_string_index-string = cl_system_uuid=>create_uuid_x16_static(  ).
      CATCH cx_root.ENDTRY.
    ENDIF.
    ls_string_index-index = sy-tabix - 1.   " The string count starts with 0
    INSERT ls_string_index INTO TABLE mth_string_index.
  ENDLOOP.

ENDMETHOD.


METHOD to_xml.
**********************************************************************
* DATA DEFINITION
**********************************************************************
  DATA: ls_sharedstrings TYPE gty_s_shared_strings.

**********************************************************************
* FUNCTIONAL BODY
**********************************************************************

  " Update the shared strings since we may have added some
  ls_sharedstrings-string_tab = mt_strings.
  ls_sharedstrings-string_count = lines( mt_strings ).
  ls_sharedstrings-string_ucount = lines( mth_string_index ).

  CALL TRANSFORMATION xl_mpa_create_strings SOURCE param = ls_sharedstrings RESULT XML rv_sharedstrings_xml.

ENDMETHOD.
ENDCLASS.
