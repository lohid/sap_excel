*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations
TYPES:
  BEGIN OF gty_s_workbook_data,
    active_sheet   TYPE i,
    sheet_ids_htab TYPE zcl_mpa_xlsx=>gty_th_sheet_info,
  END OF gty_s_workbook_data .


TYPES:
  BEGIN OF gty_s_cell,
    position      TYPE string,
    value         TYPE string,
    index         TYPE i,
    style         TYPE i,
    sharedstring  TYPE string,
    output_colnum TYPE i,
  END OF gty_s_cell .
TYPES:
  gty_t_cells TYPE STANDARD TABLE OF gty_s_cell WITH NON-UNIQUE KEY position INITIAL SIZE 1 .

TYPES: BEGIN OF gty_s_col,
         position    TYPE i,
         customwidth TYPE i,
         style       TYPE i,
         width       TYPE string,
         max         TYPE i,
         min         TYPE i,
         hidden      TYPE i,
         bestFit     TYPE i,
       END OF gty_s_col.

TYPES:
  BEGIN OF gty_s_row,
    spans     TYPE string,
    position  TYPE i,
*          outlinelevel type i,
*          hidden       type char1,
*          height       type i,
*          rowstyle     type char1,
*          range        type boolean,
    cells_tab TYPE gty_t_cells,
  END OF gty_s_row .

TYPES:
 gty_t_cols TYPE STANDARD TABLE OF gty_s_col WITH NON-UNIQUE KEY position INITIAL SIZE 1 .
TYPES:
 gty_t_rows TYPE STANDARD TABLE OF gty_s_row WITH NON-UNIQUE KEY position INITIAL SIZE 1 .
TYPES:
  BEGIN OF gty_s_sheet,
    dim       TYPE string,
    cols_tab  TYPE gty_t_cols,
    rows_tab  TYPE gty_t_rows,
    table_rid TYPE string,
  END OF gty_s_sheet .


TYPES:
  BEGIN OF gty_s_style_numfmt,
    id   TYPE i,
    code TYPE string,
  END OF gty_s_style_numfmt .
TYPES:
  BEGIN OF gty_s_style_cellxf,
    index    TYPE i,
    numfmtid TYPE i,
*      fillid TYPE i,
*      borderid TYPE i,
*      is_string TYPE i,
    indent   TYPE i,
    xfid     TYPE i,
    wrap     TYPE i,
    key      TYPE string,
  END OF gty_s_style_cellxf .
TYPES:
  BEGIN OF gty_s_style_cellxfs,
    index             TYPE i,
    numfmtid          TYPE i,
    fontid            TYPE i,
    fillid            TYPE i,
    borderid          TYPE i,
*      is_string TYPE i,
    indent            TYPE i,
    xfid              TYPE i,
    applyfont         TYPE i,
    applyfill         TYPE i,
    applyborder       TYPE i,
    applynumberformat TYPE i,
    wrap              TYPE i,
    key               TYPE string,
  END OF gty_s_style_cellxfs .
TYPES:
 gty_t_cellxfs TYPE STANDARD TABLE OF gty_s_style_cellxfs WITH NON-UNIQUE KEY index INITIAL SIZE 1 .




CLASS lcl_xlsx_sheet DEFINITION DEFERRED.
TYPES:
  BEGIN OF gty_s_sheet_obj,
    rid       TYPE string,
    sheet_obj TYPE REF TO lcl_xlsx_sheet,
  END OF gty_s_sheet_obj .
TYPES: gty_t_sheet_obj TYPE STANDARD TABLE OF gty_s_sheet_obj WITH NON-UNIQUE KEY rid INITIAL SIZE 1.




CLASS lcl_xlsx_style DEFINITION.
  PUBLIC SECTION.
    DATA: mv_has_change TYPE abap_bool.
    METHODS: init_create RAISING cx_dynamic_check.
    METHODS: init_with_xml IMPORTING iv_xml TYPE xstring
                           RAISING   cx_dynamic_check.
    METHODS: to_xml RETURNING VALUE(rv_xml) TYPE xstring
                    RAISING   cx_dynamic_check,
      get_default_date_style
        RETURNING
          VALUE(rv_style) TYPE gty_s_cell-style,
      get_default_time_style
        RETURNING
          VALUE(rv_style) TYPE gty_s_cell-style,
      get_default_tstmp_style
        RETURNING
          VALUE(rv_style) TYPE gty_s_cell-style.
  PRIVATE SECTION.
    DATA: mt_cellxfs         TYPE gty_t_cellxfs,
          mv_style_date      TYPE i,
          mv_style_time      TYPE i,
          mv_style_timestamp TYPE i,
          mv_base_xml        TYPE xstring.
    METHODS adapt_default_styles.
    METHODS get_cellxfs_xml_string
      RETURNING
        VALUE(rv_cellxfs) TYPE string
      RAISING
        cx_dynamic_check.
    METHODS prepare_styles_xml
      IMPORTING
        iv_cellxfs_count     TYPE i
      RETURNING
        VALUE(rv_styles_xml) TYPE string
      RAISING
        cx_dynamic_check.
ENDCLASS.




CLASS lcl_xlsx_doc DEFINITION.
  PUBLIC SECTION.
    INTERFACES zif_mpa_xlsx_doc.
    DATA: mo_xlsx_document TYPE REF TO cl_xlsx_document.
    DATA: mo_workbookpart  TYPE REF TO cl_xlsx_workbookpart.
    DATA: mo_string_util TYPE REF TO zcl_xlsx_string_util.
    DATA: mo_xlsx_style TYPE REF TO lcl_xlsx_style.
    METHODS constructor.
    METHODS initialize RAISING cx_openxml_not_found
                               cx_openxml_format
                               cx_openxml_not_allowed
                               cx_dynamic_check.
  PRIVATE SECTION.
    DATA ms_workbook_data TYPE gty_s_workbook_data.
    DATA mv_last_sheet_id TYPE i.
    DATA mv_workbook_xml TYPE xstring.
    DATA: mt_open_sheets TYPE gty_t_sheet_obj,
          mo_stylespart  TYPE REF TO cl_xlsx_stylespart.
    METHODS initialize_shared_strings RAISING cx_dynamic_check.
    METHODS initialize_workbook RAISING cx_openxml_not_found
                                        cx_openxml_format
                                        cx_dynamic_check.
    METHODS finalize_shared_strings RAISING cx_openxml_format cx_dynamic_check.

    METHODS finalize_workbook RAISING cx_dynamic_check.

    METHODS update_wb_xml_after_add_sheet
      IMPORTING
        is_sheet_info TYPE zcl_mpa_xlsx=>gty_s_sheet_info
      RAISING
        cx_dynamic_check.

    METHODS create_info_for_new_sheet
      IMPORTING
                iv_sheet_name        TYPE string
                io_worksheet_part    TYPE REF TO cl_xlsx_worksheetpart
      RETURNING VALUE(rs_sheet_info) TYPE zcl_mpa_xlsx=>gty_s_sheet_info
      RAISING   cx_dynamic_check.

    METHODS get_sheet_by_rid
      IMPORTING iv_rid          TYPE string
      RETURNING VALUE(ro_sheet) TYPE REF TO zif_mpa_xlsx_sheet
      RAISING   cx_openxml_format
                cx_openxml_not_found
                cx_dynamic_check.

    METHODS open_worksheet_part
      IMPORTING
        iv_rid            TYPE string
        io_worksheet_part TYPE REF TO cl_xlsx_worksheetpart
      RETURNING
        VALUE(ro_sheet)   TYPE REF TO zif_mpa_xlsx_sheet
      RAISING
        cx_dynamic_check.

    METHODS finalize_open_sheets RAISING cx_openxml_format cx_openxml_not_found cx_dynamic_check.

    METHODS initialize_styles
      RAISING cx_openxml_not_allowed cx_dynamic_check.

    METHODS finalize_styles RAISING cx_openxml_not_allowed cx_dynamic_check.

ENDCLASS.


CLASS lcl_xlsx_sheet DEFINITION.
  PUBLIC SECTION.
    INTERFACES zif_mpa_xlsx_sheet.
    DATA: mv_has_changes TYPE abap_bool VALUE abap_false.
    DATA: mo_worksheet_part TYPE REF TO cl_xlsx_worksheetpart.
    DATA: ms_sheet_data TYPE gty_s_sheet.
    DATA: mo_xlsx_doc TYPE REF TO lcl_xlsx_doc.
    METHODS: constructor IMPORTING io_xlsx_doc       TYPE REF TO lcl_xlsx_doc
                                   io_worksheet_part TYPE REF TO cl_xlsx_worksheetpart
                         RAISING   cx_dynamic_check,
      serialize
        RETURNING
          VALUE(rv_xml) TYPE xstring
        RAISING
          cx_dynamic_check.
    PRIVATE SECTION.    CONSTANTS gc_shared_string_indicator TYPE string VALUE 's' ##NO_TEXT.


    METHODS get_cell IMPORTING iv_row_number    TYPE i
                               iv_column_number TYPE i
                     RETURNING VALUE(rr_s_cell) TYPE REF TO gty_s_cell.
    METHODS add_cell IMPORTING iv_row_number    TYPE i
                               iv_column_number TYPE i
                     RETURNING VALUE(rr_s_cell) TYPE REF TO gty_s_cell.
    METHODS get_row IMPORTING iv_row_number   TYPE i
                    RETURNING VALUE(rr_s_row) TYPE REF TO gty_s_row.
    METHODS add_row IMPORTING iv_row_number   TYPE i
                    RETURNING VALUE(rr_s_row) TYPE REF TO gty_s_row.
    METHODS fill_cell
      IMPORTING
        iv_value           TYPE any
        iv_input_type      TYPE zif_mpa_xlsx_sheet=>celltype
        iv_treat_as_string TYPE abap_bool
      CHANGING
        cs_cell            TYPE gty_s_cell.
    METHODS do_input_conversion
      IMPORTING
        iv_value           TYPE any
        iv_input_type      TYPE zif_mpa_xlsx_sheet=>celltype
      EXPORTING
        ev_converted_value TYPE string.
    METHODS read_sheet RAISING cx_dynamic_check.
    METHODS determine_sheet_dimension
      RETURNING
        VALUE(rv_dimension) TYPE gty_s_sheet-dim.
    METHODS cleanup_sheet_data.
    METHODS get_output_type_for_input_type
      IMPORTING
        iv_input_type         TYPE zif_mpa_xlsx_sheet=>celltype
      RETURNING
        VALUE(rv_output_type) TYPE zif_mpa_xlsx_sheet=>cell_value_type.
    METHODS remove_cell
      IMPORTING
        iv_row_number    TYPE i
        iv_column_number TYPE i.
    METHODS get_style_for_input_type
      IMPORTING
        iv_input_type    TYPE zif_mpa_xlsx_sheet=>celltype
      RETURNING
        VALUE(rv_result) TYPE gty_s_cell-style.

ENDCLASS.


CLASS lcl_xlsx_doc IMPLEMENTATION.


  METHOD initialize.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Initialize the Workbook
    initialize_workbook( ).
*   Initialize the shared Strings
    initialize_shared_strings( ).
*   Initialize the Styles
    initialize_styles( ).
  ENDMETHOD.


  METHOD zif_mpa_xlsx_doc~save.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Finalize the open Worksheets
    me->finalize_open_sheets( ).
*   Finalize the shared strings
    me->finalize_shared_strings( ).
*   Finalize the styles
    me->finalize_styles( ).
*   Finalize the workbook
    me->finalize_workbook( ).
*   Then return the document
    rv_file_data = mo_xlsx_document->get_package_data( ).

  ENDMETHOD.


  METHOD initialize_shared_strings.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lo_sharedstrings     TYPE REF TO cl_xlsx_sharedstringspart,
          lv_sharedstrings_xml TYPE xstring.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    " Create String helper
    CREATE OBJECT mo_string_util.

    " A document may not have a shared strings part which is perfectly ok.
    " So when retrieving the shared strings part an exception could occur that we have to ignore.
    TRY.
        lo_sharedstrings = mo_workbookpart->get_sharedstringspart( ).

        lv_sharedstrings_xml = lo_sharedstrings->get_data( ).
        IF ( lv_sharedstrings_xml IS NOT INITIAL ).
          " Read the strings from the shared strings part and
          " create a lookup table for the indexes of the strings
          mo_string_util->init_from_xml( lv_sharedstrings_xml ).
          FREE lv_sharedstrings_xml.
        ENDIF.

      CATCH cx_openxml_format cx_openxml_not_found.
        " Ignore exceptions because it is ok if the document does not contain a shared strings part.
    ENDTRY.
  ENDMETHOD.



  METHOD finalize_shared_strings.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lo_sharedstrings     TYPE REF TO cl_xlsx_sharedstringspart,
          lv_sharedstrings_xml TYPE xstring.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*     If there are no Strings, we have nothing to do
    IF mo_string_util->has_strings( ) = abap_false.
      RETURN.
    ENDIF.

    " Update the shared strings since we may have added some
    lv_sharedstrings_xml = mo_string_util->to_xml( ).

    " Try to get the current Shared String part from the workbook.
    TRY.
        lo_sharedstrings = mo_workbookpart->get_sharedstringspart( ).
      CATCH cx_openxml_not_found
            cx_openxml_format.
*           Its perfectly ok, if the document did not have a sharedstring part, we
*           will create one then.
    ENDTRY.

    " If there is no shared string part yet, create a new one
    IF ( lo_sharedstrings IS INITIAL ).
      TRY.
          lo_sharedstrings = mo_workbookpart->add_sharedstringspart( ).
        CATCH cx_openxml_not_allowed.
      ENDTRY.
    ENDIF.

    lo_sharedstrings->feed_data( lv_sharedstrings_xml ).
    FREE lv_sharedstrings_xml.

  ENDMETHOD.


  METHOD initialize_workbook.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA lr_s_sheet_id TYPE REF TO zcl_mpa_xlsx=>gty_s_sheet_info.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

*   Get the Workbook
    mo_workbookpart = mo_xlsx_document->get_workbookpart( ).
    mv_workbook_xml = mo_workbookpart->get_data( ).
*   Get the names and ids of all available sheets and the number of the currently active sheet
    CALL TRANSFORMATION xl_mpa_get_sheet_names
      SOURCE XML mv_workbook_xml RESULT workbook_data = ms_workbook_data.

    " Determine the current maximum sheet ID.
    LOOP AT ms_workbook_data-sheet_ids_htab REFERENCE INTO lr_s_sheet_id.
      IF ( mv_last_sheet_id < lr_s_sheet_id->sheet_id ).
        mv_last_sheet_id = lr_s_sheet_id->sheet_id.
      ENDIF.
    ENDLOOP.


  ENDMETHOD.


  METHOD finalize_workbook.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    " Update the workbook since we may have created new sheets
    mo_workbookpart->feed_data( mv_workbook_xml ).
  ENDMETHOD.

  METHOD zif_mpa_xlsx_doc~get_sheets.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    rt_sheet_info = ms_workbook_data-sheet_ids_htab.
  ENDMETHOD.



  METHOD get_sheet_by_rid.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lo_worksheet_part TYPE REF TO cl_xlsx_worksheetpart.
    DATA: ls_open_sheet TYPE gty_s_sheet_obj.

*   If the sheet is already open, return it
    READ TABLE mt_open_sheets REFERENCE INTO DATA(lr_s_open_sheet)
      WITH TABLE KEY rid = iv_rid.
    IF sy-subrc EQ 0.
      ro_sheet = lr_s_open_sheet->sheet_obj.
    ENDIF.

*   Otherwise find the worksheet part
    lo_worksheet_part ?= mo_workbookpart->get_part_by_id( iv_id = iv_rid ).

*   And open it
    ro_sheet = open_worksheet_part(
          iv_rid            = iv_rid
          io_worksheet_part = lo_worksheet_part ).

  ENDMETHOD.

  METHOD zif_mpa_xlsx_doc~get_sheet_by_id.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    " Determine the sheet that belongs to the data selection
    READ TABLE ms_workbook_data-sheet_ids_htab REFERENCE INTO DATA(lr_s_sheet_info)
                WITH KEY sheet_id = iv_sheet_id.
    IF ( sy-subrc NE 0 ).
      RAISE EXCEPTION TYPE cx_openxml_not_found.
    ENDIF.
    ro_sheet = get_sheet_by_rid( lr_s_sheet_info->rid ).
  ENDMETHOD.


  METHOD zif_mpa_xlsx_doc~get_sheet_by_name.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*     Try to find the sheet with the given name
    READ TABLE ms_workbook_data-sheet_ids_htab REFERENCE INTO DATA(lr_s_sheet_info)
                WITH KEY name = iv_sheet_name.
    IF sy-subrc NE 0.
      RAISE EXCEPTION TYPE cx_openxml_not_found.
    ENDIF.
    ro_sheet = get_sheet_by_rid( lr_s_sheet_info->rid  ).
  ENDMETHOD.


  METHOD zif_mpa_xlsx_doc~add_new_sheet.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: ls_sheet_info LIKE LINE OF ms_workbook_data-sheet_ids_htab.
    DATA: lo_worksheet_part TYPE REF TO cl_xlsx_worksheetpart,
          lo_sheet          TYPE REF TO lcl_xlsx_sheet.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

*   Create a new Worksheet part
    lo_worksheet_part = mo_workbookpart->add_worksheetpart( ).

*   Create Sheet info for the new Sheet
    ls_sheet_info = create_info_for_new_sheet( iv_sheet_name     = iv_sheet_name
                                               io_worksheet_part = lo_worksheet_part ).

*   update the Workbook XML
    update_wb_xml_after_add_sheet( is_sheet_info = ls_sheet_info ).

*   Then open the new created Sheet
    lo_sheet ?= open_worksheet_part(
               iv_rid            = ls_sheet_info-rid
               io_worksheet_part = lo_worksheet_part
           ).
*   And flag it to say, that it has changes
    lo_sheet->mv_has_changes = abap_true.

    ro_sheet = lo_sheet.

  ENDMETHOD.


  METHOD update_wb_xml_after_add_sheet.

*   Insert the new sheet into the workbook XML
    IF ( mv_workbook_xml IS INITIAL ).
*     We do not have a workbook XML so we have to create one from scratch
      CALL TRANSFORMATION xl_mpa_create_workbook
                   SOURCE param = is_sheet_info
               RESULT XML mv_workbook_xml.
    ELSE.
*     We just have to insert a new sheet into an existing workbook XML

*     We insert the new sheet as first sheet. Thus the currently active sheet
*     moves one position to the right, i.e. the number of the active sheet
*     must be increased by 1.
      ADD 1 TO ms_workbook_data-active_sheet.

      CALL TRANSFORMATION xl_mpa_insert_sheet
               PARAMETERS active_sheet = ms_workbook_data-active_sheet
                          sheet_name   = is_sheet_info-name
                          sheet_id     = is_sheet_info-sheet_id
                          sheet_rid    = is_sheet_info-rid
               SOURCE XML mv_workbook_xml
               RESULT XML mv_workbook_xml.
    ENDIF.

  ENDMETHOD.


  METHOD create_info_for_new_sheet.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

*   Fill the Sheet Info
    rs_sheet_info-name = iv_sheet_name.
    TRY.
        rs_sheet_info-rid  = mo_workbookpart->get_id_for_part( io_worksheet_part ).
      CATCH cx_openxml_not_found.
*       This should never happen as we just created the sheet
        ASSERT 1 = 2.
    ENDTRY.

*   Generate a new sheet ID
    ADD 1 TO mv_last_sheet_id.
    rs_sheet_info-sheet_id = mv_last_sheet_id.

*   Add the info to the table
    INSERT rs_sheet_info INTO TABLE ms_workbook_data-sheet_ids_htab.

  ENDMETHOD.


  METHOD open_worksheet_part.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA ls_open_sheet TYPE gty_s_sheet_obj.

***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

*   Create a new Sheet Object
    CREATE OBJECT ls_open_sheet-sheet_obj TYPE lcl_xlsx_sheet
      EXPORTING
        io_xlsx_doc       = me
        io_worksheet_part = io_worksheet_part.
    ls_open_sheet-rid = iv_rid.

*   Insert the Sheet object to the open sheets
    APPEND ls_open_sheet TO mt_open_sheets.

*   And return the created object
    ro_sheet = ls_open_sheet-sheet_obj.

  ENDMETHOD.


  METHOD finalize_open_sheets.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lo_sheet TYPE REF TO lcl_xlsx_sheet.
    DATA: lv_xml TYPE xstring.
    DATA: lo_worksheetpart TYPE REF TO cl_xlsx_worksheetpart.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    LOOP AT mt_open_sheets REFERENCE INTO DATA(lr_s_open_sheet).
      lo_sheet = lr_s_open_sheet->sheet_obj.
      IF lo_sheet->mv_has_changes = abap_true.
*       Serialize the data to XML
        lv_xml = lo_sheet->serialize( ).
*       Then get the corresponding worksheetpart and feed the XML into it
        lo_worksheetpart ?= mo_workbookpart->get_part_by_id( iv_id = lr_s_open_sheet->rid ).
        lo_worksheetpart->feed_data( iv_data = lv_xml ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD initialize_styles.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Try to get the stylespart
    TRY.
        mo_stylespart = mo_workbookpart->get_stylespart( ).
        mo_xlsx_style->init_with_xml( iv_xml = mo_stylespart->get_data( ) ).

      CATCH cx_openxml_format cx_openxml_not_found.  "
*       There is no styles part yet, create one
        mo_xlsx_style->init_create( ).
    ENDTRY.

  ENDMETHOD.



  METHOD finalize_styles.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lv_styles_xml TYPE xstring.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    IF mo_xlsx_style->mv_has_change EQ abap_true.
      IF mo_stylespart IS INITIAL.
        mo_stylespart = mo_workbookpart->add_stylespart( ).
      ENDIF.
      lv_styles_xml = mo_xlsx_style->to_xml( ).
      mo_stylespart->feed_data( iv_data = lv_styles_xml ).
    ENDIF.
  ENDMETHOD.


  METHOD constructor.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    CREATE OBJECT mo_xlsx_style.
  ENDMETHOD.

ENDCLASS.






CLASS lcl_xlsx_sheet IMPLEMENTATION.

  METHOD constructor.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    mo_xlsx_doc = io_xlsx_doc.
    mo_worksheet_part = io_worksheet_part.
    CLEAR ms_sheet_data.
    read_sheet( ).
  ENDMETHOD.

  METHOD get_row.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Get the Row at the given position
    CLEAR rr_s_row.
    READ TABLE ms_sheet_data-rows_tab WITH KEY position = iv_row_number REFERENCE INTO rr_s_row.
    "READ TABLE ms_sheet_data-rows_tab INDEX iv_row_number REFERENCE INTO rr_s_row.
  ENDMETHOD.

  METHOD add_row.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Create a Row and add it to the sheet data
    CREATE DATA rr_s_row.
    rr_s_row->position = iv_row_number.
    APPEND rr_s_row->* TO ms_sheet_data-rows_tab REFERENCE INTO rr_s_row.
  ENDMETHOD.


  METHOD get_cell.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lr_s_row TYPE REF TO gty_s_row.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    CLEAR rr_s_cell.
*   Get the row
    lr_s_row = get_row( iv_row_number ).
*   If we found a row, try to find the cell
    IF lr_s_row IS BOUND.
      READ TABLE lr_s_row->cells_tab WITH KEY output_colnum = iv_column_number REFERENCE INTO rr_s_cell.
    ENDIF.
  ENDMETHOD.

  METHOD add_cell.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lr_s_row TYPE REF TO gty_s_row.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   First get or create the row
    lr_s_row = get_row( iv_row_number ).
    IF lr_s_row IS NOT BOUND.
      lr_s_row = add_row( iv_row_number ).
    ENDIF.
*   Then create the cell
    CREATE DATA rr_s_cell.
    rr_s_cell->output_colnum = iv_column_number.
*   Get the Excel cell name like 'A1'
    rr_s_cell->position = zcl_xlsx_helper=>get_cell_position( iv_row    = iv_row_number
                                                                       iv_column = iv_column_number ).
    APPEND rr_s_cell->* TO lr_s_row->cells_tab REFERENCE INTO rr_s_cell.
  ENDMETHOD.

  METHOD remove_cell.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lr_s_row TYPE REF TO gty_s_row.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   First get or create the row
    lr_s_row = get_row( iv_row_number ).
    IF lr_s_row IS NOT BOUND.
*     This row does not even exist, return...
      RETURN.
    ENDIF.
*   Now delete the cell from the row
    DELETE lr_s_row->cells_tab WHERE output_colnum = iv_column_number.
*   Now check, if the row is now also empty
    IF lines( lr_s_row->cells_tab ) < 1.
*     No cell is left in the row, remove the row
      DELETE ms_sheet_data-rows_tab WHERE position = iv_row_number.
    ENDIF.
  ENDMETHOD.



  METHOD zif_mpa_xlsx_sheet~set_cell_content.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lr_s_cell          TYPE REF TO gty_s_cell,
          lv_treat_as_string TYPE abap_bool,
          lv_converted_value TYPE string.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    IF iv_input_type EQ zif_mpa_xlsx_sheet=>gc_celltype-charlike.
*     If we have a character based type we can directly take it over
      lv_treat_as_string = abap_true.
      lv_converted_value = iv_value.
    ELSE.
*     Otherwise we try to do a conversion if needed
      do_input_conversion(
      EXPORTING iv_value           = iv_value
                iv_input_type      = iv_input_type
      IMPORTING ev_converted_value = lv_converted_value ).

*     Check if the value should be treated like a string in the output
      IF iv_force_string EQ abap_true.
        lv_treat_as_string = abap_true.
      ENDIF.
    ENDIF.

*   First get or add the cell
    lr_s_cell = get_cell( iv_row_number = iv_row     iv_column_number = iv_column ).
    IF lr_s_cell IS NOT BOUND.
      IF lv_converted_value IS NOT INITIAL.
*       There is a new value to be entered into the cell, add a new cell
        lr_s_cell = add_cell( iv_row_number = iv_row   iv_column_number = iv_column ).
      ELSE.
        RETURN. " There was nothing to add to this cell - so we can just leave it emtpy
      ENDIF.
    ENDIF.

    IF lv_converted_value IS INITIAL.
*     The cell is cleared, we should remove the cell from the sheet
      remove_cell( iv_row_number = iv_row     iv_column_number = iv_column ).
      RETURN.
    ENDIF.

*   Now fill the cell based on the type
    fill_cell( EXPORTING iv_value           = lv_converted_value
                         iv_input_type      = iv_input_type
                         iv_treat_as_string = lv_treat_as_string
                CHANGING cs_cell            = lr_s_cell->* ).

*   Signal that we have changes here
    mv_has_changes = abap_true.

  ENDMETHOD.

  METHOD zif_mpa_xlsx_sheet~has_cell_content.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lr_s_cell TYPE REF TO gty_s_cell.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    lr_s_cell = get_cell(
                iv_row_number    = iv_row
                iv_column_number = iv_column
            ).
    IF lr_s_cell IS BOUND.
      rv_has_content = abap_true.
    ELSE.
      rv_has_content = abap_false.
    ENDIF.

  ENDMETHOD.

  METHOD zif_mpa_xlsx_sheet~get_cell_content.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lr_s_cell TYPE REF TO gty_s_cell.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    CLEAR rv_content.

*   Find the cell at the given position
    lr_s_cell = get_cell(
                iv_row_number    = iv_row
                iv_column_number = iv_column
            ).

    IF lr_s_cell IS NOT BOUND.
*     Return an empty result if the cell does not exist
      RETURN.
    ENDIF.
    IF lr_s_cell->sharedstring EQ gc_shared_string_indicator.
*     For Sharedstrings we have to get the content from the String Util
      rv_content = mo_xlsx_doc->mo_string_util->get_string_at_index( lr_s_cell->index ).
    ELSE.
*     Otherwise we can directly return the value in the field
      rv_content = lr_s_cell->value.
    ENDIF.
  ENDMETHOD.

  METHOD zif_mpa_xlsx_sheet~get_last_row_number.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Start with -1 To signal if there are not any rows in the table
    rv_row = -1.
*   Then try to find the highest row number
    LOOP AT ms_sheet_data-rows_tab REFERENCE INTO DATA(lr_s_row).
      IF lr_s_row->position GT rv_row.
        rv_row = lr_s_row->position.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD zif_mpa_xlsx_sheet~get_last_column_number_in_row.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Start with -1 To signal if there are not any columns in the row
    rv_column = -1.
*   Try to find the row
    DATA(lr_s_row) = get_row( iv_row ).
    IF lr_s_row IS BOUND.
*     If we found the row, search the highest column number in it
      LOOP AT lr_s_row->cells_tab REFERENCE INTO DATA(lr_s_cell).
        IF lr_s_cell->output_colnum GT rv_column.
          rv_column = lr_s_cell->output_colnum.
        ENDIF.
      ENDLOOP.
    ENDIF.
  ENDMETHOD.

  METHOD zif_mpa_xlsx_sheet~change_sheet_name.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
* TODO: Implement
*   Idea: tell the document/worksheet to update
  ENDMETHOD.


  METHOD fill_cell.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    IF iv_treat_as_string EQ abap_true.
*     Flag the cell as having a shared string
      cs_cell-sharedstring = zif_mpa_xlsx_sheet=>gc_cell_value_type-shared_string.
*     Clear the value
      CLEAR cs_cell-value.
*     And set the index of the shared string
      cs_cell-index = mo_xlsx_doc->mo_string_util->get_index_of_string( iv_value ).
    ELSE.
*     Otherwise we can directly set the value
      CLEAR cs_cell-sharedstring.
      CLEAR cs_cell-index.
      cs_cell-sharedstring = get_output_type_for_input_type( iv_input_type ).
      cs_cell-style        = get_style_for_input_type( iv_input_type ).
      cs_cell-value        = iv_value.
    ENDIF.

  ENDMETHOD.


  METHOD do_input_conversion.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lv_date                   TYPE d.
    DATA: lv_time                   TYPE t.
    DATA: lv_timestamp              TYPE p.
    DATA: lv_converted_chars TYPE c LENGTH 1000.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    CASE iv_input_type.
      WHEN zif_mpa_xlsx_sheet=>gc_celltype-date.
        " convert the importing data to date
        lv_date = iv_value.
        " call date conversion
        ev_converted_value = zcl_xlsx_helper=>convert_date_to_xlsx_date( iv_date = lv_date ).

      WHEN zif_mpa_xlsx_sheet=>gc_celltype-time.
        " convert the importing data to time
        lv_time = iv_value.
        " call time conversion
        ev_converted_value = zcl_xlsx_helper=>convert_time_to_xlsx_time( iv_time = lv_time ).

      WHEN zif_mpa_xlsx_sheet=>gc_celltype-timestamp.
        " convert the importing data to timestamp
        lv_timestamp = iv_value.
        " call timestamp conversion
        ev_converted_value = zcl_xlsx_helper=>convert_timestamp_to_xlsx( iv_timestamp = lv_timestamp ).


      WHEN zif_mpa_xlsx_sheet=>gc_celltype-numericchar.
        " For NUMC values we have to check if they contain only 0s. If yes
        " we have to return just a single 0. Otherwise Microsoft Excel does
        " not recognize the value as numerical and displays an error message.
        IF ( iv_value CO '0' ).
          ev_converted_value = '0'.
        ELSE.
          ev_converted_value = iv_value.
        ENDIF.

      WHEN OTHERS.
        " per default do the output conversion
        WRITE iv_value TO lv_converted_chars.
        ev_converted_value = lv_converted_chars.

    ENDCASE.

  ENDMETHOD.


  METHOD read_sheet.
***************************************************************************
* DATA DEFINITION
***************************************************************************
*  Types for reading the Excel Worksheet
    TYPES:
      BEGIN OF lty_s_cell,
        refname   TYPE string,
        celltype  TYPE c LENGTH 1,
        cellstyle TYPE i,
        cellvalue TYPE string,
      END OF lty_s_cell .
    TYPES:
      BEGIN OF lty_s_row,
        index TYPE i,
        cells TYPE STANDARD TABLE OF lty_s_cell WITH DEFAULT KEY,
      END OF lty_s_row .
    TYPES:
      lty_t_row  TYPE STANDARD TABLE OF lty_s_row .

    DATA: lv_xml            TYPE xstring.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
    lv_xml = mo_worksheet_part->get_data( ).

*   Read the Sheet
    CALL TRANSFORMATION xl_mpa_read_sheet
        SOURCE XML lv_xml
        RESULT rows = ms_sheet_data-rows_tab
               cols = ms_sheet_data-cols_tab .

*   Now analyse the rows and correct them where needed
    LOOP AT ms_sheet_data-rows_tab REFERENCE INTO DATA(lr_s_row).
      LOOP AT lr_s_row->cells_tab  REFERENCE INTO DATA(lr_s_cell).
*       Correct the cell if we have a shared string
        IF lr_s_cell->sharedstring EQ gc_shared_string_indicator.
          lr_s_cell->index = lr_s_cell->value.
          CLEAR lr_s_cell->value.
        ENDIF.
*       Get the column index for the position
        lr_s_cell->output_colnum = zcl_xlsx_helper=>get_cell_column( lr_s_cell->position ).
      ENDLOOP.
    ENDLOOP.

  ENDMETHOD.


  METHOD serialize.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   First sort the sheet data
    cleanup_sheet_data( ).
*   Then determine the dimensions
    ms_sheet_data-dim = determine_sheet_dimension( ).
    ms_sheet_data-cols_tab = VALUE #( (  position = 1
                                         customwidth = 1
*                                         style
                                         width = 100
                                         max = 1
                                         min = 1
*                                         hidden  type i
                                         bestfit = 1
     ) ).
*   Create a new sheet XML with the data
    CALL TRANSFORMATION xl_mpa_create_sheet
               SOURCE param = ms_sheet_data
           RESULT XML rv_xml.
  ENDMETHOD.


  METHOD determine_sheet_dimension.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lv_startcell TYPE string.
    DATA: lv_endcolumn TYPE i VALUE 1.
    DATA: lv_endrow    TYPE i VALUE 1.
    DATA: lv_endcell   TYPE string.
    DATA: lv_num_rows TYPE i.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   We start always in the first cell
    lv_startcell = 'A1'.

*   Now we can loop the rows to find the highest column index
    LOOP AT ms_sheet_data-rows_tab REFERENCE INTO DATA(lr_s_sheet_row).
*     Check if the Row position is higher than the last found one
      IF lr_s_sheet_row->position GT lv_endrow.
        lv_endrow = lr_s_sheet_row->position.
      ENDIF.
*     As we have cleaned up before, the row is not empty
*     and the cells in the row are already sorted, so get the last cell in the row
      READ TABLE lr_s_sheet_row->cells_tab INDEX lines( lr_s_sheet_row->cells_tab )
              REFERENCE INTO DATA(lr_s_cell).
*     Check if the column index of the last cell is higher than the current found one
      IF lr_s_cell->output_colnum GT lv_endcolumn.
        lv_endcolumn = lr_s_cell->output_colnum.
      ENDIF.
    ENDLOOP.

*   Then determine the Position in Excel language
    lv_endcell = zcl_xlsx_helper=>get_cell_position(
                 iv_row      = lv_endrow
                 iv_column   = lv_endcolumn
             ).

*   And return the range
    CONCATENATE lv_startcell lv_endcell INTO rv_dimension SEPARATED BY ':'.

  ENDMETHOD.


  METHOD cleanup_sheet_data.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   First delete empty rows
    DELETE ms_sheet_data-rows_tab WHERE cells_tab IS INITIAL.
*   Then sort the rows
    SORT ms_sheet_data-rows_tab BY position ASCENDING.
*   Then loop over the rows and sort the cells in the row
    LOOP AT ms_sheet_data-rows_tab REFERENCE INTO DATA(lr_s_sheet_row).
      SORT lr_s_sheet_row->cells_tab BY output_colnum ASCENDING.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_output_type_for_input_type.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Start with an undefined type
    rv_output_type = zif_mpa_xlsx_sheet=>gc_cell_value_type-undefined.
*   Currently it seems, that Excel only takes one type: 's' for shared strings
*   The others (date, time, etc.) are handled by using styles for the cells

*    CASE iv_input_type.
*      WHEN zif_mpa_xlsx_sheet=>gc_celltype-date.
*        rv_output_type = zif_mpa_xlsx_sheet=>gc_cell_value_type-date.
*      WHEN zif_mpa_xlsx_sheet=>gc_celltype-float
*         OR zif_mpa_xlsx_sheet=>gc_celltype-numericchar
*        OR zif_mpa_xlsx_sheet=>gc_celltype-integer "Excel does not seem to do this...
*            .
*        rv_output_type = zif_mpa_xlsx_sheet=>gc_cell_value_type-number.
*    ENDCASE.

  ENDMETHOD.

  METHOD get_style_for_input_type.
*   We start with an empty result - so no style information is applied
    CLEAR rv_result.

    CASE iv_input_type.
      WHEN zif_mpa_xlsx_sheet=>gc_celltype-date.
        rv_result = mo_xlsx_doc->mo_xlsx_style->get_default_date_style( ).
      WHEN zif_mpa_xlsx_sheet=>gc_celltype-time.
        rv_result = mo_xlsx_doc->mo_xlsx_style->get_default_time_style( ).
      WHEN zif_mpa_xlsx_sheet=>gc_celltype-timestamp.
        rv_result = mo_xlsx_doc->mo_xlsx_style->get_default_tstmp_style( ).
    ENDCASE.


  ENDMETHOD.

ENDCLASS.







CLASS lcl_xlsx_style IMPLEMENTATION.

  METHOD init_create.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lv_styles_xml TYPE xstring.
    DATA: lv_dummy      TYPE string.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

*   As we create a new Styles.xml, we always have changes
    mv_has_change = abap_true.

*   Add style
    CALL TRANSFORMATION xl_mpa_create_styles
        SOURCE param = lv_dummy
        RESULT XML lv_styles_xml.

*   Then init with our just created XML
    me->init_with_xml( lv_styles_xml ).
  ENDMETHOD.

  METHOD init_with_xml.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************
*   Get the Base XML
    mv_base_xml = iv_xml.

*   Get existing cellXfs
    CALL TRANSFORMATION XL_MPA_get_cellxfs
    SOURCE XML iv_xml
    RESULT cellxfs = mt_cellxfs.



*   Now find or add the default styles for date, time and timestamp
    me->adapt_default_styles( ).

  ENDMETHOD.




  METHOD adapt_default_styles.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    CONSTANTS: lc_numfmtid_date      TYPE i VALUE 14,
               lc_numfmtid_time      TYPE i VALUE 21,
               lc_numfmtid_timestamp TYPE i VALUE 22.
    DATA: lv_style_date_found  TYPE abap_bool VALUE abap_false,
          lv_style_time_found  TYPE abap_bool VALUE abap_false,
          lv_style_tstmp_found TYPE abap_bool VALUE abap_false,
          lv_max_cellxfs       TYPE i,
          ls_cellxfs           TYPE gty_s_style_cellxfs.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    CLEAR: mv_style_date, mv_style_time, mv_style_timestamp.

* Check if the default styles are already in the XML
    LOOP AT mt_cellxfs REFERENCE INTO DATA(lr_s_cellxfs).
      CASE lr_s_cellxfs->numfmtid.
        WHEN lc_numfmtid_date.
          IF lr_s_cellxfs->applyfont IS INITIAL AND
             lr_s_cellxfs->applyfill IS INITIAL AND
             lr_s_cellxfs->applyborder IS INITIAL AND
             lr_s_cellxfs->xfid IS INITIAL.
            " Set index for date. Index starts at 0.
            lv_style_date_found = abap_true.
            mv_style_date = sy-tabix - 1.
          ENDIF.
        WHEN lc_numfmtid_time.
          IF lr_s_cellxfs->applyfont IS INITIAL AND
             lr_s_cellxfs->applyfill IS INITIAL AND
             lr_s_cellxfs->applyborder IS INITIAL AND
             lr_s_cellxfs->xfid IS INITIAL.
            lv_style_time_found = abap_true.
            mv_style_time = sy-tabix - 1.
          ENDIF.
        WHEN lc_numfmtid_timestamp.
          IF lr_s_cellxfs->applyfont IS INITIAL AND
             lr_s_cellxfs->applyfill IS INITIAL AND
             lr_s_cellxfs->applyborder IS INITIAL AND
             lr_s_cellxfs->xfid IS INITIAL.
            lv_style_tstmp_found = abap_true.
            mv_style_timestamp = sy-tabix - 1.
          ENDIF.
      ENDCASE.
    ENDLOOP.

*   Add the default styles that are missing
    CLEAR ls_cellxfs.
    lv_max_cellxfs = lines( mt_cellxfs ) - 1.
    " Date
    IF lv_style_date_found = abap_false.
      " remember style index number
      lv_max_cellxfs = lv_max_cellxfs + 1.
      mv_style_date = lv_max_cellxfs.
      " numFmtId
      ls_cellxfs-numfmtid = lc_numfmtid_date.
      APPEND ls_cellxfs TO mt_cellxfs.
    ENDIF.
    " Time
    IF lv_style_time_found = abap_false.
      " remember style index number
      lv_max_cellxfs = lv_max_cellxfs + 1.
      mv_style_time = lv_max_cellxfs.
      " numFmtId
      ls_cellxfs-numfmtid = lc_numfmtid_time.
      APPEND ls_cellxfs TO mt_cellxfs.
    ENDIF.
    " Timestamp
    IF lv_style_tstmp_found = abap_false.
      " remember style index number
      lv_max_cellxfs = lv_max_cellxfs + 1.
      mv_style_timestamp =  lv_max_cellxfs.
      " numFmtId
      ls_cellxfs-numfmtid = lc_numfmtid_timestamp.
      APPEND ls_cellxfs TO mt_cellxfs.
    ENDIF.

  ENDMETHOD.



  METHOD to_xml.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lv_cellxfs    TYPE string.
    DATA: lv_styles_xml TYPE string.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

*   Get the Snippte for the cellXfs
    lv_cellxfs = get_cellxfs_xml_string( ).

*   Prepare the Output Styles XML with empty cellXfs part
    lv_styles_xml = prepare_styles_xml( lines( mt_cellxfs ) ).

*   Insert our cellXfs into the styles XML
    REPLACE FIRST OCCURRENCE OF 'XFS_PLACEHOLDER' IN lv_styles_xml WITH lv_cellxfs IN CHARACTER MODE.

*   Then tranform to xstring and return
    rv_xml = cl_openxml_helper=>string_to_xstring( lv_styles_xml ).

  ENDMETHOD.


  METHOD get_cellxfs_xml_string.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    TYPES:
      BEGIN OF lty_s_style_struct,
        t_numfmts     TYPE STANDARD TABLE OF gty_s_style_numfmt WITH KEY id INITIAL SIZE 1,
        t_cellxfs     TYPE STANDARD TABLE OF gty_s_style_cellxfs WITH KEY key indent xfid wrap INITIAL SIZE 1,
        numfmts_count TYPE i,
        cellxfs_count TYPE i,
      END OF  lty_s_style_struct .
    DATA: ls_style_struct TYPE lty_s_style_struct.
    DATA: lv_dummy TYPE string.
***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    " Fill structure for style xml transformation
    CLEAR ls_style_struct.
    ls_style_struct-cellxfs_count = lines( mt_cellxfs ).
    ls_style_struct-t_cellxfs = mt_cellxfs.

    " Transform cellxfs part to XML (xstring)
    CALL TRANSFORMATION xl_mpa_set_cellxfs
    SOURCE param = ls_style_struct
    RESULT XML rv_cellxfs.

* Strip the surrounding cellXfs Tags - we will re-introduce them later
    SPLIT rv_cellxfs AT '<cellXfs>' INTO lv_dummy rv_cellxfs.
    SPLIT rv_cellxfs AT '</cellXfs>' INTO rv_cellxfs lv_dummy.
    CLEAR lv_dummy.

  ENDMETHOD.


  METHOD prepare_styles_xml.
***************************************************************************
* DATA DEFINITION
***************************************************************************
    DATA: lv_styles_string TYPE string,
          lv_dummy         TYPE string,
          lv_delete        TYPE string,
          lv_styles_xml    TYPE xstring,
          lv_dummy2        TYPE string.

***************************************************************************
* FUNCTIONAL BODY
***************************************************************************

    " Convert template style xml to string
    lv_styles_string = cl_openxml_helper=>xstring_to_string( mv_base_xml ).

    " Remove existing cellXfs
    SPLIT lv_styles_string AT '<cellXfs count' INTO lv_dummy lv_delete.
    SPLIT lv_delete AT '</cellXfs>' INTO lv_delete lv_dummy2.
    CONCATENATE lv_dummy '<cellXfs />' lv_dummy2 INTO lv_styles_string.

    " Convert back to xml
    lv_styles_xml = cl_openxml_helper=>string_to_xstring( lv_styles_string ).

    " Add placeholder for cellXfs
    CALL TRANSFORMATION xl_mpa_add_cellxfs
      SOURCE XML lv_styles_xml
      RESULT XML lv_styles_xml
      PARAMETERS cellxfs_count = iv_cellxfs_count.

    " Convert back to string
    rv_styles_xml = cl_openxml_helper=>xstring_to_string( lv_styles_xml ).
  ENDMETHOD.


  METHOD get_default_date_style.
    rv_style = mv_style_date.
  ENDMETHOD.


  METHOD get_default_time_style.
    rv_style = mv_style_time.
  ENDMETHOD.


  METHOD get_default_tstmp_style.
    rv_style = mv_style_timestamp.
  ENDMETHOD.

ENDCLASS.
