CLASS zcl_xlsx_parse_util DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES if_mpa_xlsx_parse_util .

*    TYPES:
*      BEGIN OF gty_s_struct_properties,
*        field_name   TYPE  name_feld,
*        data_element TYPE rollname,
*        data_type    TYPE datatype_d,
*        length       TYPE ddleng,
*      END OF gty_s_struct_properties .
*    TYPES:
*      gty_t_struct_properties TYPE STANDARD TABLE OF gty_s_struct_properties .

    CONSTANTS:
      BEGIN OF gc_file_type,
        excel TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        csv   TYPE string VALUE 'application/vnd.ms-excel',
      END OF gc_file_type .
    CONSTANTS:
      BEGIN OF gc_exception_type,
        business_exception  TYPE char1 VALUE '1',
        technical_exception TYPE char1 VALUE '2',
      END OF gc_exception_type .
    CONSTANTS:
      "Separators
      gc_sep_semicolon TYPE c LENGTH 1 VALUE ';' ##NO_TEXT,
      gc_sep_comma     TYPE c LENGTH 1 VALUE ',' ##NO_TEXT,
      gc_sep_pipe      TYPE c LENGTH 1 VALUE '|' ##NO_TEXT,
      gc_sep_dot       TYPE c LENGTH 1 VALUE '.' ##NO_TEXT,
      gc_sep_tab       TYPE c LENGTH 1 VALUE cl_abap_char_utilities=>horizontal_tab ##NO_TEXT.
    CLASS-DATA gc_comment_symbol TYPE string VALUE '//' ##NO_TEXT.
    CLASS-DATA gv_template_type TYPE string .

    METHODS constructor .
    CLASS-METHODS get_instance
      RETURNING
        VALUE(ro_instance) TYPE REF TO if_mpa_xlsx_parse_util .

  PROTECTED SECTION.

  PRIVATE SECTION.

    TYPES:
      BEGIN OF gty_s_field_name_mapping,
        cell_name TYPE string,
        cell_posi TYPE string,
        stru_name TYPE string,
        value     TYPE string,
      END OF gty_s_field_name_mapping .
    TYPES:
      gty_t_field_name_mapping TYPE HASHED TABLE OF gty_s_field_name_mapping WITH UNIQUE KEY cell_name cell_posi .
    TYPES:
      BEGIN OF gty_s_comment_line,
        index      TYPE i,
        skip_lines TYPE i,
      END OF gty_s_comment_line .
    TYPES:
      gty_t_comment_line TYPE STANDARD TABLE OF gty_s_comment_line .
    TYPES:
      gty_t_dd04l     TYPE STANDARD TABLE OF dd04l .
    TYPES:
      BEGIN OF ty_supported_sep,
        separator TYPE c LENGTH 1,
      END OF ty_supported_sep .
    TYPES:BEGIN OF ty_s_sequence_no.
    TYPES sequence_no          TYPE string.
    TYPES line_number  TYPE string.
    TYPES END OF ty_s_sequence_no .
    TYPES:
      ty_t_sequence_no TYPE STANDARD TABLE OF ty_s_sequence_no .
    CONSTANTS gc_nr_object TYPE nrobj VALUE 'MPA_FILEID' ##NO_TEXT.
    CONSTANTS cs_empty_symbol TYPE string VALUE '##empty##' ##NO_TEXT.
    DATA mt_dd04l TYPE gty_t_dd04l .
    DATA mt_excel_rows TYPE mpa_t_index_value_pair .
    DATA mt_line_index TYPE mpa_t_excel_doc_index .
    DATA gt_comment_line TYPE gty_t_comment_line .
    DATA mo_lcl_xlsx_parse TYPE REF TO lif_mpa_xlsx_parse_util .
    DATA gt_fieldinfo TYPE dd_x031l_table .
    DATA: mt_supported_separators  TYPE TABLE OF ty_supported_sep WITH KEY separator .
    DATA mv_separator TYPE string .
    "! Object for interface if_mpa_xlsx_parse_util
    CLASS-DATA go_instance TYPE REF TO if_mpa_xlsx_parse_util .

    "! Parse the uploaded excel file
    METHODS parse_xlsx
      IMPORTING
        !ix_file     TYPE xstring
      EXPORTING
        !ev_mpa_type TYPE mpa_template_type
        !et_asset    TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data
        !et_message  TYPE bapirettab
      RAISING
        cx_openxml_not_found
        cx_openxml_format .
    "! Save the uploaded excel file data to the DB
    METHODS save_file_to_db
      IMPORTING
        !ix_file_content   TYPE xstring
        !iv_file_name      TYPE string
        !iv_mpa_file_type  TYPE mpa_template_type
      RETURNING
        VALUE(rs_messages) TYPE bapiret2 .
    "! Map the excel file data to the corresponding structure of scenario ( Create, Change, Transfer)
    METHODS map_excel_data
      IMPORTING
        !it_table          TYPE mpa_t_index_value_pair OPTIONAL
        !is_index          TYPE mpa_s_excel_doc_index
      EXPORTING
        !ev_mpa_type       TYPE mpa_template_type
        !et_asset_xls_data TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data
        !et_message        TYPE bapirettab
      RAISING
        cx_mpa_exception_handler .
    "! Get the number range
    METHODS get_number
      RETURNING
        VALUE(rv_number) TYPE mpa_fileid .
    "! Handle the exceptions and extract messages from exception
    METHODS handle_exception
      IMPORTING
        !iref_exception   TYPE REF TO cx_mpa_exception_handler
        !iv_je_sequence   TYPE string OPTIONAL
      CHANGING
        !ct_error_message TYPE bapirettab .
    "! Remove the commented lines from the uploaded excel file
    METHODS remove_comment_line
      EXPORTING
        !et_comment_lines TYPE gty_t_comment_line
      CHANGING
        !ct_excel_lines   TYPE mpa_t_index_value_pair OPTIONAL
        !ct_csv_lines     TYPE string_table OPTIONAL .
    "! Based on user configuration format amount
    METHODS user_confign_decimal_format
      CHANGING
        !ch_value TYPE string .
    "! if used passed more then length of structure then this message get trigger
    METHODS get_excel_length_error_msg
      IMPORTING
        !iv_cell_name     TYPE string
        !iv_index         TYPE i
        !iv_length        TYPE ddleng
      RETURNING
        VALUE(rt_message) TYPE bapirettab .
    "! Check MPA status
    METHODS check_mpa_status
      EXPORTING
        !ev_adjustment_status TYPE bapi_mtype
        ev_retirement_status  TYPE bapi_mtype
        !ev_create_status     TYPE bapi_mtype
        !ev_change_status     TYPE bapi_mtype
        !ev_transfer_status   TYPE bapi_mtype .
    "! Build mass processing of asset internal table
    METHODS build_mpa_xls_itab
      IMPORTING
        !is_mass_create     TYPE mpa_s_asset_create
        !is_mass_change     TYPE mpa_s_asset_change
        !is_mass_transfer   TYPE mpa_s_asset_transfer
        !is_mass_adjustment TYPE mpa_s_asset_adjustment
        !is_mass_retirement TYPE mpa_s_asset_retirement
      CHANGING
        !ct_mass_create     TYPE mpa_t_asset_create
        !ct_mass_change     TYPE mpa_t_asset_change
        !ct_mass_transfer   TYPE mpa_t_asset_transfer
        !ct_mass_adjustment TYPE mpa_t_asset_adjustment
        !ct_mass_retirement TYPE mpa_t_asset_retirement .
    "! Insert file data into MPA_ASSET_DATA table
    METHODS insert_file_to_db
      IMPORTING
        is_file_data       TYPE mpa_asset_data
      RETURNING
        VALUE(rs_messages) TYPE bapiret2 .

    METHODS transform_xstring_to_linetab
      IMPORTING
        iv_xstring TYPE xstring
      EXPORTING
        et_linetab TYPE string_table
      RAISING
        cx_mpa_exception_handler.
    METHODS set_supported_separators .
    METHODS find_separator
      IMPORTING
        !it_linetab   TYPE string_table
      EXPORTING
        !ev_separator TYPE string
      RAISING
        cx_mpa_exception_handler .
    METHODS check_csv_layout
      IMPORTING
        !it_line       TYPE string_table
        iv_separator   TYPE string
      EXPORTING
        ev_status_flag TYPE abap_bool
      RAISING
        cx_mpa_exception_handler .
    METHODS process_lines
      IMPORTING
        !it_line       TYPE string_table
        iv_status_flag TYPE abap_bool
      EXPORTING
        !ev_mpa_type   TYPE mpa_template_type
        !et_asset_data TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data
        et_message     TYPE bapirettab
      RAISING
        cx_mpa_exception_handler .
    METHODS get_value_count
      IMPORTING
        !it_cells_tab   TYPE string_table
      RETURNING
        VALUE(ev_count) TYPE i .
    METHODS split_csv_line
      IMPORTING
        iv_line              TYPE string
        iv_empty_symbol      TYPE string OPTIONAL
      EXPORTING
        et_line_table        TYPE string_table
        ev_empty_cell_number TYPE i .
ENDCLASS.



CLASS zcl_xlsx_parse_util IMPLEMENTATION.


  METHOD map_excel_data.

    TYPES : BEGIN OF lty_s_slno,
              slno TYPE mpa_slno,
            END OF lty_s_slno,

            lty_t_slno TYPE STANDARD TABLE OF lty_s_slno.

    DATA: ls_field_mapping            TYPE gty_s_field_name_mapping,
          lt_field_mapping_header     TYPE gty_t_field_name_mapping,
          lt_str_property             TYPE if_mpa_xlsx_parse_util~ty_t_struct_properties,
          lt_mass_transfer_xls_data   TYPE mpa_t_asset_transfer,
          ls_mass_transfer            TYPE mpa_s_asset_transfer,
          lt_mass_create_xls_data     TYPE mpa_t_asset_create,
          ls_mass_create              TYPE mpa_s_asset_create,
          lt_mass_change_xls_data     TYPE mpa_t_asset_change,
          ls_mass_change              TYPE mpa_s_asset_change,
          ls_mass_adjustment          TYPE mpa_s_asset_adjustment,
          lt_mass_adjustment_xls_data TYPE mpa_t_asset_adjustment,
          ls_mass_retirement          TYPE mpa_s_asset_retirement,
          lt_mass_retirement_xls_data TYPE mpa_t_asset_retirement,
          ls_slno                     TYPE lty_s_slno,
          lt_slno                     TYPE lty_t_slno.

    FIELD-SYMBOLS: <ft_line>   TYPE mpa_t_index_value_pair,
                   <fs_cell>   TYPE mpa_s_index_value_pair,
                   <fs_value>  TYPE string,
                   <dyn_value> TYPE any.


    " -----------------Get structure property -----------"
    CASE  zcl_xlsx_parse_util=>gv_template_type.

      WHEN if_mpa_output=>gc_mpa_temp-transfer.

        if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_transfer
                                                      IMPORTING et_struct_properties = lt_str_property ).
        ev_mpa_type = if_mpa_output=>gc_mpa_scen-transfer.

      WHEN if_mpa_output=>gc_mpa_temp-create.

        if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_create
                                                      IMPORTING et_struct_properties = lt_str_property ).
        ev_mpa_type = if_mpa_output=>gc_mpa_scen-create.

      WHEN if_mpa_output=>gc_mpa_temp-change.

        if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_change
                                                      IMPORTING et_struct_properties = lt_str_property ).
        ev_mpa_type = if_mpa_output=>gc_mpa_scen-change.

      WHEN if_mpa_output=>gc_mpa_temp-adjustment.

        if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_adjustment
                                                      IMPORTING et_struct_properties = lt_str_property ).
        ev_mpa_type = if_mpa_output=>gc_mpa_scen-adjustment.

      WHEN if_mpa_output=>gc_mpa_temp-retirement.

        if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_retirement
                                                      IMPORTING et_struct_properties = lt_str_property ).
        ev_mpa_type = if_mpa_output=>gc_mpa_scen-retirement.

    ENDCASE.

    " Parse technical label line, construct field-column mapping
    READ TABLE it_table ASSIGNING FIELD-SYMBOL(<fs_header_techn>) WITH TABLE KEY index = is_index-header_techn.
    ASSIGN <fs_header_techn>-value->* TO <ft_line>.
    LOOP AT <ft_line> ASSIGNING <fs_cell>.
      ASSIGN <fs_cell>-value->* TO <fs_value>.

      " Construct Header's field mapping
      INSERT VALUE gty_s_field_name_mapping( cell_name = <fs_value>
                                            cell_posi = <fs_cell>-index
                                            stru_name = <fs_value> )
                                 INTO TABLE lt_field_mapping_header.
    ENDLOOP.

    DATA lv_excel_data_index TYPE i VALUE 1.

    " Parse data lines, base on field-column mapping
    LOOP AT is_index-data INTO DATA(lv_header_index).

      READ TABLE it_table ASSIGNING FIELD-SYMBOL(<fs_data>) WITH TABLE KEY index = lv_header_index.
      ASSIGN <fs_data>-value->* TO <ft_line>.

      LOOP AT <ft_line> ASSIGNING <fs_cell>.
        ASSIGN <fs_cell>-value->* TO <fs_value>.

        READ TABLE lt_field_mapping_header INTO ls_field_mapping WITH KEY cell_posi = <fs_cell>-index.

        READ TABLE lt_str_property INTO DATA(ls_str_property) WITH KEY field_name = ls_field_mapping-stru_name.
        IF sy-subrc IS INITIAL.

          IF ( ls_str_property-data_type <> 'DATS' AND strlen( <fs_value> ) GT ls_str_property-length ) OR
             ( ls_str_property-data_type = 'DATS' AND strlen( <fs_value> ) GT ( ls_str_property-length + 2 ) ) .

            et_message = get_excel_length_error_msg( EXPORTING iv_cell_name = ls_field_mapping-cell_name
                                                               iv_index     = lv_excel_data_index
                                                               iv_length    = ls_str_property-length ).

            check_mpa_status( IMPORTING ev_create_status     = ls_mass_create-status
                                        ev_change_status     = ls_mass_change-status
                                        ev_transfer_status   = ls_mass_transfer-status
                                        ev_adjustment_status = ls_mass_adjustment-status
                                        ev_retirement_status = ls_mass_retirement-status  ).

          ENDIF.
        ENDIF.

        DATA(lv_stru_name) = ls_field_mapping-stru_name.

        CASE zcl_xlsx_parse_util=>gv_template_type.
          WHEN if_mpa_output=>gc_mpa_temp-transfer.
            ASSIGN COMPONENT lv_stru_name OF STRUCTURE ls_mass_transfer TO <dyn_value>.
          WHEN if_mpa_output=>gc_mpa_temp-create.
            ASSIGN COMPONENT lv_stru_name OF STRUCTURE ls_mass_create TO <dyn_value>.
          WHEN if_mpa_output=>gc_mpa_temp-change.
            ASSIGN COMPONENT lv_stru_name OF STRUCTURE ls_mass_change TO <dyn_value>.
          WHEN if_mpa_output=>gc_mpa_temp-adjustment.
            ASSIGN COMPONENT lv_stru_name OF STRUCTURE ls_mass_adjustment TO <dyn_value>.
          WHEN if_mpa_output=>gc_mpa_temp-retirement.
            ASSIGN COMPONENT lv_stru_name OF STRUCTURE ls_mass_retirement TO <dyn_value> .
        ENDCASE.

*
*        IF strlen( <fs_value> ) GT ls_str_property-length.
*
*          user_confign_format_amount( CHANGING ch_value = <fs_value> ).
*
*          IF strlen( <fs_value> ) GT ls_str_property-length.
*
*            et_message = get_excel_length_error_msg( EXPORTING iv_cell_name = ls_field_mapping-cell_name
*                                                               iv_index     = lv_excel_data_index
*                                                               iv_length    = ls_str_property-length ).
*
*            check_mpa_status( IMPORTING ev_create_status     = ls_mass_transfer-status
*                                        ev_change_status     = ls_mass_transfer-status
*                                        ev_transfer_status   = ls_mass_transfer-status
*                                        ev_adjustment_status = ls_mass_transfer-status  ).
*
*          ENDIF.
*
*        ELSE.

        "Check row number should not be blank or duplicate number.
        IF ls_field_mapping-stru_name EQ 'SLNO'.
          READ TABLE lt_slno WITH KEY slno = <fs_value> TRANSPORTING NO FIELDS.
          IF sy-subrc IS INITIAL.
            et_message = VALUE #( BASE et_message ( type   = if_mpa_output=>gc_msg_type-error
                                                    id     = if_mpa_output=>gc_msgid-mpa
                                                    number = COND #( WHEN <fs_value> IS INITIAL
                                                                     THEN if_mpa_output=>gc_msgno_mpa-row_blank
                                                                     ELSE if_mpa_output=>gc_msgno_mpa-row_duplicate )
                                                    message_v1 = <fs_value> ) ).
          ELSE.
            lt_slno = VALUE #( BASE lt_slno ( slno = <fs_value> ) ).
          ENDIF.
        ENDIF.

        TRY.
            <dyn_value> = <fs_value>.


          CATCH cx_sy_conversion_no_number.

            user_confign_decimal_format( CHANGING ch_value = <fs_value> ).

            <dyn_value> = <fs_value>.
        ENDTRY.
*
*        ENDIF.


      ENDLOOP.

      build_mpa_xls_itab( EXPORTING is_mass_create     = ls_mass_create
                                    is_mass_change     = ls_mass_change
                                    is_mass_transfer   = ls_mass_transfer
                                    is_mass_adjustment = ls_mass_adjustment
                                    is_mass_retirement = ls_mass_retirement
                           CHANGING ct_mass_create     = lt_mass_create_xls_data
                                    ct_mass_change     = lt_mass_change_xls_data
                                    ct_mass_transfer   = lt_mass_transfer_xls_data
                                    ct_mass_adjustment = lt_mass_adjustment_xls_data
                                    ct_mass_retirement = lt_mass_retirement_xls_data ).

      CLEAR: ls_mass_transfer, ls_mass_create, ls_mass_change, ls_mass_adjustment, ls_mass_retirement.

      lv_excel_data_index += 1.

    ENDLOOP.

    DATA(ls_asset) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( index                = 1
                                                                         mass_transfer_data   = lt_mass_transfer_xls_data
                                                                         mass_create_data     = lt_mass_create_xls_data
                                                                         mass_change_data     = lt_mass_change_xls_data
                                                                         mass_adjustment_data = lt_mass_adjustment_xls_data
                                                                         mass_retirement_data = lt_mass_retirement_xls_data  ).
    IF ls_asset IS NOT INITIAL.
      APPEND ls_asset TO et_asset_xls_data.
    ENDIF.

    " Clear return data if contains error message
    IF et_message IS NOT INITIAL.
      CLEAR: lt_mass_transfer_xls_data, lt_mass_create_xls_data, lt_mass_change_xls_data, lt_mass_adjustment_xls_data, lt_mass_retirement_xls_data.
    ENDIF.

  ENDMETHOD.


  METHOD build_mpa_xls_itab.

    " Append data to tables
    IF is_mass_transfer IS NOT INITIAL.
      APPEND is_mass_transfer TO ct_mass_transfer.
    ELSEIF is_mass_create IS NOT INITIAL.
      APPEND is_mass_create   TO ct_mass_create.
    ELSEIF is_mass_change IS NOT INITIAL.
      APPEND is_mass_change   TO ct_mass_change.
    ELSEIF is_mass_adjustment IS NOT INITIAL.
      APPEND is_mass_adjustment TO ct_mass_adjustment.
    ELSEIF is_mass_retirement IS NOT INITIAL.
      APPEND is_mass_retirement TO ct_mass_retirement.
    ENDIF.

  ENDMETHOD.


  METHOD get_excel_length_error_msg.

    INSERT VALUE #( type       = if_mpa_output=>gc_msg_type-error
                    id         = if_mpa_output=>gc_msgid-mpa
                    number     = if_mpa_output=>gc_msgno_mpa-field_chars
                    message_v1 = iv_cell_name
                    message_v2 = iv_index
                    message_v3 = |{ iv_length ALPHA = OUT }|
                    system     = sy-sysid ) INTO TABLE rt_message.

  ENDMETHOD.


  METHOD check_mpa_status.

    IF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-transfer.
      ev_transfer_status =  if_mpa_output=>gc_msg_type-error.
    ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-create.
      ev_create_status =  if_mpa_output=>gc_msg_type-error.
    ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-change.
      ev_change_status =  if_mpa_output=>gc_msg_type-error.
    ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-adjustment.
      ev_adjustment_status =  if_mpa_output=>gc_msg_type-error.
    ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-retirement.
      ev_retirement_status =  if_mpa_output=>gc_msg_type-error.
    ENDIF.

  ENDMETHOD.


  METHOD user_confign_decimal_format.

    SELECT SINGLE dcpfm FROM usr01 WHERE bname = @sy-uname INTO @DATA(lv_decimal_format).

    CASE lv_decimal_format.

      WHEN space.
        REPLACE ALL OCCURRENCES OF '.' IN ch_value  WITH ''.
        REPLACE ALL OCCURRENCES OF ',' IN ch_value  WITH '.'.

      WHEN abap_true.
        REPLACE ALL OCCURRENCES OF ',' IN ch_value  WITH ''.

      WHEN 'Y'.
        REPLACE ALL OCCURRENCES OF REGEX '^\s*|\s*$' IN ch_value  WITH ''.
        REPLACE ALL OCCURRENCES OF ',' IN ch_value  WITH '.'.

    ENDCASE.

  ENDMETHOD.


  METHOD parse_xlsx.

    IF ix_file IS INITIAL.
      RAISE EXCEPTION TYPE cx_openxml_format
        EXPORTING
          textid = cx_openxml_format=>cx_openxml_empty.
    ENDIF.

    TRY .

        " get the file data into table "
        if_mpa_xlsx_parse_util~transform_xstring_2_tab(
         EXPORTING
           ix_file  = ix_file
         IMPORTING
           et_table = mt_excel_rows ).

        " Check the excel file format is okay "
        if_mpa_xlsx_parse_util~check_excel_layout(
          IMPORTING
            et_line_index = mt_line_index
          CHANGING
            ct_table      = mt_excel_rows ).

        " Read Excel Data based on index table"
        READ TABLE mt_line_index ASSIGNING FIELD-SYMBOL(<fs_line_index>) INDEX 1.
        IF sy-subrc IS INITIAL.

          " Map file data to internal table "
          me->map_excel_data(
            EXPORTING
              it_table                       = mt_excel_rows
              is_index                       = <fs_line_index>
            IMPORTING
              ev_mpa_type                    = ev_mpa_type
              et_asset_xls_data              = et_asset
              et_message                     = et_message ).

          CLEAR: mt_line_index.
        ENDIF.

      CATCH cx_mpa_exception_handler INTO DATA(lo_exception).
        "Handle all the exceptions and append messages to return table
        handle_exception( EXPORTING
                            iref_exception = lo_exception
                          CHANGING
                            ct_error_message = et_message ).

    ENDTRY.

  ENDMETHOD.


  METHOD get_number.

    DATA lv_returncode TYPE inri-returncode.

    CALL FUNCTION 'NUMBER_GET_NEXT'
      EXPORTING
        nr_range_nr             = '01'
        object                  = gc_nr_object
      IMPORTING
        number                  = rv_number
        returncode              = lv_returncode
      EXCEPTIONS
        interval_not_found      = 1
        number_range_not_intern = 2
        object_not_found        = 3
        quantity_is_0           = 4
        quantity_is_not_1       = 5
        interval_overflow       = 6
        buffer_overflow         = 7
        OTHERS                  = 8.

    IF sy-subrc <> 0 OR lv_returncode <> ' '.
      "message id sy-msgid type sy-msgty number sy-msgno with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 into data(lv_message).   "TODO Export the message
    ENDIF.

  ENDMETHOD.


  METHOD save_file_to_db.

    DATA(lv_file_id) = get_number( ) ."Get field id number

    IF lv_file_id IS NOT INITIAL.
      GET TIME STAMP FIELD DATA(lv_timestamp).
      DATA(ls_file_data) = VALUE mpa_asset_data( file_id     = lv_file_id
                                                 ernam       = sy-uname
                                                 erdat       = sy-datum
                                                 erzet       = sy-uzeit
                                                 timestamp   = lv_timestamp
                                                 scen_type   = iv_mpa_file_type
                                                 file_status = '1' "Initial
                                                 file_name   = iv_file_name
                                                 file_data   = ix_file_content ).

      "Insert file data into MPA_ASSET_DATA table
      rs_messages = insert_file_to_db( ls_file_data ).

    ELSE.
      " Number range not maintained for object ‘MPA_FILEID’
      rs_messages = VALUE #( type   = if_mpa_output=>gc_msg_type-error
                             id     = if_mpa_output=>gc_msgid-mpa
                             number = if_mpa_output=>gc_msgno_mpa-num_range ).
    ENDIF.


  ENDMETHOD.


  METHOD insert_file_to_db.

    INSERT mpa_asset_data FROM is_file_data.
    IF sy-subrc IS INITIAL.
      COMMIT WORK.
      "File data stored in database table for further processing
      rs_messages = VALUE #( type   = if_mpa_output=>gc_msg_type-success
                             id     = if_mpa_output=>gc_msgid-mpa
                             number = if_mpa_output=>gc_msgno_mpa-file_db ).
    ELSEIF sy-subrc IS NOT INITIAL.
      "Error while uploading file
      rs_messages = VALUE #( type   = if_mpa_output=>gc_msg_type-error
                             id     = if_mpa_output=>gc_msgid-mpa
                             number = if_mpa_output=>gc_msgno_mpa-error_file  ).
    ENDIF.

  ENDMETHOD.


  METHOD constructor.
    mo_lcl_xlsx_parse = NEW lcl_mpa_xlsx_parse_util( ).
  ENDMETHOD.


  METHOD get_instance.

    IF go_instance IS NOT BOUND.
      ro_instance = NEW zcl_xlsx_parse_util( ).
    ELSE.
      ro_instance = go_instance.
    ENDIF.

  ENDMETHOD.


  METHOD if_mpa_xlsx_parse_util~check_excel_layout.

    DATA: lt_begin_symbol_index TYPE STANDARD TABLE OF string,
          lt_begin_symbol_value TYPE HASHED TABLE OF string WITH UNIQUE KEY table_line,
          lv_start_index        TYPE i,
          lv_end_index          TYPE i,
          lv_relative_line_num  TYPE i,
          lv_label_index        TYPE string,
          lv_techn_index        TYPE string,
          ls_mass_transfer      TYPE mpa_s_asset_transfer,
          ls_mass_create        TYPE mpa_s_asset_create,
          ls_mass_change        TYPE mpa_s_asset_change,
          ls_mass_adjustment    TYPE mpa_s_asset_adjustment,
          ls_mass_retirement    TYPE mpa_s_asset_retirement,
          ls_line_index         TYPE mpa_s_excel_doc_index,
          lv_last_line_index    TYPE string,
          lv_template_type      TYPE char20,
          lv_data_begin_index   TYPE i,
          lv_data_start_column  TYPE char1.

    FIELD-SYMBOLS: <ft_line>   TYPE mpa_t_index_value_pair,
                   <fs_cell>   TYPE mpa_s_index_value_pair,
                   <ft_label>  TYPE mpa_t_index_value_pair,
                   <ft_techn>  TYPE mpa_t_index_value_pair,
                   <fs_value>  TYPE string,
                   <dyn_value> TYPE any.

    IF ct_table IS INITIAL.
      RAISE EXCEPTION TYPE cx_mpa_exception_handler
        EXPORTING
          textid = cx_mpa_exception_handler=>file_layout_not_ok.
    ENDIF.

    " Remove the comment lines in file
    remove_comment_line(
      IMPORTING
        et_comment_lines = gt_comment_line
      CHANGING
        ct_excel_lines   = ct_table ).

    " Check first line, title should not be changed"
    LOOP AT ct_table ASSIGNING FIELD-SYMBOL(<fs_title>).

      lv_data_begin_index = <fs_title>-index.

      ASSIGN <fs_title>-value->* TO <ft_line>.
      UNASSIGN <fs_cell>.

      LOOP AT <ft_line> ASSIGNING <fs_cell>.

        ASSIGN <fs_cell>-value->* TO <fs_value>.
        " check for all scenarios
        IF strlen( <fs_value> ) < 60 AND ( <fs_value> CS if_mpa_output=>gc_mpa_temp-transfer
                                        OR <fs_value> CS if_mpa_output=>gc_mpa_temp-create
                                        OR <fs_value> CS if_mpa_output=>gc_mpa_temp-change
                                        OR <fs_value> CS if_mpa_output=>gc_mpa_temp-adjustment
                                        OR <fs_value> CS if_mpa_output=>gc_mpa_temp-retirement ) .
          CLEAR: zcl_xlsx_parse_util=>gv_template_type.
          zcl_xlsx_parse_util=>gv_template_type = <fs_value>.
        ELSE.
          UNASSIGN <fs_value>.
        ENDIF.
        EXIT.
      ENDLOOP.

      "loop over the table until the template type is found
      IF <fs_value> IS ASSIGNED.
        EXIT.
      ENDIF.
    ENDLOOP.


    " Check begin symbol line
    LOOP AT ct_table ASSIGNING FIELD-SYMBOL(<fs_line>) WHERE index > lv_data_begin_index.

      ASSIGN <fs_line>-value->* TO <ft_line>.
      IF <ft_line> IS INITIAL.
        CONTINUE.
      ENDIF.

      lv_last_line_index = <fs_line>-index.
      UNASSIGN <fs_cell>.

      IF lv_data_start_column IS INITIAL.
        LOOP AT <ft_line> ASSIGNING <fs_cell>. " the index of the 1st column with data is stored here
          lv_data_start_column = <fs_cell>-index.
          EXIT.
        ENDLOOP.
      ELSE.
        READ TABLE <ft_line> ASSIGNING <fs_cell> WITH TABLE KEY index = lv_data_start_column.
      ENDIF.

      IF <fs_cell> IS NOT ASSIGNED.
        CONTINUE.
      ENDIF.

      APPEND <fs_line>-index TO lt_begin_symbol_index.
      INSERT <fs_value> INTO TABLE lt_begin_symbol_value.

    ENDLOOP.

    "-------Template contains at least one index----------"
    IF lt_begin_symbol_index IS INITIAL.  "
      RAISE EXCEPTION TYPE cx_mpa_exception_handler
        EXPORTING
          textid = cx_mpa_exception_handler=>file_layout_not_ok.
    ENDIF.

    "check content base on begin symbol line
    DO lines( lt_begin_symbol_index ) TIMES.

      lv_start_index = lt_begin_symbol_index[ sy-index ].
      IF sy-index = lines( lt_begin_symbol_index ).
        lv_end_index = + lv_last_line_index + 1.
      ELSE.
        lv_end_index = lt_begin_symbol_index[ sy-index + 1 ].
      ENDIF.

      IF lv_start_index IS INITIAL OR lv_end_index IS INITIAL OR lv_start_index >= lv_end_index.
        RAISE EXCEPTION TYPE cx_mpa_exception_handler
          EXPORTING
            textid = cx_mpa_exception_handler=>file_layout_not_ok.
      ENDIF.

      lv_relative_line_num  = 1.

      CLEAR: lv_label_index, lv_techn_index,  ls_line_index.

      " store begin symbol line number
      ls_line_index-begin_symbol = lt_begin_symbol_index[ sy-index ].

      "validate the heading, data type and data
      LOOP AT ct_table ASSIGNING <fs_line> WHERE index >= lv_data_begin_index.
        ASSIGN <fs_line>-value->* TO <ft_line>.

        IF lv_relative_line_num = 2.  " Technical Name line

          lv_techn_index = <fs_line>-index.
          "loop over the technical names in the uploaded excel and check the content against the structure
          LOOP AT <ft_line> ASSIGNING <fs_cell>.
            UNASSIGN <dyn_value>.
            ASSIGN <fs_cell>-value->* TO <fs_value>.

            "set indicator to identify if the template contains status
            ev_slno_status_flag = COND #( WHEN <fs_value> = 'STATUS'
                                          THEN abap_true
                                          ELSE ev_slno_status_flag ).

            IF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-transfer.
              ASSIGN COMPONENT <fs_value> OF STRUCTURE ls_mass_transfer TO <dyn_value>.
            ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-create.
              ASSIGN COMPONENT <fs_value> OF STRUCTURE ls_mass_create TO <dyn_value>.
            ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-change.
              ASSIGN COMPONENT <fs_value> OF STRUCTURE ls_mass_change TO <dyn_value>.
            ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-adjustment.
              ASSIGN COMPONENT <fs_value> OF STRUCTURE ls_mass_adjustment TO <dyn_value>.
            ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-retirement.
              ASSIGN COMPONENT <fs_value> OF STRUCTURE ls_mass_retirement TO <dyn_value>.
            ENDIF.

            IF <dyn_value> IS NOT ASSIGNED.
              RAISE EXCEPTION TYPE cx_mpa_exception_handler
                EXPORTING
                  textid = cx_mpa_exception_handler=>file_layout_not_ok.
            ENDIF.
          ENDLOOP.

          ls_line_index-header_techn    = <fs_line>-index.     " store the index of technical fields name
          DATA(lv_possible_label_index) = lv_techn_index + 1 .

        ELSEIF lv_relative_line_num = 3 AND <fs_line>-index = lv_possible_label_index.  " Label name line -- The label will  be in the next row as the technical field names

          lv_label_index = <fs_line>-index.
          IF lv_label_index IS INITIAL OR lv_techn_index IS INITIAL. " Technical name line should be before label name line
            RAISE EXCEPTION TYPE cx_mpa_exception_handler
              EXPORTING
                textid = cx_mpa_exception_handler=>file_layout_not_ok.
          ENDIF.

          " check if the technical name has label name
          LOOP AT <ft_line> ASSIGNING <fs_cell>.

            READ TABLE ct_table ASSIGNING FIELD-SYMBOL(<fs_label_header>) WITH TABLE KEY index = lv_techn_index.
            ASSIGN <fs_label_header>-value->* TO <ft_label>.
            IF NOT line_exists( <ft_label>[ index = <fs_cell>-index ] ).  " technical name without label name
              RAISE EXCEPTION TYPE cx_mpa_exception_handler
                EXPORTING
                  textid = cx_mpa_exception_handler=>file_layout_not_ok.
            ENDIF.
          ENDLOOP.

        ELSEIF lv_relative_line_num =  3.    "data line should be after label name line and technical name line
          IF  lv_techn_index IS INITIAL.
            RAISE EXCEPTION TYPE cx_mpa_exception_handler
              EXPORTING
                textid = cx_mpa_exception_handler=>file_layout_not_ok.
          ENDIF.

          LOOP AT <ft_line> ASSIGNING <fs_cell>.
            READ TABLE ct_table ASSIGNING FIELD-SYMBOL(<fs_techn_header>) WITH TABLE KEY index = lv_techn_index.
            ASSIGN <fs_techn_header>-value->* TO <ft_techn>.

            IF NOT line_exists( <ft_techn>[ index = <fs_cell>-index ] ).  " technical name without label name
              RAISE EXCEPTION TYPE cx_mpa_exception_handler
                EXPORTING
                  textid = cx_mpa_exception_handler=>file_layout_not_ok.
            ENDIF.
          ENDLOOP.

          APPEND <fs_line>-index TO ls_line_index-data. " store the index

        ELSEIF lv_relative_line_num = 4 AND lv_label_index IS NOT INITIAL.

          IF lv_label_index IS INITIAL OR lv_techn_index IS INITIAL.  "data line should be after label name line and technical name line
            RAISE EXCEPTION TYPE cx_mpa_exception_handler
              EXPORTING
                textid = cx_mpa_exception_handler=>file_layout_not_ok.
          ENDIF.

          LOOP AT <ft_line> ASSIGNING <fs_cell>.
            READ TABLE ct_table ASSIGNING FIELD-SYMBOL(<fs_techn_head_of_label_>) WITH TABLE KEY index = lv_techn_index.
            ASSIGN <fs_techn_head_of_label_>-value->* TO <ft_techn>.

            IF NOT line_exists( <ft_techn>[ index = <fs_cell>-index ] ).  " technical name without label name
              RAISE EXCEPTION TYPE cx_mpa_exception_handler
                EXPORTING
                  textid = cx_mpa_exception_handler=>file_layout_not_ok.
            ENDIF.
          ENDLOOP.

          APPEND <fs_line>-index TO ls_line_index-data. " store the index

        ELSEIF  lv_relative_line_num > 3.

          APPEND <fs_line>-index TO ls_line_index-data. " store the index

        ENDIF.

        lv_relative_line_num = lv_relative_line_num + 1. " step to next line

      ENDLOOP.

      IF ls_line_index IS NOT INITIAL.
        " Output line index
        IF ls_line_index-header_techn IS INITIAL.
          RAISE EXCEPTION TYPE cx_mpa_exception_handler
            EXPORTING
              textid = cx_mpa_exception_handler=>file_layout_not_ok.
        ENDIF.

        IF ls_line_index-data IS INITIAL.
          RAISE EXCEPTION TYPE cx_mpa_exception_handler
            EXPORTING
              textid = cx_mpa_exception_handler=>empty_file.
        ENDIF.
        APPEND ls_line_index TO et_line_index.
      ENDIF.

    ENDDO.

  ENDMETHOD.


  METHOD if_mpa_xlsx_parse_util~delete_file_from_db.

    DATA: ls_key  TYPE /iwbep/s_mgw_tech_pair,
          ls_file TYPE mpa_asset_data.

    READ TABLE it_key_tab WITH KEY name = 'MASSPROCGASTUPLOADFILEID' INTO ls_key.
    IF sy-subrc IS INITIAL.
      SELECT SINGLE * FROM mpa_asset_data INTO ls_file WHERE file_id = ls_key-value.
      IF sy-subrc IS INITIAL.
        IF sy-uname EQ ls_file-ernam.
          DELETE FROM mpa_asset_data WHERE file_id     EQ ls_key-value
                                     AND   file_status EQ if_mpa_output=>gc_file_status-initial.
          IF sy-subrc IS INITIAL.
            COMMIT WORK.
            "Success
            rs_message = VALUE #( type   = if_mpa_output=>gc_msg_type-success
                                  id     = if_mpa_output=>gc_msgid-mpa
                                  number = if_mpa_output=>gc_msgno_mpa-file_deleted ).
          ELSE.
            "Error
            rs_message = VALUE #( type   = if_mpa_output=>gc_msg_type-error
                                  id     = if_mpa_output=>gc_msgid-mpa
                                  number = if_mpa_output=>gc_msgno_mpa-file_processed ).
          ENDIF.

        ELSE.
          "User is not authorize to delete some one's file
          rs_message = VALUE #( type   = if_mpa_output=>gc_msg_type-error
                                id     = if_mpa_output=>gc_msgid-mpa
                                number = if_mpa_output=>gc_msgno_mpa-file_del_authzd ).
        ENDIF.
      ELSE.
        "File does not exist, can not be deleted
        rs_message = VALUE #( type   = if_mpa_output=>gc_msg_type-error
                              id     = if_mpa_output=>gc_msgid-mpa
                              number = if_mpa_output=>gc_msgno_mpa-file_not_exist ).
      ENDIF.
    ELSE.
      "Key field is not passed to delete record
      rs_message = VALUE #( type   = if_mpa_output=>gc_msg_type-error
                            id     = if_mpa_output=>gc_msgid-mpa
                            number = if_mpa_output=>gc_msgno_mpa-key_field ).
    ENDIF.

  ENDMETHOD.


  METHOD if_mpa_xlsx_parse_util~get_struct_properties.

    TYPES : BEGIN OF lty_data_element,
              field_name   TYPE name_feld,
              data_element TYPE rollname,
            END OF lty_data_element,
            lty_t_data_element TYPE STANDARD TABLE OF lty_data_element.

    DATA: lt_data_element TYPE lty_t_data_element,
          ls_data_element TYPE lty_data_element,
          lo_strucdescr   TYPE REF TO cl_abap_structdescr.

    lo_strucdescr ?= cl_abap_typedescr=>describe_by_data( iv_struct_name ).

    LOOP AT lo_strucdescr->components ASSIGNING FIELD-SYMBOL(<ls_components>).

      DATA(lo_datadescr) = lo_strucdescr->get_component_type( p_name = <ls_components>-name ).
      INSERT VALUE #( data_element = lo_datadescr->get_relative_name( )
                      field_name   = <ls_components>-name )
           INTO TABLE lt_data_element.
    ENDLOOP.

    SELECT leng AS length, datatype AS data_type, rollname AS data_element,decimals AS decimals
       INTO TABLE @DATA(lt_data) FROM dd04l
       FOR ALL ENTRIES IN @lt_data_element
       WHERE rollname = @lt_data_element-data_element AND as4local   = 'A'.

    LOOP AT lt_data_element INTO ls_data_element.

      READ TABLE lt_data INTO DATA(ls_data) WITH KEY data_element = ls_data_element-data_element.
      IF sy-subrc IS INITIAL.
        INSERT VALUE #( length       = ls_data-length
                        data_type    = ls_data-data_type
                        field_name   = ls_data_element-field_name
                        data_element = ls_data_element-data_element
                        decimals = ls_data-decimals )
             INTO TABLE et_struct_properties.

      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD if_mpa_xlsx_parse_util~process_upload_request.

    DATA: lt_message_log               TYPE bapirettab,
          lt_asset_in                  TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data,
          lv_mpa_file_type             TYPE mpa_template_type,
          lv_batchid_string            TYPE string,
          lv_error_occurred_at_least_1 TYPE abap_bool,
          lv_total_num                 TYPE i VALUE 0,
          lv_action_type               TYPE char1.

    IF iv_mime_type EQ zcl_xlsx_parse_util=>gc_file_type-excel OR iv_mime_type EQ zcl_xlsx_parse_util=>gc_file_type-csv.

      CASE iv_mime_type.
        WHEN zcl_xlsx_parse_util=>gc_file_type-excel.
          TRY .
              CALL METHOD me->parse_xlsx
                EXPORTING
                  ix_file     = ix_file
                IMPORTING
                  ev_mpa_type = lv_mpa_file_type
                  et_asset    = lt_asset_in
                  et_message  = et_message.

            CATCH cx_openxml_not_found cx_openxml_format INTO DATA(lo_execl_parsing_exception).
              "You have changed the template. Restore the template or download it again
              INSERT VALUE bapiret2( type   = if_mpa_output=>gc_msg_type-error
                                     id     = if_mpa_output=>gc_msgid-mpa
                                     number = if_mpa_output=>gc_msgno_mpa-chgd_tmpl ) INTO TABLE et_message.
          ENDTRY.

        WHEN zcl_xlsx_parse_util=>gc_file_type-csv.

          me->if_mpa_xlsx_parse_util~parse_csv( EXPORTING ix_file         = ix_file
                                                IMPORTING ev_mpa_type = lv_mpa_file_type
                                                          et_asset    = lt_asset_in
                                                          et_message  = et_message ).

*      when others.  "TODO 00handle the exception here
          "Handle if file mime type is not correct

      ENDCASE.

      " Save the upload file data to DB for further processing
      IF et_message IS INITIAL AND lv_mpa_file_type IS NOT INITIAL.

        DATA(ls_message) = save_file_to_db( EXPORTING ix_file_content  = ix_file
                                                      iv_file_name     = iv_file_name
                                                      iv_mpa_file_type = lv_mpa_file_type ).
        INSERT ls_message INTO TABLE et_message.
      ENDIF.

    ELSE.
      "Invalid file format   :TODO
      "Handle if file mime type is not correct
    ENDIF.

  ENDMETHOD.


  METHOD if_mpa_xlsx_parse_util~transform_xstring_2_tab.

    mo_lcl_xlsx_parse->transform_xstring_2_tab(
                         EXPORTING
                           ix_file         = ix_file
                           iv_download_file_ind = iv_download_file_ind
                         IMPORTING
                           et_table        = et_table
                           et_dd04l        = mt_dd04l
                         CHANGING
                           ct_fieldinfo    = gt_fieldinfo ).

  ENDMETHOD.


  METHOD handle_exception.

    DATA: lref_t100_message  TYPE REF TO if_t100_message,
          lv_uploading_error TYPE bapiret2,
          lv_current_line    TYPE i,
          lt_msg_longtext    TYPE TABLE OF bapitgb.

    lref_t100_message = iref_exception.
    DATA(lv_key)      = lref_t100_message->t100key.

    lv_uploading_error-type    = if_mpa_output=>gc_msg_type-error.
    lv_uploading_error-id      = lv_key-msgid.
    lv_uploading_error-number  = lv_key-msgno.
    lv_uploading_error-message = iref_exception->get_text( ).

    IF lv_uploading_error-number EQ 006.   "TODO --Check
      CALL FUNCTION 'BAPI_MESSAGE_GETDETAIL'
        EXPORTING
          id         = lv_key-msgid
          number     = lv_key-msgno
          language   = sy-langu
          textformat = 'ASC'
        TABLES
          text       = lt_msg_longtext.
      LOOP AT lt_msg_longtext ASSIGNING FIELD-SYMBOL(<fs>).
        CONCATENATE lv_uploading_error-message <fs>-line INTO lv_uploading_error-message SEPARATED BY cl_abap_char_utilities=>newline.
      ENDLOOP.
    ENDIF.

    APPEND lv_uploading_error TO ct_error_message.
    CLEAR lv_uploading_error.

  ENDMETHOD.


  METHOD remove_comment_line.

    DATA: lt_excel_lines      TYPE mpa_t_index_value_pair,
          lv_deleted_line_num TYPE i VALUE 0,
          lv_line_index_str   TYPE string.

    CLEAR gt_comment_line.

    IF ct_excel_lines IS SUPPLIED.  " this // is comment symbol
      LOOP AT ct_excel_lines ASSIGNING FIELD-SYMBOL(<ls_excel_line>).
        DATA(lv_new_line_index) = <ls_excel_line>-index - lv_deleted_line_num.
        DATA(lv_line_index) = <ls_excel_line>-index.
        DATA(lt_remark) = CAST mpa_t_index_value_pair( <ls_excel_line>-value )->*.
        READ TABLE lt_remark WITH KEY index = 'A' TRANSPORTING NO FIELDS.
        IF sy-subrc EQ 0.
          DATA(ls_remark) = CAST string( lt_remark[ index = 'A' ]-value )->* .

          IF ls_remark IS NOT INITIAL AND strlen( ls_remark ) > 2 AND ls_remark+0(2) EQ gc_comment_symbol.
            lv_deleted_line_num = lv_deleted_line_num + 1.
            APPEND VALUE #( index = lv_line_index skip_lines = lv_line_index - lv_deleted_line_num ) TO et_comment_lines.
            CONTINUE.
          ENDIF.

        ENDIF.

        lv_line_index_str = lv_new_line_index.
        INSERT VALUE #( index = lv_line_index_str value = <ls_excel_line>-value ) INTO TABLE lt_excel_lines.
      ENDLOOP.
      ct_excel_lines = lt_excel_lines.
    ENDIF.

    IF ct_csv_lines IS SUPPLIED.
      LOOP AT ct_csv_lines ASSIGNING FIELD-SYMBOL(<ls_csv_line>).
        IF <ls_csv_line> IS NOT INITIAL AND
          strlen( <ls_csv_line> ) > 2 AND
           ( <ls_csv_line>+0(2) = gc_comment_symbol OR  <ls_csv_line>+1(2) = gc_comment_symbol ).
          lv_deleted_line_num = lv_deleted_line_num + 1.
          APPEND VALUE #( index = sy-tabix skip_lines = sy-tabix - lv_deleted_line_num ) TO et_comment_lines.
        ENDIF.
      ENDLOOP.

      " delete the lines
      SORT et_comment_lines BY index DESCENDING.
      LOOP AT et_comment_lines INTO DATA(ls_comment_line).
        DELETE ct_csv_lines INDEX ls_comment_line-index.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.


  METHOD if_mpa_xlsx_parse_util~parse_csv.

    TRY.
        " supported separators are semicolon and comma
        set_supported_separators( ).

        " convert into string line table
        transform_xstring_to_linetab( EXPORTING iv_xstring                      =  ix_file
                                      IMPORTING et_linetab                      =  DATA(lt_linetab) ).
        "from lt_linetab the 1st row would give the type

        remove_comment_line( IMPORTING et_comment_lines = gt_comment_line
                              CHANGING ct_csv_lines     = lt_linetab )." Table of Strings
        " here the row 2 and 3 are are deleted so now lt_linetab only has data from row 4

        " find the delimiter( ; or , )
        find_separator(  EXPORTING it_linetab =  lt_linetab
                         IMPORTING ev_separator = ev_seperator ).


        " check format of csv file
        check_csv_layout( EXPORTING it_line = lt_linetab
                                    iv_separator = ev_seperator
                          IMPORTING ev_status_flag = ev_status_flag ).


        " process every line: header lineitems
        process_lines( EXPORTING  it_line                        = lt_linetab
                                  iv_status_flag = ev_status_flag
                        IMPORTING ev_mpa_type                    = ev_mpa_type
                                  et_asset_data                = et_asset
                                  et_message = et_message ).

      CATCH cx_mpa_exception_handler INTO DATA(lo_exception).
        "Handle all the exceptions and append messages to return table
        handle_exception( EXPORTING iref_exception = lo_exception
                           CHANGING ct_error_message = et_message ).
    ENDTRY.

    " delete the entries that error exist
    LOOP AT et_asset INTO DATA(ls_asset_item).
      IF ls_asset_item-error_exist EQ abap_true.
        DELETE et_asset INDEX sy-tabix.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD transform_xstring_to_linetab.
    DATA: lv_length   TYPE i,
          lv_string   TYPE string,
          lv_mimetype TYPE w3conttype,
          lt_binary   TYPE STANDARD TABLE OF x255.

*   convert file
    lcl_function_module=>get_instance( )->scms_xstring_to_binary(
      EXPORTING
        buffer          = iv_xstring
      IMPORTING
        output_length   = lv_length
      CHANGING
        binary_tab      = lt_binary ).

*    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
*      EXPORTING
*        buffer        = iv_xstring
*      IMPORTING
*        output_length = lv_length
*      TABLES
*        binary_tab    = lt_binary.

    lv_mimetype = zcl_xlsx_parse_util=>gc_file_type-csv.

    lcl_function_module=>get_instance( )->scms_binary_to_string(
      EXPORTING
        input_length  = lv_length
        mimetype      = lv_mimetype
      IMPORTING
        text_buffer   = lv_string
      CHANGING
        binary_tab    = lt_binary
      EXCEPTIONS
        failed        = 1
        OTHERS        = 2 ).
*    CALL FUNCTION 'SCMS_BINARY_TO_STRING'
*      EXPORTING
*        input_length = lv_length
*        mimetype     = lv_mimetype
*      IMPORTING
*        text_buffer  = lv_string
*      TABLES
*        binary_tab   = lt_binary
*      EXCEPTIONS
*        failed       = 1
*        OTHERS       = 2.

*     if conversion fails return error message and exit
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE cx_mpa_exception_handler
        EXPORTING
          textid = cx_mpa_exception_handler=>file_not_ok.
    ELSE.
*     handle carriage return
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN lv_string WITH '??'.
*     create lines
      SPLIT lv_string AT '??' INTO TABLE DATA(line_tab).
      et_linetab = line_tab.

    ENDIF.

  ENDMETHOD.


  METHOD set_supported_separators.
    mt_supported_separators = VALUE #( ( separator = gc_sep_semicolon )
                                       ( separator = gc_sep_comma ) ).
  ENDMETHOD.


  METHOD find_separator.

    DATA: lv_current_line TYPE string,
          lv_separator    TYPE string VALUE gc_sep_comma.

    IF lines( it_linetab ) < 3.
      RAISE EXCEPTION TYPE cx_mpa_exception_handler EXPORTING textid = cx_mpa_exception_handler=>file_layout_not_ok.
    ENDIF.

    LOOP AT it_linetab INTO lv_current_line FROM 2.
      " find the separator from line 2, because the 5th charactor is separator from second line (since the 1st column is SLNO )
      "change condition to regex and delete all char and numeric to find the symbol - lohid
      IF lv_current_line IS NOT INITIAL.
        lv_separator = lv_current_line+4(1).
        IF line_exists( mt_supported_separators[ separator = lv_separator ] ).
          ev_separator = mv_separator = lv_separator. "change from instance variable to parameter
          RETURN.
        ELSE.
          RAISE EXCEPTION TYPE cx_mpa_exception_handler EXPORTING textid = cx_mpa_exception_handler=>file_layout_not_ok.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD check_csv_layout.
    DATA:
      lv_line            TYPE string,
      lv_cell            TYPE string,
      lt_cells           TYPE string_table,
      lv_block_line_num  TYPE i VALUE 0,
      lt_labels          TYPE string_table,
      lt_fieldnames      TYPE string_table,
      lv_block_num       TYPE i VALUE 0,
      lv_lineitem_symbol TYPE string,
      lv_batch_id_line   TYPE string,
      lv_forth_line      TYPE string,
      lv_fifth_line      TYPE string,
      ls_mass_transfer   TYPE mpa_s_asset_transfer,
      ls_mass_create     TYPE mpa_s_asset_create,
      ls_mass_change     TYPE mpa_s_asset_change,
      ls_mass_adjustment TYPE mpa_s_asset_adjustment,
      ls_mass_retirement TYPE mpa_s_asset_retirement,
      lv_cell_index      TYPE i VALUE 1,
      lt_mand_cell_index TYPE TABLE OF i,
      lv_header_symbol   TYPE string,
      lv_desc_sep        TYPE string.

    FIELD-SYMBOLS: <dyn_value>     TYPE any,
                   <lv_cell_value> TYPE any.

    " check csv file template, csv file must have more than 3 lines
    IF lines( it_line ) < 3.
      RAISE EXCEPTION TYPE cx_mpa_exception_handler
        EXPORTING
          textid = cx_mpa_exception_handler=>file_layout_not_ok.
    ELSE.
      CLEAR zcl_xlsx_parse_util=>gv_template_type.
    ENDIF.


    LOOP AT it_line INTO lv_line.
      CONDENSE lv_line.
      IF lv_block_line_num <> 3.
        SPLIT lv_line AT iv_separator INTO TABLE lt_cells.
      ELSE.
        lv_desc_sep = |"{ iv_separator }"|.
        SPLIT lv_line AT lv_desc_sep INTO TABLE lt_cells.
      ENDIF.
      CONCATENATE LINES OF lt_cells INTO DATA(lv_line_without_sep) SEPARATED BY ''.
      CONDENSE lv_line_without_sep.
      IF lv_line_without_sep IS NOT INITIAL.

        " header line(Sequence No,Header)
        "from row 2
        IF lt_cells[ 1 ] IS NOT INITIAL AND zcl_xlsx_parse_util=>gv_template_type IS INITIAL."mpa file header
          " header title only first cell have value, e.g. 1, header
          " or else raise format excepton.
          LOOP AT lt_cells INTO lv_cell FROM 2 WHERE table_line IS NOT INITIAL.
            RAISE EXCEPTION TYPE cx_mpa_exception_handler
              EXPORTING
                textid = cx_mpa_exception_handler=>file_layout_not_ok.
          ENDLOOP.
          " set header and lineitem symbol
          IF lv_header_symbol IS INITIAL.
            lv_header_symbol = lt_cells[ 1 ].
          ENDIF.
*        ELSE.
*          " first cell has value at no header title part, raise format exception
*          IF lv_header_symbol NE lt_cells[ 1 ].
*            RAISE EXCEPTION TYPE cx_mpa_exception_handler
*              EXPORTING
*                textid = cx_mpa_exception_handler=>file_layout_not_ok.
*          ENDIF.
*        ENDIF.
        " block start with 1
        lv_block_line_num = 1.

        " check for all scenarios
        IF strlen( lt_cells[ 1 ] ) < 60 AND ( lt_cells[ 1 ] CS if_mpa_output=>gc_mpa_temp-transfer
                                        OR lt_cells[ 1 ] CS if_mpa_output=>gc_mpa_temp-create
                                        OR lt_cells[ 1 ] CS if_mpa_output=>gc_mpa_temp-change
                                        OR lt_cells[ 1 ] CS if_mpa_output=>gc_mpa_temp-adjustment
                                        OR lt_cells[ 1 ] CS if_mpa_output=>gc_mpa_temp-retirement ) .
          CLEAR: zcl_xlsx_parse_util=>gv_template_type.
          zcl_xlsx_parse_util=>gv_template_type = lt_cells[ 1 ].
        ENDIF.

      ELSE.
        IF lv_header_symbol EQ lt_cells[ 1 ].
          " block start with 1
          lv_block_line_num = 1.
        ENDIF.
      ENDIF.

      " field name line, validate type will be executed in process line
      IF lv_block_line_num = 2.
        "check if all the fields from the structure is present in the template
        LOOP AT lt_cells ASSIGNING <lv_cell_value>.
          IF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-transfer.
            ASSIGN COMPONENT <lv_cell_value> OF STRUCTURE ls_mass_transfer TO <dyn_value>.
          ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-create.
            ASSIGN COMPONENT <lv_cell_value> OF STRUCTURE ls_mass_create TO <dyn_value>.
          ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-change.
            ASSIGN COMPONENT <lv_cell_value> OF STRUCTURE ls_mass_change TO <dyn_value>.
          ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-adjustment.
            ASSIGN COMPONENT <lv_cell_value> OF STRUCTURE ls_mass_adjustment TO <dyn_value>.
          ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-retirement.
            ASSIGN COMPONENT <lv_cell_value> OF STRUCTURE ls_mass_retirement TO <dyn_value>.
          ENDIF.

          IF <dyn_value> IS NOT ASSIGNED.
            RAISE EXCEPTION TYPE cx_mpa_exception_handler
              EXPORTING
                textid = cx_mpa_exception_handler=>file_layout_not_ok.
          ENDIF.
        ENDLOOP.

        lt_fieldnames = lt_cells.

        IF lt_fieldnames[ 2 ] = 'STATUS'.
          ev_status_flag = abap_true.
        ENDIF.

        UNASSIGN <lv_cell_value>.
      ENDIF.

      " label line, will compare with field line
      IF lv_block_line_num = 3.
        lt_labels = lt_cells.
        DATA(lv_labels_count) = get_value_count( lt_labels ).
        DATA(lv_fieldnames_count) = get_value_count( lt_fieldnames ).
        " raise exception, if label and fieldname can't match
        IF lv_labels_count NE lv_fieldnames_count.
          RAISE EXCEPTION TYPE cx_mpa_exception_handler
            EXPORTING
              textid = cx_mpa_exception_handler=>file_layout_not_ok.
        ELSE.
          "get cell index for all mandatory fields
          LOOP AT lt_cells ASSIGNING <lv_cell_value>.

            IF <lv_cell_value> CS '*'.
              APPEND lv_cell_index TO lt_mand_cell_index.
            ENDIF.
            lv_cell_index += 1.

          ENDLOOP.
          UNASSIGN <lv_cell_value>.

        ENDIF.

      ENDIF.

      IF lv_block_line_num > 3.

        IF lines( lt_cells ) > lines( lt_fieldnames ).
          RAISE EXCEPTION TYPE cx_mpa_exception_handler
            EXPORTING
              textid = cx_mpa_exception_handler=>row_has_too_many_entries_error
              msgv1  = CONV #( lv_block_line_num + 2 ).
        ELSE.
          "check if the mandatory fields are filled
          LOOP AT lt_mand_cell_index ASSIGNING FIELD-SYMBOL(<lv_mand_field_index>).

            IF lt_cells[ <lv_mand_field_index> ] IS INITIAL.
              RAISE EXCEPTION TYPE cx_mpa_exception_handler
                EXPORTING
                  textid = cx_mpa_exception_handler=>mand_field_empty.
            ENDIF.
          ENDLOOP.
        ENDIF.
      ENDIF.

      lv_block_line_num = lv_block_line_num + 1.
    ENDIF.

  ENDLOOP.


ENDMETHOD.


METHOD process_lines.

  DATA:
    lv_line               TYPE string,
    lv_cell               TYPE string,
    lt_cells              TYPE string_table,
    lv_fname              TYPE fac_posting_fieldname,
    lv_current_line       TYPE i VALUE 0,
    lv_empty_cell_number  TYPE i,
    " header and lineitems as different block
    lv_block_line_num     TYPE i VALUE 0,
    lv_block_content_type TYPE string,
    lt_data_fields        TYPE string_table,
    ls_line_items         TYPE fac_s_accdoc_itm_odata,
    ls_line_copas         TYPE fac_s_accdoc_itm_copa_odata,
    lv_block_cursor       TYPE string,
    lv_sequence_no        TYPE i,
    lv_message_counter    TYPE i VALUE 0.

  FIELD-SYMBOLS: <fs_field>        TYPE any,
                 <ls_accdoc_items> TYPE cl_fac_financials_post_dpc_ext=>ty_s_accdoc_gl_entry_odata,
                 <ls_gl_hdr>       TYPE fac_s_accdoc_hdr_odata,
                 <lt_gl_items>     TYPE fac_t_accdoc_itm_odata,
                 <lt_gl_copas>     TYPE fac_t_accdoc_itm_copa_odata,
                 <dyn_value>       TYPE any.

  DATA :
    lt_mass_transfer_data   TYPE mpa_t_asset_transfer,
    ls_mass_transfer        TYPE mpa_s_asset_transfer,
    lt_mass_create_data     TYPE mpa_t_asset_create,
    ls_mass_create          TYPE mpa_s_asset_create,
    lt_mass_change_data     TYPE mpa_t_asset_change,
    ls_mass_change          TYPE mpa_s_asset_change,
    ls_mass_adjustment      TYPE mpa_s_asset_adjustment,
    lt_mass_adjustment_data TYPE mpa_t_asset_adjustment,
    ls_mass_retirement      TYPE mpa_s_asset_retirement,
    lt_mass_retirement_data TYPE mpa_t_asset_retirement,
    lt_str_property         TYPE if_mpa_xlsx_parse_util~ty_t_struct_properties,
    lv_loop_count           TYPE i VALUE 1,
    lv_flag                 TYPE abap_bool.


  " -----------------Get structure property -----------"
  IF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-transfer.

    if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_transfer
                                                  IMPORTING et_struct_properties = lt_str_property ).
    ev_mpa_type = if_mpa_output=>gc_mpa_scen-transfer.

  ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-create.

    if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_create
                                                  IMPORTING et_struct_properties = lt_str_property ).
    ev_mpa_type = if_mpa_output=>gc_mpa_scen-create.

  ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-change.

    if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_change
                                                  IMPORTING et_struct_properties = lt_str_property ).
    ev_mpa_type = if_mpa_output=>gc_mpa_scen-change.

  ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-adjustment.

    if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_adjustment
                                                  IMPORTING et_struct_properties = lt_str_property ).
    ev_mpa_type = if_mpa_output=>gc_mpa_scen-adjustment.

  ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-retirement.

    if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = ls_mass_retirement
                                                  IMPORTING et_struct_properties = lt_str_property ).
    ev_mpa_type = if_mpa_output=>gc_mpa_scen-retirement.

  ENDIF.

  LOOP AT it_line INTO lv_line.
    " start from 1 and just if start to loop lines, line number will add 1
    lv_current_line = lv_current_line + 1.

    CONDENSE lv_line.
    " split line into array
    split_csv_line( EXPORTING iv_line         = lv_line
                    IMPORTING et_line_table = lt_cells
                              ev_empty_cell_number = lv_empty_cell_number ).
    " if current line is empty/blank, skip
    IF lv_empty_cell_number = lines( lt_cells ).
      CONTINUE.
    ENDIF.

    IF lv_current_line > 3.


      LOOP AT lt_str_property ASSIGNING FIELD-SYMBOL(<ls_str_property>).

        "check if the status field is filled in the template ie not 1st run of the data
        IF lv_loop_count = 2 AND iv_status_flag = abap_false AND lv_flag = abap_false.
          lv_flag = abap_true.
          CONTINUE.
        ENDIF.

        IF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-transfer.
          ASSIGN COMPONENT <ls_str_property>-field_name  OF STRUCTURE ls_mass_transfer TO <dyn_value>.
        ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-create.
          ASSIGN COMPONENT <ls_str_property>-field_name  OF STRUCTURE ls_mass_create TO <dyn_value>.
        ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-change.
          ASSIGN COMPONENT <ls_str_property>-field_name  OF STRUCTURE ls_mass_change TO <dyn_value>.
        ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-adjustment.
          ASSIGN COMPONENT <ls_str_property>-field_name  OF STRUCTURE ls_mass_adjustment TO <dyn_value>.
        ELSEIF zcl_xlsx_parse_util=>gv_template_type EQ if_mpa_output=>gc_mpa_temp-retirement.
          ASSIGN COMPONENT <ls_str_property>-field_name  OF STRUCTURE ls_mass_retirement TO <dyn_value> .
        ENDIF.

        IF lt_cells[ lv_loop_count ] <> '##empty##'.
          IF ( <ls_str_property>-data_type <> 'DATS' AND strlen( lt_cells[ lv_loop_count ] ) GT CONV i( <ls_str_property>-length ) ) OR
                       ( <ls_str_property>-data_type = 'DATS' AND strlen( lt_cells[ lv_loop_count ] ) GT ( CONV i( <ls_str_property>-length ) + 2 ) ) .

            APPEND LINES OF get_excel_length_error_msg( EXPORTING iv_cell_name =  CONV #( <ls_str_property>-field_name )
                                                                  iv_index     = CONV #( lt_cells[ 1 ] ) "the 1st row of every structure is serial number
                                                                  iv_length    = <ls_str_property>-length  ) TO et_message .

            check_mpa_status( IMPORTING ev_create_status     = ls_mass_create-status
                                        ev_change_status     = ls_mass_change-status
                                        ev_transfer_status   = ls_mass_transfer-status
                                        ev_adjustment_status = ls_mass_adjustment-status
                                        ev_retirement_status = ls_mass_retirement-status  ).
            CONTINUE.
          ENDIF.
          "convert the number with decimal into internal format
          IF <ls_str_property>-decimals IS NOT INITIAL.
            me->user_confign_decimal_format( CHANGING ch_value = lt_cells[ lv_loop_count ] ).
          ENDIF.
          "convert date into internal format
          IF <ls_str_property>-data_type = 'DATS'.
*              TRY.
            DATA(lv_date_internal) = lt_cells[ lv_loop_count ].
            REPLACE ALL OCCURRENCES OF REGEX '[^a-zA-Z\d:]' IN lv_date_internal WITH ''.

*                  cl_abap_datfm=>conv_date_int_to_ext( EXPORTING im_datint = CONV #( lv_date_internal ) ).

            lcl_function_module=>get_instance( )->date_check_plausibility( EXPORTING date = CONV syst_datum( lv_date_internal )
                                                                          EXCEPTIONS plausibility_check_failed = 1
                                                                                     OTHERS                    = 2 ).

            IF sy-subrc = 0.

              lt_cells[ lv_loop_count ] =  lv_date_internal.

            ELSE.

              APPEND VALUE #( type       = if_mpa_output=>gc_msg_type-error
                              id         = if_mpa_output=>gc_msgid-mpa
                              number     = if_mpa_output=>gc_msgno_mpa-date_format
                              message_v1 = lv_current_line
                              system     = sy-sysid ) TO et_message.
              check_mpa_status( IMPORTING ev_create_status     = ls_mass_create-status
                                          ev_change_status     = ls_mass_change-status
                                          ev_transfer_status   = ls_mass_transfer-status
                                          ev_adjustment_status = ls_mass_adjustment-status
                                          ev_retirement_status = ls_mass_retirement-status  ).
              lv_loop_count += 1.
              CONTINUE.

            ENDIF.


*                CATCH cx_abap_datfm_format_unknown.
*
*                  APPEND VALUE #( type       = if_mpa_output=>gc_msg_type-error
*                                  id         = if_mpa_output=>gc_msgid-mpa
*                                  number     = if_mpa_output=>gc_msgno_mpa-date_format
*                                  system     = sy-sysid ) TO et_message.
*                  check_mpa_status( IMPORTING ev_create_status     = ls_mass_create-status
*                                            ev_change_status     = ls_mass_change-status
*                                            ev_transfer_status   = ls_mass_transfer-status
*                                            ev_adjustment_status = ls_mass_adjustment-status
*                                            ev_retirement_status = ls_mass_retirement-status  ).
*                  CONTINUE.

          ENDIF.

          TRY.
              <dyn_value> = CONV #( lt_cells[ lv_loop_count ] ).

            CATCH cx_sy_conversion_no_number.

              user_confign_decimal_format( CHANGING ch_value = lt_cells[ lv_loop_count ] ).

              <dyn_value> = lt_cells[ lv_loop_count ].
          ENDTRY.

        ENDIF.

        lv_loop_count += 1.
      ENDLOOP.

      build_mpa_xls_itab( EXPORTING is_mass_create     = ls_mass_create
                                    is_mass_change     = ls_mass_change
                                    is_mass_transfer   = ls_mass_transfer
                                    is_mass_adjustment = ls_mass_adjustment
                                    is_mass_retirement = ls_mass_retirement
                           CHANGING ct_mass_create     = lt_mass_create_data
                                    ct_mass_change     = lt_mass_change_data
                                    ct_mass_transfer   = lt_mass_transfer_data
                                    ct_mass_adjustment = lt_mass_adjustment_data
                                    ct_mass_retirement = lt_mass_retirement_data ).

      CLEAR: ls_mass_transfer, ls_mass_create, ls_mass_change, ls_mass_adjustment, ls_mass_retirement,lv_flag.
      lv_loop_count = 1.
    ENDIF.

  ENDLOOP.

  DATA(ls_asset) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( index                = 1
                                                                       mass_transfer_data   = lt_mass_transfer_data
                                                                       mass_create_data     = lt_mass_create_data
                                                                       mass_change_data     = lt_mass_change_data
                                                                       mass_adjustment_data = lt_mass_adjustment_data
                                                                       mass_retirement_data = lt_mass_retirement_data  ).
  IF ls_asset IS NOT INITIAL.
    APPEND ls_asset TO et_asset_data.
  ENDIF.

ENDMETHOD.


METHOD get_value_count.
  LOOP AT it_cells_tab INTO DATA(lv_cell).
    IF lv_cell NE space.
      ev_count = ev_count + 1.
    ENDIF.
  ENDLOOP.
ENDMETHOD.


METHOD split_csv_line.
  DATA: lv_string_left        TYPE string,
        lv_csv_regx           TYPE string,
        lv_csv_regx_lookahead TYPE string VALUE '(?=([^\"]*\"[^\"]*\")*(?![^\"]*\"))',
        lv_quote_regx         TYPE string VALUE '"(?:[^"\\]|\\.)*"',
        lt_regex_result       TYPE match_result_tab,
        lv_cell_value         TYPE string,
        lv_regex_length       TYPE i,
        lt_result             TYPE STANDARD TABLE OF string,
        lv_empty_number       TYPE i VALUE 0.

  CONCATENATE mv_separator lv_csv_regx_lookahead INTO lv_csv_regx.
  lv_string_left = iv_line.
  DO.
    " find comma as separator
    FIND REGEX lv_csv_regx IN lv_string_left RESULTS lt_regex_result.
    IF lt_regex_result IS INITIAL.
      " last cell
      IF lv_string_left IS INITIAL.
        APPEND cs_empty_symbol TO lt_result.
        lv_empty_number = lv_empty_number + 1.
      ELSE.
        APPEND lv_string_left TO lt_result.
      ENDIF.
      EXIT.
    ENDIF.
    " get cell value
    lv_regex_length = lt_regex_result[ 1 ]-offset.
    lv_cell_value = lv_string_left+0(lv_regex_length).

    " delete quote
    FIND REGEX lv_quote_regx IN lv_cell_value RESULTS lt_regex_result.
    IF lt_regex_result IS NOT INITIAL.
      " delete previous quote
      SHIFT lv_cell_value LEFT DELETING LEADING '"'.
      " delete last quote
      SHIFT lv_cell_value RIGHT DELETING TRAILING '"'.
    ENDIF.

    " move to next
    SHIFT lv_string_left BY lv_regex_length + 1 PLACES LEFT.
    CONDENSE lv_cell_value.
    IF lv_cell_value IS INITIAL.
      APPEND cs_empty_symbol TO lt_result.
      lv_empty_number = lv_empty_number + 1.
    ELSE.
      APPEND lv_cell_value TO lt_result.
    ENDIF.
  ENDDO.

  et_line_table = lt_result.
  ev_empty_cell_number = lv_empty_number.
ENDMETHOD.


ENDCLASS.
