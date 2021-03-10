**"* use this source file for your ABAP unit test classes
CLASS ltc_mpa_xlsx_parse_util DEFINITION DEFERRED.
CLASS zcl_xlsx_parse_util DEFINITION LOCAL FRIENDS ltc_mpa_xlsx_parse_util.

CLASS ltc_mpa_xlsx_parse_util DEFINITION FOR TESTING
  FINAL
  DURATION MEDIUM
  RISK LEVEL HARMLESS .

  PRIVATE SECTION.
    CONSTANTS:
      gc_cc  TYPE bukrs VALUE '0001',
      gc_one TYPE char1 VALUE '1'.

    CLASS-DATA:
      go_osql_environment TYPE REF TO if_osql_test_environment,    "* Global instance of the OSQL test environment
      mo_mpa_parse_util   TYPE REF TO zcl_xlsx_parse_util,
      mo_lcl_parse_util   TYPE REF TO lcl_mpa_xlsx_parse_util,
      gv_mime_type        TYPE string.

    CLASS-METHODS:
      class_setup,                  "* Set up the OSQL test environment
      class_teardown.               "* Destroy the OSQL test environment

    DATA mt_csv_line_data TYPE string_table.

    METHODS:

      setup,                        "* Prepare the run before starting
      teardown,
      fill_mock_tables,             "* Fill tables in OSQL environment

      fill_xstring_file_data
        EXPORTING
          ex_file TYPE xstring,

      fill_table_file_data
        IMPORTING
          iv_scen_type           TYPE mpa_template_type
          iv_clear_title         TYPE boolean OPTIONAL
          iv_change_title_name   TYPE boolean OPTIONAL
          iv_error_data          TYPE boolean OPTIONAL
          iv_comment_cell_change TYPE boolean OPTIONAL
          iv_unwanted_field      TYPE boolean OPTIONAL
          iv_increase_row        TYPE boolean OPTIONAL
          iv_clear_techdesc      TYPE boolean OPTIONAL
          iv_no_data             TYPE boolean OPTIONAL
        EXPORTING
          et_table               TYPE mpa_t_index_value_pair,
      fill_table_dile_with_slno IMPORTING iv_scen_type TYPE mpa_template_type
                                EXPORTING
                                          et_table     TYPE mpa_t_index_value_pair,
      fill_table_dile_with_long_date IMPORTING iv_scen_type TYPE mpa_template_type
                                     EXPORTING
                                               et_table     TYPE mpa_t_index_value_pair,

      process_upload_request    FOR TESTING RAISING cx_static_check,

      "Other public methods
      get_instance            FOR TESTING RAISING cx_static_check,
      delete_file_from_db     FOR TESTING RAISING cx_static_check,

      "Local interface methods
      get_attr_from_node      FOR TESTING RAISING cx_static_check,

      test_lcl_private_method FOR TESTING RAISING cx_static_check,

      user_confign_format_amount FOR TESTING,

      save_file_to_db FOR TESTING,

      handle_exception FOR TESTING,

      map_excel_data FOR TESTING RAISING cx_mpa_exception_handler,

      check_mpa_status FOR TESTING,

      build_mpa_xls_itab FOR TESTING,

      insert_file_to_db FOR TESTING,

      get_excel_length_error_msg FOR TESTING,

      convert_dec_time_to_hhmmss FOR TESTING,

      convert_long_to_date FOR TESTING,

      check_excel_layout FOR TESTING RAISING cx_mpa_exception_handler,

      parse_xlsx FOR TESTING RAISING cx_openxml_format cx_openxml_not_found,

      remove_comment_line FOR TESTING RAISING cx_root,

      get_struct_properties FOR TESTING RAISING cx_root ,

      convert_dec_time FOR TESTING RAISING cx_static_check,

      set_supported_separators FOR TESTING RAISING cx_static_check,

      remove_comment_line_csv FOR TESTING RAISING cx_static_check,

      find_separator FOR TESTING RAISING cx_static_check,

      find_separator_invalid_file FOR TESTING RAISING cx_static_check,

      check_csv_layout FOR TESTING RAISING cx_static_check,

      check_csv_layout_with_status FOR TESTING RAISING cx_static_check,

      check_csv_exp_layout_not_ok FOR TESTING RAISING cx_static_check,

      process_lines FOR TESTING RAISING cx_static_check,

      parse_csv_macro FOR TESTING RAISING cx_static_check,

      process_lines_create FOR TESTING RAISING cx_static_check,

      process_lines_change FOR TESTING RAISING cx_static_check,

      process_lines_adjustment FOR TESTING RAISING cx_static_check,

      process_lines_retirement FOR TESTING RAISING cx_static_check,

      check_csv_layout_line_lt_3 FOR TESTING RAISING cx_static_check,

      check_csv_layout_create FOR TESTING RAISING cx_static_check,

      check_csv_layout_retirement FOR TESTING RAISING cx_static_check,

      check_csv_layout_change FOR TESTING RAISING cx_static_check,

      check_csv_layout_adjustment FOR TESTING RAISING cx_static_check,

      check_csv_layout_hdr_not_ok FOR TESTING RAISING cx_static_check,

      check_csv_layout_hdr_not_ok_2 FOR TESTING RAISING cx_static_check,

      check_csv_layout_hdr_not_ok_3 FOR TESTING RAISING cx_static_check,

      check_csv_layout_hdr_not_ok_4 FOR TESTING RAISING cx_static_check,

      check_csv_layout_hdr_not_ok_5 FOR TESTING RAISING cx_static_check,

      find_separator_invalid_sep FOR TESTING RAISING cx_static_check,

      map_excel_data_exp FOR TESTING RAISING cx_static_check,

    map_excel_data_exp_long_date FOR TESTING RAISING cx_static_check,

    check_csv_layout_more_cells FOR TESTING RAISING cx_static_check,

    get_value_count_initial FOR TESTING RAISING cx_static_check.

ENDCLASS.

CLASS ltc_mpa_xlsx_parse_util IMPLEMENTATION.

  METHOD class_setup.
    DATA lt_tables TYPE if_osql_test_environment=>ty_t_sobjnames.

*   List of tables that will be abstracted by OSQL
    lt_tables = VALUE #( ( 'MPA_ASSET_DATA' )
                         ( 'DD04T' )
                         ( 'DD04L' )
                         ( 'DD07V' )
                         ( 'TNRO' )
                         ( 'USR01' )
                         ( 'T002' )
     ) .

*   Register tables with OSQL environment
    ltc_mpa_xlsx_parse_util=>go_osql_environment = cl_osql_test_environment=>create( lt_tables ).


** Initialize all classes and interfaces
    mo_mpa_parse_util = NEW zcl_xlsx_parse_util( ).
    mo_lcl_parse_util = NEW lcl_mpa_xlsx_parse_util( ).
    mo_mpa_parse_util->mo_lcl_xlsx_parse = NEW lcl_mpa_xlsx_prs_util_mock( ).

    gv_mime_type  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'.

  ENDMETHOD.

  METHOD class_teardown.
    ltc_mpa_xlsx_parse_util=>go_osql_environment->disable_double_redirection( ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->destroy( ).
  ENDMETHOD.


  METHOD setup. "Instance Method
    fill_mock_tables( ). " Fill data to mock tables

    mt_csv_line_data = VALUE string_table( ( `Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                       ( `// Do not change the template. Instead; add the data in the corresponding field based on the scenario.;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                       ( `// Fields marked with an asterisk (*) are mandatory. After filling the template; upload it for further processing.;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                       ( `SLNO;BLART;BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR` )
                                       ( `"*Row";"Document";"*Document";"*Posting Date";"*Asset Value";"Fiscal";"*Translation";"*Company";"KJ";"hui";"as";"aa";"qq";"qw";"yx";"ff";"tt";"gh";"qwhh";"w";"q";"a";"y";"x";"c";"e";"f";"s";"g"` )
                                       ( `4;;2021-01-01;2021-01-01;2021-01-01;;12;2020-10-25;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;` )
                                       ( `5;AA;2021-01-01;2021-01-01;2021-01-01;;9;2020-06-15;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;` )  ).

  ENDMETHOD.

  METHOD teardown. "Instance method
    " Delete all test data generated by the test methods
    go_osql_environment->clear_doubles( ).
  ENDMETHOD.

  METHOD process_upload_request.

    DATA: lv_file_name TYPE string,
          lx_file      TYPE xstring,
          lv_scen_type TYPE mpa_template_type.

    fill_xstring_file_data( IMPORTING ex_file = lx_file ).

    DATA(lt_line_index) =  VALUE mpa_t_line_index( ( '4' )  ).
    mo_mpa_parse_util->mt_line_index = VALUE mpa_t_excel_doc_index( ( begin_symbol = '2'
                                                                      header_techn = '2'
                                                                      data         = lt_line_index ) ).
    "When 1
    lv_file_name  = 'AssetMassCreate_Template.XLSX'.
    lv_scen_type  =  if_mpa_output=>gc_mpa_scen-create.

    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-create.

    fill_table_file_data( EXPORTING iv_scen_type = lv_scen_type
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->if_mpa_xlsx_parse_util~process_upload_request(
     EXPORTING
       ix_file      = lx_file
       iv_mime_type = gv_mime_type
       iv_file_name = lv_file_name
     IMPORTING
       et_message   = DATA(lt_message_log) ).
    "Then 1
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-type
                                         exp  = if_mpa_output=>gc_msg_type-error
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-number
                                         exp  = '012'
                                         quit  = if_aunit_constants=>quit-no ).
    "When 2
    lv_file_name  = 'AssetMassTransfer_Template.XLSX'.
    lv_scen_type  =  if_mpa_output=>gc_mpa_scen-transfer.

    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.

    fill_table_file_data( EXPORTING iv_scen_type = lv_scen_type
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->if_mpa_xlsx_parse_util~process_upload_request(
     EXPORTING
       ix_file      = lx_file
       iv_mime_type = gv_mime_type
       iv_file_name = lv_file_name
     IMPORTING
       et_message   = lt_message_log ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-type
                                         exp  = if_mpa_output=>gc_msg_type-error
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-number
                                         exp  = '012'
                                         quit  = if_aunit_constants=>quit-no ).
    "When 3
    fill_table_file_data( EXPORTING iv_scen_type = lv_scen_type
                                    iv_unwanted_field  = abap_true
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->if_mpa_xlsx_parse_util~process_upload_request(
     EXPORTING
       ix_file      = lx_file
       iv_mime_type = gv_mime_type
       iv_file_name = lv_file_name
     IMPORTING
       et_message   = lt_message_log ).
    "Then 3
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-type
                                         exp  = if_mpa_output=>gc_msg_type-error
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-number
                                         exp  = '012'
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-type
                                         exp  = if_mpa_output=>gc_msg_type-error
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 2 ]-number
                                         exp  = '006'
                                         quit  = if_aunit_constants=>quit-no ).
    "When 4
    lv_file_name  = 'AssetMasschange_Template.XLSX'.
    lv_scen_type  =  if_mpa_output=>gc_mpa_scen-change.

    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-change.

    fill_table_file_data( EXPORTING iv_scen_type = lv_scen_type
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->if_mpa_xlsx_parse_util~process_upload_request(
     EXPORTING
       ix_file      = lx_file
       iv_mime_type = gv_mime_type
       iv_file_name = lv_file_name
     IMPORTING
       et_message   = lt_message_log ).
    "Then 4
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-type
                                         exp  = if_mpa_output=>gc_msg_type-error
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-number
                                         exp  = '012'
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 1 ]-type
                                         exp  = if_mpa_output=>gc_msg_type-error
                                         quit  = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act   = lt_message_log[ 2 ]-number
                                         exp  = '006'
                                         quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD delete_file_from_db.

    DATA : lt_mpa_asset_data TYPE STANDARD TABLE OF mpa_asset_data,
           lt_key_tab        TYPE /iwbep/t_mgw_tech_pairs.

    "When 1
    DATA(ls_message) = mo_mpa_parse_util->if_mpa_xlsx_parse_util~delete_file_from_db( EXPORTING it_key_tab = lt_key_tab ).
    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-key_field
                                        quit = if_aunit_constants=>quit-no ).

    "When 2
    lt_key_tab = VALUE #( ( name = 'MASSPROCGASTUPLOADFILEID' value = '12345' ) ).
    ls_message = mo_mpa_parse_util->if_mpa_xlsx_parse_util~delete_file_from_db( EXPORTING it_key_tab = lt_key_tab ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-file_not_exist
                                        quit = if_aunit_constants=>quit-no ).

    "When 3
    ls_message = mo_mpa_parse_util->if_mpa_xlsx_parse_util~delete_file_from_db( EXPORTING it_key_tab = lt_key_tab ).
    "Then 3
    lt_mpa_asset_data = VALUE #( ( file_id = '12345' ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_mpa_asset_data ).
    ls_message = mo_mpa_parse_util->if_mpa_xlsx_parse_util~delete_file_from_db( EXPORTING it_key_tab = lt_key_tab ).
    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-file_del_authzd
                                        quit = if_aunit_constants=>quit-no ).

    DELETE mpa_asset_data FROM TABLE lt_mpa_asset_data.
    COMMIT WORK.



    "When 4
    lt_mpa_asset_data = VALUE #( ( file_id = '12345'
                                  file_status = if_mpa_output=>gc_file_status-completed
                                   ernam   = sy-uname ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_mpa_asset_data ).

    ls_message = mo_mpa_parse_util->if_mpa_xlsx_parse_util~delete_file_from_db( EXPORTING it_key_tab = lt_key_tab ).
    "Then 4
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-file_processed
                                        quit = if_aunit_constants=>quit-no ).

    DELETE mpa_asset_data FROM TABLE lt_mpa_asset_data.
    COMMIT WORK.
    "When 5

    lt_mpa_asset_data = VALUE #( ( file_id = '12345'
                                  file_status = if_mpa_output=>gc_file_status-initial
                                   ernam   = sy-uname ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_mpa_asset_data ).

    ls_message = mo_mpa_parse_util->if_mpa_xlsx_parse_util~delete_file_from_db( EXPORTING it_key_tab = lt_key_tab ).
    "Then 5
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-success
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-file_deleted
                                        quit = if_aunit_constants=>quit-no ).

    SELECT * FROM mpa_asset_data INTO TABLE @DATA(lt_data) WHERE file_id = '12345'.

    cl_abap_unit_assert=>assert_initial( act  = lt_data
                                         quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD get_instance.
    DATA(lo) = zcl_xlsx_parse_util=>get_instance( ).
    cl_abap_unit_assert=>assert_not_initial( act   = lo
                                             msg   = 'Error in creating object of the class'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

    mo_mpa_parse_util->go_instance = mo_mpa_parse_util.
    lo = zcl_xlsx_parse_util=>get_instance( ).
    cl_abap_unit_assert=>assert_not_initial( act   = lo
                                         msg   = 'Error in object assignment'
                                         level = if_abap_unit_constant=>severity-medium
                                         quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD get_attr_from_node.

    DATA: lv_name TYPE string,
          lo_node TYPE REF TO if_ixml_node.

    DATA(lv_value) = mo_mpa_parse_util->mo_lcl_xlsx_parse->get_attr_from_node( EXPORTING iv_name = lv_name io_node = lo_node ).

    cl_abap_unit_assert=>assert_initial( act   = lv_value
                                         msg   = 'Error in adding additional format'
                                         level = if_abap_unit_constant=>severity-medium
                                         quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_lcl_private_method.

    DATA: lt_numfmtids  TYPE gty_t_numfmtids,
          lv_cell_value TYPE string.

    "When 1
    DATA(lv_formatted_value) = mo_lcl_parse_util->convert_cell_value_by_numfmt( EXPORTING iv_cell_value    = lv_cell_value
                                                                                          iv_number_format = ' '  ).
    "Then 1

    cl_abap_unit_assert=>assert_initial( act   = lv_formatted_value
                                         quit  = if_aunit_constants=>quit-no ).

    mo_lcl_parse_util->add_additional_format( CHANGING ct_numfmtids = lt_numfmtids ).

    cl_abap_unit_assert=>assert_not_initial( act   = lt_numfmtids
                                             msg   = 'Error in adding additional format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).


    lv_cell_value = 'DOC_TYPE'.
    READ TABLE lt_numfmtids INTO DATA(ls_numfmtids) WITH KEY id = 1.
    lv_formatted_value = mo_lcl_parse_util->convert_cell_value_by_numfmt(
                                                                       EXPORTING
                                                                         iv_cell_value    = lv_cell_value
                                                                         iv_number_format = ls_numfmtids-formatcode  ).

    cl_abap_unit_assert=>assert_not_initial( act   = lv_formatted_value
                                             msg   = 'Error'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

    lv_cell_value = '10052020'.
    READ TABLE lt_numfmtids INTO ls_numfmtids WITH KEY id = 14.
    lv_formatted_value = mo_lcl_parse_util->convert_cell_value_by_numfmt(
                                              EXPORTING
                                                iv_cell_value    = lv_cell_value
                                                iv_number_format = ls_numfmtids-formatcode  ).

    cl_abap_unit_assert=>assert_not_initial( act   = lv_formatted_value
                                             msg   = 'Error'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).


    lv_cell_value = '100520'.
    READ TABLE lt_numfmtids INTO ls_numfmtids WITH KEY id = 21.
    lv_formatted_value = mo_lcl_parse_util->convert_cell_value_by_numfmt(
                                              EXPORTING
                                                iv_cell_value    = lv_cell_value
                                                iv_number_format = ls_numfmtids-formatcode  ).

    cl_abap_unit_assert=>assert_not_initial( act   = lv_formatted_value
                                             msg   = 'Error'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).



    lv_cell_value = '1.1694'.
    READ TABLE lt_numfmtids INTO ls_numfmtids WITH KEY id = 18.
    lv_formatted_value = mo_lcl_parse_util->convert_cell_value_by_numfmt(
                                              EXPORTING
                                                iv_cell_value    = lv_cell_value
                                                iv_number_format = ls_numfmtids-formatcode  ).

    cl_abap_unit_assert=>assert_not_initial( act   = lv_formatted_value
                                             msg   = 'Error'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).


    mo_lcl_parse_util->convert_ser_val_to_date_time(
            EXPORTING
              iv_serial_value_string = '2020'
            IMPORTING
              ev_date                = DATA(lv_date) ).

    cl_abap_unit_assert=>assert_not_initial( act   = lv_date
                                             msg   = 'Error'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).


    mo_lcl_parse_util->convert_ser_val_to_date_time(
            EXPORTING
              iv_serial_value_string = 'abc'
            IMPORTING
              ev_date                = lv_date ).

    cl_abap_unit_assert=>assert_equals( act  = lv_date
                                        exp  = '19050712'
                                        quit = if_aunit_constants=>quit-no ).


    mo_lcl_parse_util->convert_ser_val_to_date_time(
            EXPORTING
              iv_serial_value_string = '0'
            IMPORTING
              ev_date                = lv_date ).

    cl_abap_unit_assert=>assert_equals( act  = lv_date
                                        exp  = '19050712'
                                        quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD fill_table_file_data.

    DATA : ls_line           TYPE mpa_s_index_value_pair,
           ls_cell           TYPE mpa_s_index_value_pair,
           lt_mpa_asset_data TYPE mpa_t_index_value_pair.

    FIELD-SYMBOLS: <ft_linedata> TYPE mpa_t_index_value_pair,
                   <fs_celldata> TYPE any.

    CASE iv_scen_type.
      WHEN if_mpa_output=>gc_mpa_scen-transfer.

        "First record
        ls_line-index = gc_one.

        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        ls_cell-index = 'A'.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        <fs_celldata> = 'Asset Mass Transfer'.
        IF iv_change_title_name IS NOT INITIAL.
          <fs_celldata> = 'Mass Transfer'.
        ENDIF.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Second record
        ls_line-index = '2'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = 'BLART'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = 'BLDAT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'BUDAT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'BZDAT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'SGTXT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'F'.
        <fs_celldata> = 'MONAT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'G'.
        <fs_celldata> = 'WWERT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'H'.
        <fs_celldata> = 'BUKRS'.
        IF iv_unwanted_field IS NOT INITIAL.
          <fs_celldata> = 'BUKRS1'.
        ENDIF.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'I'.
        <fs_celldata> = 'ANLN1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'J'.
        <fs_celldata> = 'ANLN2'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'K'.
        <fs_celldata> = 'ACC_PRINCIPLE'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'L'.
        <fs_celldata> = 'AFABER'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'M'.
        <fs_celldata> = 'PBUKRS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'N'.
        <fs_celldata> = 'PANL1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'O'.
        <fs_celldata> = 'PANL2'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'P'.
        <fs_celldata> = 'ANLKL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'Q'.
        <fs_celldata> = 'KOSTL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'R'.
        <fs_celldata> = 'TEXT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'S'.
        <fs_celldata> = 'TRAVA'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'T'.
        <fs_celldata> = 'ANBTR'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'U'.
        <fs_celldata> = 'WAERS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'V'.
        <fs_celldata> = 'MENGE'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'W'.
        <fs_celldata> = 'MEINS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'X'.
        <fs_celldata> = 'PROZS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'Y'.
        <fs_celldata> = 'XANEU'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'Z'.
        <fs_celldata> = 'RECID'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'AA'.
        <fs_celldata> = 'XBLNR'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'AB'.
        <fs_celldata> = 'DZUONR'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Third record
        IF iv_no_data IS INITIAL.
          IF iv_clear_techdesc IS INITIAL.
            ls_line-index = '3'.
            CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
            ASSIGN ls_line-value->* TO <ft_linedata>.
            CREATE DATA ls_cell-value TYPE string.
            ASSIGN ls_cell-value->* TO <fs_celldata>.
            ls_cell-index = 'A'.
            <fs_celldata> = ' '.
            INSERT ls_cell INTO TABLE <ft_linedata>.
            INSERT ls_line INTO TABLE lt_mpa_asset_data.
          ENDIF.
          "Fourth record
          ls_line-index = '4'.
          IF  iv_increase_row IS NOT INITIAL.
            ls_line-index = '5'.
          ENDIF.
          CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
          ASSIGN ls_line-value->* TO <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'B'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'C'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'D'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'G'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'H'.
          <fs_celldata> = gc_cc.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'I'.
          <fs_celldata> = '10000000073'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'J'.
          <fs_celldata> = ' '.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'N'.
          <fs_celldata> = '10000000074'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'O'.
          <fs_celldata> = ' '.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'T'.
          <fs_celldata> = gc_one.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'U'.
          <fs_celldata> = 'EUR'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'V'.
          <fs_celldata> = gc_one.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'W'.
          <fs_celldata> = 'KG'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'Y'.
          <fs_celldata> = 'X'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          INSERT ls_line INTO TABLE lt_mpa_asset_data.

          "Fifth record
          ls_line-index = '5'.
          CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
          ASSIGN ls_line-value->* TO <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'B'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'C'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'D'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'G'.
          <fs_celldata> = sy-datum.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'H'.
          <fs_celldata> = '0002'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'I'.
          <fs_celldata> = '10000073'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'J'.
          <fs_celldata> = ' '.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'N'.
          <fs_celldata> = '10000000074'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'O'.
          <fs_celldata> = ' '.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'T'.
          <fs_celldata> = gc_one.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'U'.
          <fs_celldata> = 'EUR'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'V'.
          <fs_celldata> = gc_one.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'W'.
          <fs_celldata> = 'KG'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          ls_cell-index = 'Y'.
          <fs_celldata> = 'X'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          INSERT ls_line INTO TABLE lt_mpa_asset_data.
        ENDIF.

      WHEN if_mpa_output=>gc_mpa_scen-create.

        "First record
        ls_line-index = gc_one.

        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        ls_cell-index = 'A'.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        <fs_celldata> = 'Asset Mass Create'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Second record
        ls_line-index = '2'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = 'BUKRS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = 'ANLKL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'TXA50_ANLT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'TXA50_MORE'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'MEINS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'F'.
        <fs_celldata> = 'KOSTL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Third record
        ls_line-index = '3'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = ' '.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Fourth record
        ls_line-index = '4'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = gc_cc.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = '1100'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'KG'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'F'.
        <fs_celldata> = 'CC_JV00356'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.

        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

      WHEN if_mpa_output=>gc_mpa_scen-change.

        IF iv_clear_title IS INITIAL.
          "First record
          ls_line-index = gc_one.
          CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
          ASSIGN ls_line-value->* TO <ft_linedata>.
          ls_cell-index = 'A'.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          <fs_celldata> = 'Asset Mass Change'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          INSERT ls_line INTO TABLE lt_mpa_asset_data.
        ENDIF.


        ls_line-index = '2'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        ls_cell-index = 'A'.
        IF iv_comment_cell_change IS NOT INITIAL.
          ls_cell-index = 'B'.
        ENDIF.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        <fs_celldata> = '// Comments'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        ls_line-index = '3'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        IF iv_error_data IS INITIAL.
          ls_cell-index = 'A'.
          <fs_celldata> = 'BUKRS'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
        ENDIF.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = 'ANLN1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'TXA50_ANLT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'TXA50_MORE'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'MEINS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'F'.
        <fs_celldata> = 'KOSTL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Third record
        ls_line-index = '4'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = ' '.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.
        .
        "Fourth record
        ls_line-index = '5'.
        IF iv_error_data IS NOT INITIAL.
          ls_line-index = '1'.
        ENDIF..
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = gc_cc.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = '1100'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'KG'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        IF iv_error_data IS INITIAL.
          ls_cell-index = 'F'.
          <fs_celldata> = 'CC_JV00356'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
        ENDIF.

        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

      WHEN if_mpa_output=>gc_mpa_scen-adjustment.

        IF iv_clear_title IS INITIAL.
          "First record
          ls_line-index = gc_one.
          CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
          ASSIGN ls_line-value->* TO <ft_linedata>.
          ls_cell-index = 'A'.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          <fs_celldata> = 'Asset Mass Adjustment'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          INSERT ls_line INTO TABLE lt_mpa_asset_data.
        ENDIF.


        ls_line-index = '2'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        ls_cell-index = 'A'.
        IF iv_comment_cell_change IS NOT INITIAL.
          ls_cell-index = 'B'.
        ENDIF.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        <fs_celldata> = '// Comments'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        ls_line-index = '3'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        IF iv_error_data IS INITIAL.
          ls_cell-index = 'A'.
          <fs_celldata> = 'BUKRS'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
        ENDIF.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = 'ANLN1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'TXA50_ANLT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'TXA50_MORE'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'MEINS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'F'.
        <fs_celldata> = 'KOSTL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Third record
        ls_line-index = '4'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = ' '.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.
        .
        "Fourth record
        ls_line-index = '5'.
        IF iv_error_data IS NOT INITIAL.
          ls_line-index = '1'.
        ENDIF..
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = gc_cc.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = '1100'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'KG'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        IF iv_error_data IS INITIAL.
          ls_cell-index = 'F'.
          <fs_celldata> = 'CC_JV00356'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
        ENDIF.

        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

      WHEN if_mpa_output=>gc_mpa_scen-retirement.

        IF iv_clear_title IS INITIAL.
          "First record
          ls_line-index = gc_one.
          CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
          ASSIGN ls_line-value->* TO <ft_linedata>.
          ls_cell-index = 'A'.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
          <fs_celldata> = 'Asset Mass Retirement'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          INSERT ls_line INTO TABLE lt_mpa_asset_data.
        ENDIF.


        ls_line-index = '2'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        ls_cell-index = 'A'.
        IF iv_comment_cell_change IS NOT INITIAL.
          ls_cell-index = 'B'.
        ENDIF.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        <fs_celldata> = '// Comments'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        ls_line-index = '3'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        IF iv_error_data IS INITIAL.
          ls_cell-index = 'A'.
          <fs_celldata> = 'BUKRS'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
        ENDIF.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = 'ANLN1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'TXA50_ANLT'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'TXA50_MORE'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'MEINS'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'F'.
        <fs_celldata> = 'KOSTL'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

        "Third record
        ls_line-index = '4'.
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = ' '.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.
        .
        "Fourth record
        ls_line-index = '5'.
        IF iv_error_data IS NOT INITIAL.
          ls_line-index = '1'.
        ENDIF..
        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'A'.
        <fs_celldata> = gc_cc.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'B'.
        <fs_celldata> = '1100'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'C'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'D'.
        <fs_celldata> = 'Test'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'E'.
        <fs_celldata> = 'KG'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        IF iv_error_data IS INITIAL.
          ls_cell-index = 'F'.
          <fs_celldata> = 'CC_JV00356'.
          INSERT ls_cell INTO TABLE <ft_linedata>.
          CREATE DATA ls_cell-value TYPE string.
          ASSIGN ls_cell-value->* TO <fs_celldata>.
        ENDIF.

        INSERT ls_cell INTO TABLE <ft_linedata>.
        INSERT ls_line INTO TABLE lt_mpa_asset_data.

    ENDCASE.

    et_table = lt_mpa_asset_data.

  ENDMETHOD.


  METHOD fill_xstring_file_data.

    DATA: lv_file1  TYPE string,
          lv_file80 TYPE string.
    lv_file1  = '504B030414000600080000002100413782CF6E01000004050000130008025B436F6E74656E745F54797065735D2E786D6C20A2040228A00002000000000000000000000000000000000000'.
    lv_file80 = '000000000C000C0026030000202B00000000'.

    CONCATENATE lv_file1
                lv_file80 INTO DATA(lv_file).
    MOVE lv_file TO ex_file.
  ENDMETHOD.


  METHOD fill_mock_tables.

    DATA: lt_dd04t TYPE TABLE OF dd04t.   "R/3 DD: Data element texts

    "R/3 DD: Data element texts
    lt_dd04t = VALUE #(
                        ( rollname = 'BUKRS' ddlanguage = 'E' as4local = 'A' ddtext = 'Company Code' reptext = 'CoCd' scrtext_s = 'CoCode' scrtext_m = 'Company Code' scrtext_l = 'Company Code' )
                        ( rollname = 'MONAT' ddlanguage = 'E' as4local = 'A' ddtext = 'Fiscal Period' scrtext_s = 'Period' scrtext_m = 'Period' scrtext_l = 'Posting Period' )
                        ( rollname = 'BF_PANL1' ddlanguage = 'E' as4local = 'A' ddtext = 'Main Number Partner Asset (Intercompany Transfer)' scrtext_s = 'PrtnrAsset' scrtext_m = 'Partner Asset' scrtext_l = 'Partner Asset' )
                        ( rollname = 'BF_PANL2' ddlanguage = 'E' as4local = 'A' ddtext = 'Partner Asset Subnumber (Intercompany Transfer)' scrtext_s = 'P. Sub-No.' scrtext_m = 'Partner Sub-No.' scrtext_l = 'Partner Subnumber' )
                        ( rollname = 'BF_TRANSVAR' ddlanguage = 'E' as4local = 'A' ddtext = 'Transfer Variant for Intercompany Asset Transfers' scrtext_s = 'Variant' scrtext_m = 'Transfer Var.' scrtext_l = 'Transfer Variant' )
                        ( rollname = 'TXA50_ANLT' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Description' reptext = 'Asset Description' scrtext_s = 'Descript.' scrtext_m = 'Description' scrtext_l = 'Description' )
                        ( rollname = 'WAERS' ddlanguage = 'E' as4local = 'A' ddtext = 'Currency Key' reptext = 'Crcy' scrtext_s = 'Currency' scrtext_m = 'Currency' scrtext_l = 'Currency' )
                        ( rollname = 'SGTXT' ddlanguage = 'E' as4local = 'A' ddtext = 'Item Text' reptext = 'Text' scrtext_s = 'Text' scrtext_m = 'Text' scrtext_l = 'Text' )
                        ( rollname = 'MENGE_D' ddlanguage = 'E' as4local = 'A' ddtext = 'Quantity' reptext = 'Quantity' scrtext_s = 'Quantity' scrtext_m = 'Quantity' scrtext_l = 'Quantity' )
                        ( rollname = 'BLART' ddlanguage = 'E' as4local = 'A' ddtext = 'Document Type' reptext = 'Doc. Type' scrtext_s = 'Type' scrtext_m = 'Document Type' scrtext_l = 'Document Type' )
                        ( rollname = 'WERKS_D' ddlanguage = 'E' as4local = 'A' ddtext = 'Plant' reptext = 'Plnt' scrtext_s = 'Plant' scrtext_m = 'Plant' scrtext_l = 'Plant' )
                        ( rollname = 'XBLNR' ddlanguage = 'E' as4local = 'A' ddtext = 'Reference Document Number' reptext = 'Reference' scrtext_s = 'Reference' scrtext_m = 'Reference' scrtext_l = 'Reference' )
                        ( rollname = 'BUDAT' ddlanguage = 'E' as4local = 'A' ddtext = 'Posting Date in the Document' reptext = 'Pstng Date' scrtext_s = 'Pstng Date' scrtext_m = 'Posting Date' scrtext_l = 'Posting Date' )
                        ( rollname = 'BF_ANBTR' ddlanguage = 'E' as4local = 'A' ddtext = 'Amount Posted' reptext = 'Amount Posted' scrtext_s = 'Amount' scrtext_m = 'Amount Posted' scrtext_l = 'Amount Posted' )
                        ( rollname = 'MEINS' ddlanguage = 'E' as4local = 'A' ddtext = 'Base Unit of Measure' reptext = 'BUn' scrtext_s = 'Unit' scrtext_m = 'Base Unit' scrtext_l = 'Base Unit of Measure' )
                        ( rollname = 'BF_ANLKL' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset class' reptext = 'Class' scrtext_s = 'Class' scrtext_m = 'Asset class' scrtext_l = 'Asset class' )
                        ( rollname = 'AFABE_POST' ddlanguage = 'E' as4local = 'A' ddtext = 'Posting Depreciation Area' reptext = 'Ar.' scrtext_s = 'Area' scrtext_m = 'Deprec. Area' scrtext_l = 'Depreciation Area' )
                        ( rollname = 'PS_PSP_PNR' ddlanguage = 'E' as4local = 'A' ddtext = 'Work Breakdown Structure Element (WBS Element)' reptext = 'WBS Element' scrtext_s = 'WBS Elem.' scrtext_m = 'WBS Element' scrtext_l = 'WBS Element' )
                       ( rollname = 'XANEU' ddlanguage = 'E' as4local = 'A' ddtext = 'Indicator: Transaction Relates to Curr.-Yr Acquisition' reptext = 'New' scrtext_s = 'Cur.Yr.Acq' scrtext_m = 'Curr.Yr.Acquis.' scrtext_l = 'Curr.-Year Acquisition' )
                       ( rollname = 'BLDAT' ddlanguage = 'E' as4local = 'A' ddtext = 'Document Date in Document' reptext = 'Doc. Date' scrtext_s = 'Doc. Date' scrtext_m = 'Document Date' scrtext_l = 'Document Date' )
                       ( rollname = 'JV_RECIND' ddlanguage = 'E' as4local = 'A' ddtext = 'Recovery Indicator' reptext = 'RI' scrtext_s = 'Rec.Indic.' scrtext_m = 'Recovery Indic.' scrtext_l = 'Recovery Indicator' )
                       ( rollname = 'BF_ANLN1' ddlanguage = 'E' as4local = 'A' ddtext = 'Main Asset Number' reptext = 'Asset' scrtext_s = 'Asset' scrtext_m = 'Asset' scrtext_l = 'Asset' )
                       ( rollname = 'DZUONR' ddlanguage = 'E' as4local = 'A' ddtext = 'Assignment number' reptext = 'Assignment' scrtext_s = 'Assign.' scrtext_m = 'Assignment' scrtext_l = 'Assignment' )
                       ( rollname = 'KOSTL' ddlanguage = 'E' as4local = 'A' ddtext = 'Cost Center' reptext = 'Cost Ctr' scrtext_s = 'Cost Ctr' scrtext_m = 'Cost Center' scrtext_l = 'Cost Center' )
                       ( rollname = 'BF_ANLN2' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Subnumber' reptext = 'SNo.' scrtext_s = 'Sub-number' scrtext_m = 'Sub-number' scrtext_l = 'Sub-number' )
                      ( rollname = 'ACCOUNTING_PRINCIPLE' ddlanguage = 'E' as4local = 'A' ddtext = 'Accounting Principle' reptext = 'Accounting Principle' scrtext_s = 'Acc.Princ.' scrtext_m = 'Accounting Principle' scrtext_l = 'Accounting Principle' )
                      ( rollname = 'BAPI_MTYPE' ddlanguage = 'E' as4local = 'A' ddtext = 'Message type: S Success, E Error, W Warning, I Info, A Abort' reptext = 'MsgType' scrtext_s = 'Msg type' scrtext_m = 'Message type' scrtext_l = 'Message type' )
                      ( rollname = 'INVNR_ANLA' ddlanguage = 'E' as4local = 'A' ddtext = 'Inventory Number' reptext = 'Inventory Number' scrtext_s = 'Invent. No' scrtext_m = 'Inventory No.' scrtext_l = 'Inventory Number' )
                      ( rollname = 'WWERT_D' ddlanguage = 'E' as4local = 'A' ddtext = 'Translation date' reptext = 'TranslDate' scrtext_s = 'TranslDate' scrtext_m = 'Translation dte' scrtext_l = 'Translation date' )
                      ( rollname = 'BF_PBUKR' ddlanguage = 'E' as4local = 'A' ddtext = 'Partner Company Code' reptext = 'PCCo' scrtext_s = 'PartCoCd' scrtext_m = 'PartnerCoCode' scrtext_l = 'Partner Comp. Code' )
                      ( rollname = 'BZDAT' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Value Date' reptext = 'AssetValDate' scrtext_s = 'AsstValDat' scrtext_m = 'Asset Val. Date' scrtext_l = 'Asset Value Date' )
                      ( rollname = 'BF_PROZS' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Retirement: Percentage Rate' reptext = '% Rate' scrtext_s = 'Perc.Rate' scrtext_m = 'Percentage Rate' scrtext_l = 'Percentage Rate' )

                      ( rollname = 'XNEU_AM' ddlanguage = 'E' as4local = 'A' ddtext = 'Indicator: Asset purchased new' scrtext_l = 'Asset purchased new' )
                      ( rollname = 'BF_AM_LAND1' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Country of Origin' scrtext_s = 'Orig Cntry' scrtext_m = 'Ctry of Origin' scrtext_l = 'Country of Origin' )
                      ( rollname = 'BF_TXA50' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Description' reptext = 'Asset Description' scrtext_s = 'Descript.' scrtext_m = 'Description' scrtext_l = 'Description' )
                      ( rollname = 'STORT' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset location' reptext = 'Location' scrtext_s = 'Location' scrtext_m = 'Location' scrtext_l = 'Location' )
                      ( rollname = 'BF_AFASL' ddlanguage = 'E' as4local = 'A' ddtext = 'Depreciation Key' reptext = 'DepKy' scrtext_s = 'Key' scrtext_m = 'Dep. Key' scrtext_l = 'Depreciation Key' )
                      ( rollname = 'BF_AFABE_D' ddlanguage = 'E' as4local = 'A' ddtext = 'Real Depreciation Area' reptext = 'Ar.' scrtext_s = 'Area' scrtext_m = 'Deprec. Area' scrtext_l = 'Depreciation Area' )
                      ( rollname = 'BF_AM_LIFNR' ddlanguage = 'E' as4local = 'A' ddtext = 'Account Number of Supplier (Other Key Word)' reptext = 'Supplier' scrtext_s = 'Supplier' scrtext_m = 'Supplier' scrtext_l = 'Supplier' )
                      ( rollname = 'BF_NDPER' ddlanguage = 'E' as4local = 'A' ddtext = 'Planned Useful Life in Periods' reptext = 'Per' scrtext_s = '/' scrtext_m = 'Periods' scrtext_l = 'Usef.Life in Periods' )
                      ( rollname = 'BF_NDJAR' ddlanguage = 'E' as4local = 'A' ddtext = 'Planned Useful Life in Years' reptext = 'Use' scrtext_s = 'Usefl.Life' scrtext_m = 'Useful Life' scrtext_l = 'Useful Life' )
                      ( rollname = 'BF_AM_SERNR' ddlanguage = 'E' as4local = 'A' ddtext = 'Serial Number' reptext = 'Serial Number' scrtext_s = 'Serial No.' scrtext_m = 'Serial Number' scrtext_l = 'Serial Number' )
                      ( rollname = 'FB_SEGMENT' ddlanguage = 'E' as4local = 'A' ddtext = 'Segment for Segmental Reporting' reptext = 'Segment' scrtext_s = 'Segment' scrtext_m = 'Segment' scrtext_l = 'Segment' )
                      ( rollname = 'FINS_LEDGER' ddlanguage = 'E' as4local = 'A' ddtext = 'Ledger in General Ledger Accounting' reptext = 'Ld' scrtext_s = 'Ledger' scrtext_m = 'Ledger' scrtext_l = 'Ledger' )
                      ( rollname = 'BF_RAUMNR' ddlanguage = 'E' as4local = 'A' ddtext = 'Room' reptext = 'Room' scrtext_s = 'Room' scrtext_m = 'Room' scrtext_l = 'Room' )
                      ( rollname = 'BF_TYPBZ_ANLA' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset Type Name' reptext = 'Type Name' scrtext_s = 'Type Name' scrtext_m = 'Type Name' scrtext_l = 'Type Name' )
                      ( rollname = 'PRCTR' ddlanguage = 'E' as4local = 'A' ddtext = 'Profit Center' reptext = 'Profit Ctr' scrtext_s = 'Profit Ctr' scrtext_m = 'Profit Center' scrtext_l = 'Profit Center' )
                      ( rollname = 'BF_INKEN' ddlanguage = 'E' as4local = 'A' ddtext = 'Inventory Indicator' reptext = 'II' scrtext_s = 'Invent.Ind' scrtext_m = 'Inventory Ind.' scrtext_l = 'Include Asset' )
                      ( rollname = 'BF_URWRT' ddlanguage = 'E' as4local = 'A' ddtext = 'Original Acquisition Value' reptext = 'Original Value' scrtext_s = 'Orig.Value' scrtext_m = 'Original Value' scrtext_l = 'Original Value' )
                      ( rollname = 'BF_RASSC' ddlanguage = 'E' as4local = 'A' ddtext = 'Company ID of trading partner' reptext = 'Tr.prt' scrtext_s = 'Tradg part' scrtext_m = 'Trading partner' scrtext_l = 'Trading partner' )
                      ( rollname = 'BF_INVNR_ANLA' ddlanguage = 'E' as4local = 'A' ddtext = 'Inventory number' reptext = 'Inventory number' scrtext_s = 'Invent. no' scrtext_m = 'Inventory no.' scrtext_l = 'Inventory number' )
                      ( rollname = 'BF_AKTIVD' ddlanguage = 'E' as4local = 'A' ddtext = 'Asset capitalization date' reptext = 'Cap.date' scrtext_s = 'Cap.date' scrtext_m = 'Activated On' scrtext_l = 'Activated On' )
                      ( rollname = 'BF_IVDAT_ANLA' ddlanguage = 'E' as4local = 'A' ddtext = 'Last Inventory Date' reptext = 'Inv.Date' scrtext_s = 'Last Inv.' scrtext_m = 'Last Inventory' scrtext_l = 'Last Inventory On' )
                      ( rollname = 'BF_HERST' ddlanguage = 'E' as4local = 'A' ddtext = 'Manufacturer of Asset' reptext = 'Manufacturer of Asset' scrtext_s = 'Mfr' scrtext_m = 'Manufacturer' scrtext_l = 'Manufacturer' )
                      ( rollname = 'BF_ANTEI' ddlanguage = 'E' as4local = 'A' ddtext = 'In-House Production Percentage' reptext = 'In-H.Pro' scrtext_s = 'In-H.Prod.' scrtext_m = 'In-H.Prod.Perc.' scrtext_l = 'In-House Prod.Perc.' )
                      ( rollname = 'TXJCD' ddlanguage = 'E' as4local = 'A' ddtext = 'Tax Jurisdiction' reptext = 'Tax Jur.' scrtext_s = 'Tax Jur.' scrtext_m = 'Tax Jur.' scrtext_l = 'Tax Jurisdiction' )
                      ( rollname = 'BAPI1022_POSNR_EXT2' ddlanguage = 'E' as4local = 'A' ddtext = 'WBS Element - External Key'  scrtext_s = 'WBS Elem.' scrtext_m = 'WBS Element' scrtext_l = 'WBSElement for Costs' )
                      ( rollname = 'XNACH_ANLA' ddlanguage = 'E' as4local = 'A' ddtext = 'Post-capitalization of asset' reptext = 'Post-capitalization' scrtext_s = 'Post-cap.' scrtext_m = 'Post-capital.' scrtext_l = 'Post-capitalization' )
                      ( rollname = 'BF_INVZU_ANLA' ddlanguage = 'E' as4local = 'A' ddtext = 'Supplementary Inventory Specifications' reptext = 'Inventory Note' scrtext_s = 'Inventory' scrtext_m = 'Inventory Note' scrtext_l = 'Inventory Note' )
                      ( rollname = 'BF_AFABG' ddlanguage = 'E' as4local = 'A' ddtext = 'Depreciation Calculation Start Date' reptext = 'ODep.Start' scrtext_s = 'Ord. Depr.' scrtext_m = 'Ordinary Dep.' scrtext_l = 'Ordinary Depreciat.' )
                     ( rollname = 'BAPI1022_TXA50_MORE' ddlanguage = 'E' as4local = 'A' ddtext = 'Additional Asset Description' reptext = 'Asset Description 2' scrtext_s = 'Descrip. 2' scrtext_m = 'Description 2' scrtext_l = 'Additional Description' )
                   ).

    go_osql_environment->insert_test_data( lt_dd04t ).

    DATA: lt_dd07v TYPE TABLE OF dd07v.   "Generated Table for View

    "Generated Table for View
    lt_dd07v = VALUE #(
                        ( domname = 'MPA_TEMPLATE_TYPE' valpos = '0002' ddlanguage = 'E' domvalue_l = 'CR' ddtext = 'Mass Creation' )
                        ( domname = 'MPA_TEMPLATE_TYPE' valpos = '0003' ddlanguage = 'E' domvalue_l = 'CH' ddtext = 'Mass Change' )
                        ( domname = 'MPA_TEMPLATE_TYPE' valpos = '0001' ddlanguage = 'E' domvalue_l = 'MT' ddtext = 'Mass Transfer' )
                        ( domname = 'MPA_TEMPLATE_TYPE' valpos = '0004' ddlanguage = 'E' domvalue_l = 'RT' ddtext = 'Mass Retirement' )
                      ).

    go_osql_environment->insert_test_data( lt_dd07v ).


    DATA: lt_mpa_asset_data TYPE TABLE OF mpa_asset_data.   "File content for MPA

    "File content for MPA
    lt_mpa_asset_data = VALUE #(
                                (
                                  mandt = sy-mandt
                                  ernam = 'C5232603'
                                  erdat = '20200325'
                                  erzet = '163151'
                                  file_id = '0000000168'
                                  file_name = 'AssetMassTransfer_Template.XLSX'
                                  scen_type = 'MT'
                                  file_status = '1'
                                )
                                (
                                  mandt = sy-mandt
                                  ernam = 'C5232604'
                                  erdat = '20200325'
                                  erzet = '163151'
                                  file_id = '0000000169'
                                  file_name = 'AssetMassCreate_Template.XLSX'
                                  scen_type = 'MT'
                                  file_status = '1'
                                )
                              ).

    go_osql_environment->insert_test_data( lt_mpa_asset_data ).

    DATA: lt_dd04l TYPE TABLE OF dd04l.   "Data elements

    "Data elements
    lt_dd04l = VALUE #(
                        ( rollname = 'BF_ANLKL' as4local = 'A' domname = 'BF_ANLKL' memoryid = 'ANK' logflag = 'X' headlen = '08' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '19980218' as4time = '064343'
                          dtelmaster = 'D' deffdname = 'ASSETCLASS' datatype = 'CHAR' leng = '000008' outputlen = '000008' convexit = 'ALPHA' entitytab = 'ANKA' refkind = 'D' )
                        ( rollname = 'TXA50_ANLT' as4local = 'A' domname = 'TEXT50' logflag = 'X' headlen = '50' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20130815' as4time = '143728' dtelmaster = 'D'
                          deffdname = 'DESCRIPT' datatype = 'CHAR' leng = '000050' outputlen = '000050' lowercase = 'X' refkind = 'D' )
                    ( rollname = 'SGTXT' as4local = 'A' domname = 'TEXT50' logflag = 'X' headlen = '50' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'SAP' as4user = 'SAP' as4date = '20130815' as4time = '143700' dtelmaster = 'D' deffdname =
                      'ITEM_TEXT' datatype = 'CHAR' leng = '000050' outputlen = '000050' lowercase = 'X' refkind = 'D' )
                    ( rollname = 'BUKRS' as4local = 'A' domname = 'BUKRS' memoryid = 'BUK' logflag = 'X' headlen = '04' scrlen1 = '06' scrlen2 = '15' scrlen3 = '15' applclass = 'FB' as4user = 'LISSITSYNA' as4date = '20200213' as4time = '142320'
                      dtelmaster = 'D' shlpname = 'C_T001' shlpfield = 'BUKRS' deffdname = 'COMP_CODE' datatype = 'CHAR' leng = '000004' outputlen = '000004' entitytab = 'T001' refkind = 'D' )
                    ( rollname = 'BF_ANLN1' as4local = 'A' domname = 'BF_ANLN1' memoryid = 'AN1' logflag = 'X' headlen = '12' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20120404' as4time = '183839'
                      dtelmaster = 'D' deffdname = 'ASSETMAINO' datatype = 'CHAR' leng = '000012' outputlen = '000012' convexit = 'ALPHA' refkind = 'D' )
                    ( rollname = 'KOSTL' as4local = 'A' domname = 'KOSTL' memoryid = 'KOS' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'KS' as4user = 'LISSITSYNA' as4date = '20170807' as4time = '232918'
                      dtelmaster = 'D' deffdname = 'COSTCENTER' datatype = 'CHAR' leng = '000010' outputlen = '000010' convexit = 'ALPHA' entitytab = 'CSKS' refkind = 'D' )
                    ( rollname = 'MONAT' as4local = 'A' domname = 'MONAT' logflag = 'X' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'MG' as4user = 'SAP' as4date = '19980218' as4time = '073249' dtelmaster = 'D' deffdname = 'FIS_PERIOD'
                      datatype = 'NUMC' leng = '000002' outputlen = '000002' refkind = 'D' )
                    ( rollname = 'INVNR_ANLA' as4local = 'A' domname = 'INVNR_ANLA' logflag = 'X' headlen = '25' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '161501' dtelmaster = 'D'
                      deffdname = 'INVENT_NO' datatype = 'CHAR' leng = '000025' outputlen = '000025' refkind = 'D' )
                    ( rollname = 'XBLNR' as4local = 'A' domname = 'XBLNR' logflag = 'X' headlen = '16' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '19980218' as4time = '080231' dtelmaster = 'D' deffdname =
                      'REF_DOC_NO' datatype = 'CHAR' leng = '000016' outputlen = '000016' refkind = 'D' )
                    ( rollname = 'PS_PSP_PNR' as4local = 'A' domname = 'PS_POSNR' logflag = 'X' headlen = '24' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'LISSITSYNA' as4date = '20170807' as4time = '232918' dtelmaster = 'D' deffdname =
                      'WBS_ELEM' datatype = 'NUMC' leng = '000008' outputlen = '000024' convexit = 'ABPSP' entitytab = 'PRPS' refkind = 'D' )
                    ( rollname = 'BLART' as4local = 'A' domname = 'BLART' memoryid = 'BAR' logflag = 'X' headlen = '10' scrlen1 = '06' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '19990526' as4time = '085809' dtelmaster =
                      'D' deffdname = 'DOC_TYPE' datatype = 'CHAR' leng = '000002' outputlen = '000002' entitytab = 'T003' refkind = 'D' )
                    ( rollname = 'XANEU' as4local = 'A' domname = 'XFELD' headlen = '03' scrlen1 = '10' scrlen2 = '15' scrlen3 = '24' as4user = 'SAP' as4date = '20040825' as4time = '183908' dtelmaster = 'D' deffdname = 'NEW_ACQ_IN' datatype = 'CHAR'
                      leng = '000001' outputlen = '000001' valexi = 'X' refkind = 'D' )
                    ( rollname = 'WWERT_D' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '19980218' as4time = '080157' dtelmaster = 'D' deffdname =
                      'TRANS_DATE' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BZDAT' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '12' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AB' as4user = 'SAP' as4date = '20010607' as4time = '155518' dtelmaster = 'D' deffdname =
                      'ASVAL_DATE' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BUDAT' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '19980713' as4time = '175204' dtelmaster = 'D' deffdname =
                      'PSTNG_DATE' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BLDAT' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '19980713' as4time = '175202' dtelmaster = 'D' deffdname =
                      'DOC_DATE' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BF_ANLN2' as4local = 'A' domname = 'BF_ANLN2' memoryid = 'AN2' logflag = 'X' headlen = '04' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '19980218' as4time = '064343'
                      dtelmaster = 'D' deffdname = 'ASSETSUBNO' datatype = 'CHAR' leng = '000004' outputlen = '000004' convexit = 'ALPHA' refkind = 'D' )
                    ( rollname = 'BF_ANBTR' as4local = 'A' domname = 'BAPICURR' logflag = 'X' headlen = '16' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'ABB' as4user = 'SAP' as4date = '20171129' as4time = '134618' dtelmaster = 'D'
                      deffdname = 'AMOUNT' datatype = 'DEC' leng = '000023' decimals = '000004' outputlen = '000030' refkind = 'D' )
                    ( rollname = 'BF_PANL1' as4local = 'A' domname = 'BF_ANLN1' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'SAP' as4date = '20010607' as4time = '155506' dtelmaster = 'D' deffdname = 'PART_ASSET' datatype = 'CHAR' leng =
                      '000012' outputlen = '000012' convexit = 'ALPHA' refkind = 'D' )
                    ( rollname = 'BF_PANL2' as4local = 'A' domname = 'BF_ANLN2' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'SAP' as4date = '20010607' as4time = '155506' dtelmaster = 'D' deffdname = 'PART_SUBNO' datatype = 'CHAR' leng =
                      '000004' outputlen = '000004' convexit = 'ALPHA' refkind = 'D' )
                    ( rollname = 'BF_TRANSVAR' as4local = 'A' domname = 'BF_TRANSVAR' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'SAP' as4date = '20010607' as4time = '155507' dtelmaster = 'D' deffdname = 'TRANSVAR' datatype = 'CHAR' leng =
                      '000004' outputlen = '000004' convexit = 'ALPHA' refkind = 'D' )
                    ( rollname = 'WAERS' as4local = 'A' domname = 'WAERS' memoryid = 'FWS' logflag = 'X' headlen = '05' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '19980218' as4time = '080102' dtelmaster =
                      'D' deffdname = 'CURRENCY' datatype = 'CUKY' leng = '000005' outputlen = '000005' entitytab = 'TCURC' refkind = 'D' )
                    ( rollname = 'MEINS' as4local = 'A' domname = 'MEINS' logflag = 'X' headlen = '03' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'MG' as4user = 'SAP' as4date = '20161114' as4time = '082417' dtelmaster = 'D' deffdname =
                      'BASE_UOM' datatype = 'UNIT' leng = '000003' outputlen = '000003' lowercase = 'X' convexit = 'CUNIT' entitytab = 'T006' refkind = 'D' )
                    ( rollname = 'DZUONR' as4local = 'A' domname = 'ZUONR' logflag = 'X' headlen = '18' scrlen1 = '10' scrlen2 = '12' scrlen3 = '15' applclass = 'FB' as4user = 'SAP' as4date = '20170427' as4time = '180329' dtelmaster = 'D' deffdname =
                      'ALLOC_NMBR' datatype = 'CHAR' leng = '000018' outputlen = '000018' lowercase = 'X' refkind = 'D' )
                    ( rollname = 'MENGE_D' as4local = 'A' domname = 'MENG13' logflag = 'X' headlen = '17' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'MG' as4user = 'LISSITSYNA' as4date = '20170807' as4time = '232918' dtelmaster = 'D'
                      deffdname = 'QUANTITY' datatype = 'QUAN' leng = '000013' decimals = '000003' outputlen = '000017' refkind = 'D' )
                    ( rollname = 'JV_RECIND' as4local = 'A' domname = 'JV_RECIND' logflag = 'X' headlen = '02' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FGJ' as4user = 'LISSITSYNA' as4date = '20170807' as4time = '232918' dtelmaster =
                      'D' deffdname = 'REC_IND' datatype = 'CHAR' leng = '000002' outputlen = '000002' convexit = 'ALPHA' entitytab = 'T8JJ' refkind = 'D' )
                    ( rollname = 'BF_PROZS' as4local = 'A' domname = 'BF_PRZ32' headlen = '06' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155506' dtelmaster = 'D' deffdname =
                      'PERC_RATE' datatype = 'DEC' leng = '000005' decimals = '000002' outputlen = '000006' refkind = 'D' )
                    ( rollname = 'ACCOUNTING_PRINCIPLE' as4local = 'A' domname = 'ACCOUNTING_PRINCIPLE' memoryid = 'ACCOUNTING_PRINCIPLE' logflag = 'X' headlen = '55' scrlen1 = '10' scrlen2 = '20' scrlen3 = '40' as4user = 'SWONKE' as4date = '20191125'
                      as4time = '211648' dtelmaster = 'D' deffdname = 'ACC_PRINCIPLE' datatype = 'CHAR' leng = '000004' outputlen = '000004' entitytab = 'TACC_PRINCIPLE' refkind = 'D' )
                    ( rollname = 'AFABE_POST' as4local = 'A' domname = 'AFABE' memoryid = 'AFB' headlen = '03' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AB' as4user = 'SAP' as4date = '20140226' as4time = '151651' dtelmaster = 'D'
                      deffdname = 'DEPR_AREA' datatype = 'NUMC' leng = '000002' outputlen = '000002' entitytab = 'T093' refkind = 'D' )
                    ( rollname = 'BAPI_MTYPE' as4local = 'A' domname = 'SYCHAR01' headlen = '07' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'SAP' as4date = '20001108' as4time = '191707' dtelmaster = 'D' deffdname = 'TYPE' datatype = 'CHAR'
                      leng = '000001' outputlen = '000001' refkind = 'D' )
                    ( rollname = 'BF_PBUKR' as4local = 'A' domname = 'BUKRS' headlen = '04' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'KA' as4user = 'SAP' as4date = '20010607' as4time = '155506' dtelmaster = 'D' deffdname = 'PART_COMCO'
                      datatype = 'CHAR' leng = '000004' outputlen = '000004' entitytab = 'T001' refkind = 'D' )

                    ( rollname = 'BF_ANTEI' as4local = 'A' domname = 'BF_PRZ32' logflag = 'X' headlen = '08' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155456' dtelmaster = 'D'
                      deffdname = 'PER_IH_PRO' datatype = 'DEC' leng = '000005' decimals = '000002' outputlen = '000006' refkind = 'D' )
                    ( rollname = 'BAPI1022_TXA50_MORE' as4local = 'A' domname = 'TEXT50' logflag = 'X' headlen = '50' scrlen1 = '10' scrlen2 = '15' scrlen3 = '23' applclass = 'AA' as4user = 'SAP' as4date = '20130815' as4time = '143412' dtelmaster =
                      'D' deffdname = 'DESCRIPT' datatype = 'CHAR' leng = '000050' outputlen = '000050' lowercase = 'X' refkind = 'D' )
                    ( rollname = 'BF_RAUMNR' as4local = 'A' domname = 'CHAR8' logflag = 'X' headlen = '08' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'INST' as4user = 'SAP' as4date = '20010607' as4time = '155506' dtelmaster = 'D'
                      deffdname = 'MAINTROOM' datatype = 'CHAR' leng = '000008' outputlen = '000008' refkind = 'D' )
                    ( rollname = 'BF_AM_SERNR' as4local = 'A' domname = 'BF_GERNR' memoryid = 'SER' logflag = 'X' headlen = '18' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'INST' as4user = 'SAP' as4date = '20010607' as4time = '155455'
                      dtelmaster = 'D' deffdname = 'SERIAL_NO' datatype = 'CHAR' leng = '000018' outputlen = '000018' convexit = 'ALPHA' refkind = 'D' )
                    ( rollname = 'BF_NDPER' as4local = 'A' domname = 'BF_PERAF' logflag = 'X' headlen = '03' scrlen1 = '01' scrlen2 = '10' scrlen3 = '20' applclass = 'ABA' as4user = 'SAP' as4date = '20010607' as4time = '155505' dtelmaster = 'D'
                      deffdname = 'USF_LIF_PER' datatype = 'NUMC' leng = '000003' outputlen = '000003' valexi = 'X' refkind = 'D' )
                    ( rollname = 'BF_AM_LIFNR' as4local = 'A' domname = 'LIFNR' memoryid = 'LIF' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'SAP' as4date = '20010607' as4time = '155455'
                      dtelmaster = 'D' deffdname = 'VENDOR_NO' datatype = 'CHAR' leng = '000010' outputlen = '000010' convexit = 'ALPHA' entitytab = 'LFA1' refkind = 'D' )
                    ( rollname = 'BF_INVNR_ANLA' as4local = 'A' domname = 'BF_INVNR_ANLA' logflag = 'X' headlen = '25' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '19980218' as4time = '064353' dtelmaster =
                      'D' deffdname = 'INVENT_NO' datatype = 'CHAR' leng = '000025' outputlen = '000025' refkind = 'D' )
                    ( rollname = 'FINS_LEDGER' as4local = 'A' domname = 'FINS_LEDGER' memoryid = 'GLN_FLEX' headlen = '02' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'BORECKY' as4date = '20180517' as4time = '092604' dtelmaster = 'E'
                      shlpname = 'FINS_LEDGER' shlpfield = 'RLDNR' datatype = 'CHAR' leng = '000002' outputlen = '000002' convexit = 'ALPHA' entitytab = 'FINSC_LEDGER' refkind = 'D' )
                    ( rollname = 'BF_INKEN' as4local = 'A' domname = 'BF_XFELD' logflag = 'X' headlen = '02' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155500' dtelmaster = 'D'
                      deffdname = 'INVENT_IND' datatype = 'CHAR' leng = '000001' outputlen = '000001' valexi = 'X' refkind = 'D' )
                    ( rollname = 'FB_SEGMENT' as4local = 'A' domname = 'FB_SEGMENT' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'KENNTNER' as4date = '20041220' as4time = '123507' dtelmaster = 'D' deffdname =
                      'SEGMENT' datatype = 'CHAR' leng = '000010' outputlen = '000010' convexit = 'ALPHA' entitytab = 'FAGL_SEGM' refkind = 'D' )
                    ( rollname = 'PRCTR' as4local = 'A' domname = 'PRCTR' memoryid = 'PRC' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'KE' as4user = 'LISSITSYNA' as4date = '20170807' as4time = '232918'
                      dtelmaster = 'D' shlpname = 'PRCTR_EMPTY' shlpfield = 'PRCTR' deffdname = 'PROFIT_CTR' datatype = 'CHAR' leng = '000010' outputlen = '000010' convexit = 'ALPHA' entitytab = 'CEPC' refkind = 'D' )
                    ( rollname = 'XNEU_AM' as4local = 'A' domname = 'XFELD' scrlen3 = '20' as4user = 'SAP' as4date = '20010607' as4time = '160042' dtelmaster = 'D' deffdname = 'PURCH_NEW' datatype = 'CHAR' leng = '000001' outputlen = '000001' valexi
                      = 'X' refkind = 'D' )
                    ( rollname = 'XNACH_ANLA' as4local = 'A' domname = 'XFELD' headlen = '20' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '160041' dtelmaster = 'D' datatype = 'CHAR'
                      leng = '000001' outputlen = '000001' valexi = 'X' refkind = 'D' )
                    ( rollname = 'BF_AKTIVD' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '19980218' as4time = '064341' dtelmaster = 'D'
                      deffdname = 'CAP_DATE' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BF_AFABG' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'ABA' as4user = 'SAP' as4date = '20010607' as4time = '155453' dtelmaster = 'D'
                      deffdname = 'ORD_DEP_DT' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BF_IVDAT_ANLA' as4local = 'A' domname = 'DATUM' logflag = 'X' headlen = '10' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155501' dtelmaster = 'D'
                      deffdname = 'LASTINVDAT' datatype = 'DATS' leng = '000008' outputlen = '000010' refkind = 'D' )
                    ( rollname = 'BF_URWRT' as4local = 'A' domname = 'BAPICURR' logflag = 'X' headlen = '17' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20180219' as4time = '111515' dtelmaster = 'D'
                      deffdname = 'VALUE_OAS' datatype = 'DEC' leng = '000023' decimals = '000004' outputlen = '000030' refkind = 'D' )
                    ( rollname = 'BF_NDJAR' as4local = 'A' domname = 'BF_JARAF' logflag = 'X' headlen = '03' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'ABA' as4user = 'SAP' as4date = '20010607' as4time = '155505' dtelmaster = 'D'
                      deffdname = 'USF_LIF_YR' datatype = 'NUMC' leng = '000003' outputlen = '000003' refkind = 'D' )
                    ( rollname = 'BF_AFASL' as4local = 'A' domname = 'BF_AFASL' logflag = 'X' headlen = '05' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AB' as4user = 'SAP' as4date = '20010607' as4time = '155453' dtelmaster = 'D'
                      deffdname = 'DEPREC_KEY' datatype = 'CHAR' leng = '000004' outputlen = '000004' refkind = 'D' )
                    ( rollname = 'TXJCD' as4local = 'A' domname = 'TXJCD' memoryid = 'TXJ' logflag = 'X' headlen = '15' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FB' as4user = 'LISSITSYNA' as4date = '20170807' as4time = '232918'
                      dtelmaster = 'E' deffdname = 'TAXJURCODE' datatype = 'CHAR' leng = '000015' outputlen = '000015' entitytab = 'TTXJ' refkind = 'D' )
                    ( rollname = 'BF_TYPBZ_ANLA' as4local = 'A' domname = 'BF_TYPBZ_ANLA' logflag = 'X' headlen = '15' scrlen1 = '10' scrlen2 = '16' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155507' dtelmaster =
                      'D' deffdname = 'TYPE_NAME' datatype = 'CHAR' leng = '000015' outputlen = '000015' refkind = 'D' )
                    ( rollname = 'BF_INVZU_ANLA' as4local = 'A' domname = 'BF_INVZU_ANLA' logflag = 'X' headlen = '15' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155501' dtelmaster =
                      'D' deffdname = 'INV_NOTE' datatype = 'CHAR' leng = '000015' outputlen = '000015' refkind = 'D' )
                    ( rollname = 'BF_AFABE_D' as4local = 'A' domname = 'BF_AFABE' memoryid = 'AFB' headlen = '03' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AB' as4user = 'SAP' as4date = '20010607' as4time = '155453' dtelmaster = 'D'
                      deffdname = 'DEPR_AREA' datatype = 'NUMC' leng = '000002' outputlen = '000002' refkind = 'D' )
                    ( rollname = 'BF_AM_LAND1' as4local = 'A' domname = 'LAND1' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' as4user = 'SAP' as4date = '20010607' as4time = '155454' dtelmaster = 'D' deffdname = 'COUNTRY' datatype = 'CHAR' leng =
                      '000003' outputlen = '000003' entitytab = 'T005' refkind = 'D' )
                    ( rollname = 'BF_HERST' as4local = 'A' domname = 'TEXT30' logflag = 'X' headlen = '30' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'AA' as4user = 'SAP' as4date = '20010607' as4time = '155500' dtelmaster = 'D'
                      deffdname = 'MANUFACT' datatype = 'CHAR' leng = '000030' outputlen = '000030' lowercase = 'X' refkind = 'D' )
                    ( rollname = 'BF_RASSC' as4local = 'A' domname = 'RCOMP' logflag = 'X' headlen = '06' scrlen1 = '10' scrlen2 = '15' scrlen3 = '20' applclass = 'FG' as4user = 'SAP' as4date = '19980218' as4time = '064411' dtelmaster = 'D'
                      deffdname = 'TRADE_ID' datatype = 'CHAR' leng = '000006' outputlen = '000006' convexit = 'ALPHA' entitytab = 'T880' refkind = 'D' )
                    ( rollname = 'BAPI1022_POSNR_EXT2' as4local = 'A' domname = 'PS_POSID' headlen = '47' scrlen1 = '10' scrlen2 = '15' scrlen3 = '32' as4user = 'SAP' as4date = '20011120' as4time = '141421' dtelmaster = 'D' datatype = 'CHAR' leng =
                      '000024' outputlen = '000024' convexit = 'ABPSN' refkind = 'D' )
                  ).

    go_osql_environment->insert_test_data( lt_dd04l ).

    DATA: lt_t002 TYPE TABLE OF t002.   "Language Keys (Component BC-I18)

    "Language Keys (Component BC-I18)
    lt_t002 = VALUE #(
                       ( spras = '0' laspez = 'S' lahq = '0' laiso = 'SR' )
                       ( spras = '1' laspez = 'D' lahq = '0' laiso = 'ZH' )
                       ( spras = '2' laspez = 'M' lahq = '0' laiso = 'TH' )
                       ( spras = '3' laspez = 'D' lahq = '0' laiso = 'KO' )
                       ( spras = '4' laspez = 'S' lahq = '0' laiso = 'RO' )
                       ( spras = '5' laspez = 'S' lahq = '0' laiso = 'SL' )
                       ( spras = '6' laspez = 'S' lahq = '0' laiso = 'HR' )
                       ( spras = '7' laspez = 'S' lahq = '4' laiso = 'MS' )
                       ( spras = '8' laspez = 'S' lahq = '0' laiso = 'UK' )
                       ( spras = '9' laspez = 'S' lahq = '0' laiso = 'ET' )
                       ( spras = 'A' laspez = 'L' lahq = '0' laiso = 'AR' )
                       ( spras = 'B' laspez = 'L' lahq = '0' laiso = 'HE' )
                       ( spras = 'C' laspez = 'S' lahq = '4' laiso = 'CS' )
                       ( spras = 'D' laspez = 'S' lahq = '1' laiso = 'DE' )
                       ( spras = 'E' laspez = 'S' lahq = '1' laiso = 'EN' )
                       ( spras = 'F' laspez = 'S' lahq = '2' laiso = 'FR' )
                       ( spras = 'G' laspez = 'S' lahq = '0' laiso = 'EL' )
                       ( spras = 'H' laspez = 'S' lahq = '4' laiso = 'HU' )
                       ( spras = 'I' laspez = 'S' lahq = '2' laiso = 'IT' )
                       ( spras = 'J' laspez = 'D' lahq = '2' laiso = 'JA' )
                       ( spras = 'K' laspez = 'S' lahq = '3' laiso = 'DA' )
                       ( spras = 'L' laspez = 'S' lahq = '0' laiso = 'PL' )
                       ( spras = 'M' laspez = 'D' lahq = '0' laiso = 'ZF' )
                       ( spras = 'N' laspez = 'S' lahq = '2' laiso = 'NL' )
                       ( spras = 'O' laspez = 'S' lahq = '3' laiso = 'NO' )
                     ).

    go_osql_environment->insert_test_data( lt_t002 ).
  ENDMETHOD.

  METHOD user_confign_format_amount.

    DATA: lv_value TYPE string,
          lt_usr01 TYPE STANDARD TABLE OF usr01.
    "When 1
    lt_usr01 = VALUE #( ( bname = sy-uname
                          dcpfm = space ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_usr01 ).
    lv_value = '1.101.123,01'.
    mo_mpa_parse_util->user_confign_decimal_format( CHANGING ch_value = lv_value ).

    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = lv_value
                                        quit = if_aunit_constants=>quit-no
                                        exp  = '1101123.01' ).

    "Then 2
    DELETE usr01 FROM TABLE lt_usr01.
    COMMIT WORK.
    lt_usr01 = VALUE #( ( bname = sy-uname
                          dcpfm = abap_true ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_usr01 ).
    lv_value = '1,101,123.01 '.
    mo_mpa_parse_util->user_confign_decimal_format( CHANGING ch_value = lv_value ).

    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lv_value
                                        quit = if_aunit_constants=>quit-no
                                        exp  = '1101123.01' ).

    "Then 3
    DELETE usr01 FROM TABLE lt_usr01.
    COMMIT WORK.
    lt_usr01 = VALUE #( ( bname = sy-uname
                          dcpfm = 'Y' ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_usr01 ).
    lv_value = ' 1101123,01'.
    mo_mpa_parse_util->user_confign_decimal_format( CHANGING ch_value = lv_value ).

    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = lv_value
                                        quit = if_aunit_constants=>quit-no
                                        exp  = '1101123.01' ).
    "Then 4
    DELETE usr01 FROM TABLE lt_usr01.
    COMMIT WORK.
    lt_usr01 = VALUE #( ( bname = sy-uname
                          dcpfm = 'D' ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_usr01 ).
    lv_value = ' 1101123,01'.
    mo_mpa_parse_util->user_confign_decimal_format( CHANGING ch_value = lv_value ).

    "Then 4
    cl_abap_unit_assert=>assert_equals( act  = lv_value
                                        quit = if_aunit_constants=>quit-no
                                        exp  = ' 1101123,01' ).
  ENDMETHOD.

  METHOD save_file_to_db.

    DATA :
      lx_file_content  TYPE xstring,
      lv_file_name     TYPE string,
      lv_mpa_file_type TYPE mpa_template_type,
      lt_tnro          TYPE STANDARD TABLE OF tnro.

    "When 1
    DATA(ls_message) = mo_mpa_parse_util->save_file_to_db( EXPORTING ix_file_content  = lx_file_content
                                                                     iv_file_name     = lv_file_name
                                                                     iv_mpa_file_type = lv_mpa_file_type ).

    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-num_range
                                        quit = if_aunit_constants=>quit-no ).

    "When 2
    lt_tnro = VALUE #( ( object = mo_mpa_parse_util->gc_nr_object ) ).
    ltc_mpa_xlsx_parse_util=>go_osql_environment->insert_test_data( lt_tnro ).

    ls_message = mo_mpa_parse_util->save_file_to_db( EXPORTING ix_file_content  = lx_file_content
                                                               iv_file_name     = lv_file_name
                                                               iv_mpa_file_type = lv_mpa_file_type ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = ls_message-type
                                        exp  = if_mpa_output=>gc_msg_type-success
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_message-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-file_db
                                        quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD handle_exception.

    DATA : lr_exception     TYPE REF TO cx_mpa_exception_handler,
           lv_je_sequence   TYPE string,
           lt_error_message TYPE bapirettab.
    "When 1
    lr_exception = NEW cx_mpa_exception_handler(  ).

    mo_mpa_parse_util->handle_exception( EXPORTING iref_exception   = lr_exception
                                                   iv_je_sequence   = lv_je_sequence
                                         CHANGING  ct_error_message = lt_error_message ).
    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = lt_error_message[ 1 ]-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_error_message[ 1 ]-id
                                        exp  = 'SY'
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_error_message[ 1 ]-number
                                        exp  = 530
                                        quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD map_excel_data.

    DATA : lt_table          TYPE mpa_t_index_value_pair,
           ls_index          TYPE mpa_s_excel_doc_index,
           lv_mpa_type       TYPE mpa_template_type,
           lt_asset_xls_data TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data,
           lt_message        TYPE bapirettab.

    "When 1
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.

    ls_index-header_techn = '2'.

    fill_table_file_data( EXPORTING iv_scen_type = if_mpa_output=>gc_mpa_scen-transfer
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-transfer
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).


    "When 2
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.

    ls_index-header_techn = '2'.

    fill_table_file_data( EXPORTING iv_scen_type = if_mpa_output=>gc_mpa_scen-transfer
                                    iv_unwanted_field = abap_true
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-transfer
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).

    "When 3
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-change.

    ls_index-header_techn = '2'.

    fill_table_file_data( EXPORTING iv_scen_type = if_mpa_output=>gc_mpa_scen-change
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-change
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).

    "When 4
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-adjustment.

    ls_index-header_techn = '2'.

    fill_table_file_data( EXPORTING iv_scen_type = if_mpa_output=>gc_mpa_scen-adjustment
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 4
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-adjustment
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).

    "When 5
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-retirement.

    ls_index-header_techn = '2'.

    fill_table_file_data( EXPORTING iv_scen_type = if_mpa_output=>gc_mpa_scen-retirement
                          IMPORTING et_table     =  mo_mpa_parse_util->mt_excel_rows ).

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 5
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-retirement
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD check_mpa_status.

    DATA: lv_create_status   TYPE bapi_mtype,
          lv_change_status   TYPE bapi_mtype,
          lv_transfer_status TYPE bapi_mtype.

    CLEAR zcl_xlsx_parse_util=>gv_template_type.
    "When 1
    mo_mpa_parse_util->check_mpa_status( IMPORTING ev_create_status   = lv_create_status
                                                   ev_change_status   = lv_change_status
                                                   ev_transfer_status = lv_transfer_status ).
    "Then 1
    cl_abap_unit_assert=>assert_initial( act  = lv_create_status
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lv_change_status
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lv_transfer_status
                                         quit = if_aunit_constants=>quit-no ).

    "When 2
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.

    mo_mpa_parse_util->check_mpa_status( IMPORTING ev_create_status   = lv_create_status
                                                   ev_change_status   = lv_change_status
                                                   ev_transfer_status = lv_transfer_status ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lv_transfer_status
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lv_create_status
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lv_change_status
                                         quit = if_aunit_constants=>quit-no ).

    "When 3
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-create.

    mo_mpa_parse_util->check_mpa_status( IMPORTING ev_create_status   = lv_create_status
                                                   ev_change_status   = lv_change_status
                                                   ev_transfer_status = lv_transfer_status ).
    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = lv_transfer_status
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lv_create_status
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lv_change_status
                                         quit = if_aunit_constants=>quit-no ).
    "When 4
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-change.

    mo_mpa_parse_util->check_mpa_status( IMPORTING ev_create_status   = lv_create_status
                                                   ev_change_status   = lv_change_status
                                                   ev_transfer_status = lv_transfer_status ).
    "Then 4
    cl_abap_unit_assert=>assert_equals( act  = lv_transfer_status
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lv_create_status
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lv_change_status
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD map_excel_data_exp.

    DATA : lt_table          TYPE mpa_t_index_value_pair,
           ls_index          TYPE mpa_s_excel_doc_index,
           lv_mpa_type       TYPE mpa_template_type,
           lt_asset_xls_data TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data,
           lt_message        TYPE bapirettab.

    "When 1
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.

    ls_index-header_techn = '2'.
    ls_index = VALUE #( header_techn = '2'
                        data = VALUE #( ( '4' ) ( '5' ) ) ).

    fill_table_dile_with_slno(
      EXPORTING
        iv_scen_type = if_mpa_output=>gc_mpa_scen-transfer
      IMPORTING
        et_table     = mo_mpa_parse_util->mt_excel_rows
    ).

*ASSIGN mo_mpa_parse_util->mt_excel_rows[ index = 2 ] to FIELD-SYMBOL(<lt_row_2>).

*<lt_row_2>[ index = 1 ] =

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-transfer
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_not_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD map_excel_data_exp_long_date.

    DATA : lt_table          TYPE mpa_t_index_value_pair,
           ls_index          TYPE mpa_s_excel_doc_index,
           lv_mpa_type       TYPE mpa_template_type,
           lt_asset_xls_data TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data,
           lt_message        TYPE bapirettab.

    "When 1
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.

    ls_index-header_techn = '2'.
    ls_index = VALUE #( header_techn = '2'
                        data = VALUE #( ( '4' ) ( '5' ) ) ).

    fill_table_dile_with_long_date(
      EXPORTING
        iv_scen_type = if_mpa_output=>gc_mpa_scen-transfer
      IMPORTING
        et_table     = mo_mpa_parse_util->mt_excel_rows
    ).

*ASSIGN mo_mpa_parse_util->mt_excel_rows[ index = 2 ] to FIELD-SYMBOL(<lt_row_2>).

*<lt_row_2>[ index = 1 ] =

    mo_mpa_parse_util->map_excel_data( EXPORTING it_table          = mo_mpa_parse_util->mt_excel_rows
                                                 is_index          = ls_index
                                       IMPORTING ev_mpa_type       = lv_mpa_type
                                                 et_asset_xls_data = lt_asset_xls_data
                                                 et_message        = lt_message ).
    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp  = if_mpa_output=>gc_mpa_scen-transfer
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_not_initial( act  = lt_message
                                         quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD build_mpa_xls_itab.

    CONSTANTS lc_bukrs TYPE bukrs VALUE '0001'.

    DATA: ls_mass_create     TYPE mpa_s_asset_create,
          ls_mass_change     TYPE mpa_s_asset_change,
          ls_mass_transfer   TYPE mpa_s_asset_transfer,
          ls_mass_adjustment TYPE mpa_s_asset_adjustment,
          lt_mass_adjustment TYPE mpa_t_asset_adjustment,
          ls_mass_retirement TYPE mpa_s_asset_retirement,
          lt_mass_retirement TYPE mpa_t_asset_retirement,
          lt_mass_create     TYPE mpa_t_asset_create,
          lt_mass_change     TYPE mpa_t_asset_change,
          lt_mass_transfer   TYPE mpa_t_asset_transfer.
    "When 1
    mo_mpa_parse_util->build_mpa_xls_itab( EXPORTING is_mass_create     = ls_mass_create
                                                     is_mass_change     = ls_mass_change
                                                     is_mass_transfer   = ls_mass_transfer
                                                     is_mass_adjustment = ls_mass_adjustment
                                                     is_mass_retirement = ls_mass_retirement
                                            CHANGING ct_mass_create     = lt_mass_create
                                                     ct_mass_change     = lt_mass_change
                                                     ct_mass_transfer   = lt_mass_transfer
                                                     ct_mass_adjustment = lt_mass_adjustment
                                                     ct_mass_retirement = lt_mass_retirement ).
    "Then 1
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_create
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_change
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_transfer
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_adjustment
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_retirement
                                         quit = if_aunit_constants=>quit-no ).
    "When 2
    ls_mass_create = VALUE #( bukrs = lc_bukrs ).
    mo_mpa_parse_util->build_mpa_xls_itab( EXPORTING is_mass_create     = ls_mass_create
                                                     is_mass_change     = ls_mass_change
                                                     is_mass_transfer   = ls_mass_transfer
                                                     is_mass_adjustment = ls_mass_adjustment
                                                     is_mass_retirement = ls_mass_retirement
                                            CHANGING ct_mass_create     = lt_mass_create
                                                     ct_mass_change     = lt_mass_change
                                                     ct_mass_transfer   = lt_mass_transfer
                                                     ct_mass_adjustment = lt_mass_adjustment
                                                     ct_mass_retirement = lt_mass_retirement ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lt_mass_create[ 1 ]-bukrs
                                        exp  = lc_bukrs
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_change
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_transfer
                                         quit = if_aunit_constants=>quit-no ).
    "When 3
    ls_mass_transfer = VALUE #( bukrs = lc_bukrs ).
    mo_mpa_parse_util->build_mpa_xls_itab( EXPORTING is_mass_create     = ls_mass_create
                                                     is_mass_change     = ls_mass_change
                                                     is_mass_transfer   = ls_mass_transfer
                                                     is_mass_adjustment = ls_mass_adjustment
                                                     is_mass_retirement = ls_mass_retirement
                                            CHANGING ct_mass_create     = lt_mass_create
                                                     ct_mass_change     = lt_mass_change
                                                     ct_mass_transfer   = lt_mass_transfer
                                                     ct_mass_adjustment = lt_mass_adjustment
                                                     ct_mass_retirement = lt_mass_retirement ).
    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = lt_mass_create[ 1 ]-bukrs
                                        exp  = lc_bukrs
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_initial( act  = lt_mass_change
                                         quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_mass_transfer[ 1 ]-bukrs
                                        exp  = lc_bukrs
                                        quit = if_aunit_constants=>quit-no ).
    "When 4
    CLEAR : ls_mass_create, ls_mass_transfer.

    ls_mass_change = VALUE #( bukrs = lc_bukrs ).
    mo_mpa_parse_util->build_mpa_xls_itab( EXPORTING is_mass_create     = ls_mass_create
                                                     is_mass_change     = ls_mass_change
                                                     is_mass_transfer   = ls_mass_transfer
                                                     is_mass_adjustment = ls_mass_adjustment
                                                     is_mass_retirement = ls_mass_retirement
                                            CHANGING ct_mass_create     = lt_mass_create
                                                     ct_mass_change     = lt_mass_change
                                                     ct_mass_transfer   = lt_mass_transfer
                                                     ct_mass_adjustment = lt_mass_adjustment
                                                     ct_mass_retirement = lt_mass_retirement ).
    "Then 4
    cl_abap_unit_assert=>assert_equals( act  = lt_mass_create[ 1 ]-bukrs
                                        exp  = lc_bukrs
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_mass_change[ 1 ]-bukrs
                                        exp  = lc_bukrs
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_mass_transfer[ 1 ]-bukrs
                                        exp  = lc_bukrs
                                        quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD get_excel_length_error_msg.


    DATA : lv_cell_name TYPE string VALUE 'A',
           lv_index     TYPE i VALUE '1',
           lv_length    TYPE ddleng VALUE '1'.
    "When
    DATA(lt_message) = mo_mpa_parse_util->get_excel_length_error_msg( EXPORTING iv_cell_name = lv_cell_name
                                                                                iv_index     = lv_index
                                                                                iv_length    = lv_length ).
    "Then
    cl_abap_unit_assert=>assert_equals( act  = lt_message[ 1 ]-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_message[ 1 ]-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_message[ 1 ]-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-field_chars
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_message[ 1 ]-message_v1
                                        exp  = lv_cell_name
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_message[ 1 ]-message_v2
                                        exp  = lv_index
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = lt_message[ 1 ]-system
                                        exp  = sy-sysid
                                        quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD insert_file_to_db.

    "When 1
    DATA(ls_file_data) = VALUE mpa_asset_data( file_id  = '123456'
                                               ernam    = sy-uname
                                               erdat    = sy-datum ).

    DATA(ls_messages) = mo_mpa_parse_util->insert_file_to_db( is_file_data = ls_file_data ).
    "Then 1
    SELECT SINGLE * FROM mpa_asset_data INTO @DATA(ls_data) WHERE file_id  = '123456'.

    cl_abap_unit_assert=>assert_equals( act  = ls_data-file_id
                                        exp  = ls_file_data-file_id
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_data-ernam
                                        exp  = sy-uname
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_file_data-erdat
                                        exp  = sy-datum
                                        quit = if_aunit_constants=>quit-no ).

    cl_abap_unit_assert=>assert_equals( act  = ls_messages-type
                                        exp  = if_mpa_output=>gc_msg_type-success
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_messages-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_messages-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-file_db
                                        quit = if_aunit_constants=>quit-no ).

    "When 2
    ls_file_data = VALUE mpa_asset_data( file_id  = '123456'
                                         ernam    = sy-uname
                                         erdat    = sy-datum ).

    ls_messages = mo_mpa_parse_util->insert_file_to_db( is_file_data = ls_file_data ).
    "Then 2
    SELECT SINGLE * FROM mpa_asset_data INTO ls_data WHERE file_id  = '123456'.

    cl_abap_unit_assert=>assert_equals( act  = ls_data-file_id
                                        exp  = ls_file_data-file_id
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_data-ernam
                                        exp  = sy-uname
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_file_data-erdat
                                        exp  = sy-datum
                                        quit = if_aunit_constants=>quit-no ).


    cl_abap_unit_assert=>assert_equals( act  = ls_messages-type
                                        exp  = if_mpa_output=>gc_msg_type-error
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_messages-id
                                        exp  = if_mpa_output=>gc_msgid-mpa
                                        quit = if_aunit_constants=>quit-no ).
    cl_abap_unit_assert=>assert_equals( act  = ls_messages-number
                                        exp  = if_mpa_output=>gc_msgno_mpa-error_file
                                        quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD convert_dec_time_to_hhmmss.

    DATA lv_dec_time_string  TYPE string.
    "When 1
    lv_dec_time_string = '-99.59'.

    DATA(lv_time) = mo_lcl_parse_util->convert_dec_time_to_hhmmss( iv_dec_time_string = lv_dec_time_string ).

    "Then 1
    cl_abap_unit_assert=>assert_equals( act  = lv_time
                                        exp  = '0-1550'
                                        quit = if_aunit_constants=>quit-no ).
    "When 2
    lv_dec_time_string = 'ABC'.

    lv_time = mo_lcl_parse_util->convert_dec_time_to_hhmmss( iv_dec_time_string = lv_dec_time_string ).

    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lv_time
                                        exp  = 'ABC000'
                                        quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD convert_long_to_date.

    "When 1
    mo_lcl_parse_util->dateformat1904 = abap_true.
    DATA(lv_date) = mo_lcl_parse_util->convert_long_to_date( iv_date_string = '9999999' ).
    "Then 1
    cl_abap_unit_assert=>assert_initial( act  = lv_date
                                         quit = if_aunit_constants=>quit-no ).

    "When 2
    mo_lcl_parse_util->dateformat1904 = abap_false.
    lv_date = mo_lcl_parse_util->convert_long_to_date( iv_date_string = '60' ).
    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lv_date
                                        exp  = '19000228'
                                        quit = if_aunit_constants=>quit-no ).
    "When 3
    mo_lcl_parse_util->dateformat1904 = abap_false.
    lv_date = mo_lcl_parse_util->convert_long_to_date( iv_date_string = '6' ).
    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = lv_date
                                        exp  = '19000106'
                                        quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD check_excel_layout.

    DATA: lt_check_table TYPE mpa_t_index_value_pair,
          lt_table       TYPE mpa_t_index_value_pair.
    "When 1
    TRY .
        mo_mpa_parse_util->if_mpa_xlsx_parse_util~check_excel_layout( IMPORTING et_line_index = DATA(lt_line_index)
                                                                      CHANGING  ct_table      = lt_table ).
      CATCH  cx_mpa_exception_handler .

    ENDTRY.

    "Then 1
    cl_abap_unit_assert=>assert_initial( act  = lt_table
                                         quit = if_aunit_constants=>quit-no ).

    "When 2
    fill_table_file_data( EXPORTING iv_scen_type   = if_mpa_output=>gc_mpa_scen-transfer
                                    iv_clear_title = abap_true
                          IMPORTING et_table       = DATA(t_index_value_pair) ).

    lt_check_table = t_index_value_pair.
    TRY .
        mo_mpa_parse_util->if_mpa_xlsx_parse_util~check_excel_layout( IMPORTING et_line_index = lt_line_index
                                                                      CHANGING  ct_table      = t_index_value_pair ).
      CATCH  cx_mpa_exception_handler .
    ENDTRY.

    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = t_index_value_pair
                                        exp  = lt_check_table
                                        quit = if_aunit_constants=>quit-no ).

    "When 3
    fill_table_file_data( EXPORTING iv_scen_type      = if_mpa_output=>gc_mpa_scen-transfer
                                    iv_clear_techdesc = abap_true
                          IMPORTING et_table          = t_index_value_pair ).

    lt_check_table = t_index_value_pair.
    TRY .
        mo_mpa_parse_util->if_mpa_xlsx_parse_util~check_excel_layout( IMPORTING et_line_index = lt_line_index
                                                                      CHANGING  ct_table      = t_index_value_pair ).
      CATCH  cx_mpa_exception_handler .
    ENDTRY.

    "Then 3
    cl_abap_unit_assert=>assert_equals( act  = t_index_value_pair
                                        exp  = lt_check_table
                                        quit = if_aunit_constants=>quit-no ).

    "When 4
    fill_table_file_data( EXPORTING iv_scen_type   = if_mpa_output=>gc_mpa_scen-transfer
                                    iv_no_data     = abap_true
                          IMPORTING et_table       = t_index_value_pair ).

    lt_check_table = t_index_value_pair.
    TRY .
        mo_mpa_parse_util->if_mpa_xlsx_parse_util~check_excel_layout( IMPORTING et_line_index = lt_line_index
                                                                       CHANGING ct_table      = t_index_value_pair ).
      CATCH  cx_mpa_exception_handler .
    ENDTRY.

    "Then 4
    cl_abap_unit_assert=>assert_equals( act  = t_index_value_pair
                                        exp  = lt_check_table
                                        quit = if_aunit_constants=>quit-no ).


    "When 5
    fill_table_file_data( EXPORTING iv_scen_type   = if_mpa_output=>gc_mpa_scen-transfer
                                    iv_change_title_name  = abap_true
                          IMPORTING et_table              = t_index_value_pair ).

    lt_check_table = t_index_value_pair.
    TRY .
        mo_mpa_parse_util->if_mpa_xlsx_parse_util~check_excel_layout( IMPORTING et_line_index = lt_line_index
                                                                      CHANGING  ct_table      = t_index_value_pair ).
      CATCH  cx_mpa_exception_handler .
    ENDTRY.

    "Then 5
    cl_abap_unit_assert=>assert_equals( act  = t_index_value_pair
                                        exp  = lt_check_table
                                        quit = if_aunit_constants=>quit-no ).
    "When 6
    fill_table_file_data( EXPORTING iv_scen_type           = if_mpa_output=>gc_mpa_scen-change
                                    iv_comment_cell_change = abap_true
                          IMPORTING et_table               = t_index_value_pair ).

    lt_check_table = t_index_value_pair.
    TRY .
        mo_mpa_parse_util->if_mpa_xlsx_parse_util~check_excel_layout( IMPORTING et_line_index = lt_line_index
                                                                      CHANGING  ct_table      = t_index_value_pair ).
      CATCH  cx_mpa_exception_handler .
    ENDTRY.

    "Then 6
    cl_abap_unit_assert=>assert_equals( act  = t_index_value_pair
                                        exp  = lt_check_table
                                        quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD parse_xlsx.

    DATA lv_file  TYPE xstring.
    "When 1
    TRY.
        mo_mpa_parse_util->parse_xlsx(
          EXPORTING
            ix_file     = lv_file
          IMPORTING
            ev_mpa_type = DATA(lv_mpa_type) ).
      CATCH cx_openxml_not_found.
      CATCH cx_openxml_format.
    ENDTRY.

    "Then 1
    cl_abap_unit_assert=>assert_initial( act  = lv_mpa_type
                                         quit = if_aunit_constants=>quit-no ).
    "When 2
    TRY.
        lv_file = '123'.
        mo_mpa_parse_util->parse_xlsx(
          EXPORTING
            ix_file     = lv_file
          IMPORTING
            ev_mpa_type = lv_mpa_type ).
      CATCH cx_openxml_not_found.
      CATCH cx_openxml_format.
    ENDTRY.

    "Then 2
    cl_abap_unit_assert=>assert_equals( act  = lv_mpa_type
                                        exp = ''
                                        quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD remove_comment_line.

    DATA : lt_comment_lines TYPE zcl_xlsx_parse_util=>gty_t_comment_line,
           lt_excel_lines   TYPE mpa_t_index_value_pair,
           lt_csv_lines     TYPE string_table.

    mo_mpa_parse_util->remove_comment_line( IMPORTING  et_comment_lines = lt_comment_lines
                                            CHANGING   ct_excel_lines   = lt_excel_lines
                                                       ct_csv_lines     = lt_csv_lines  ).

  ENDMETHOD.

  METHOD get_struct_properties.

    DATA lv_struct_name  TYPE mpa_s_asset_create.
    "When 1
    mo_mpa_parse_util->if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = lv_struct_name
                                                                     IMPORTING et_struct_properties = DATA(lt_struct_properties) ).

    "Then 1
    cl_abap_unit_assert=>assert_not_initial( act  = lt_struct_properties
                                         quit = if_aunit_constants=>quit-no ).

    DATA lv_mpa_struct_name  TYPE mpa_asset_data.
    CLEAR lt_struct_properties.
    "When 2
    mo_mpa_parse_util->if_mpa_xlsx_parse_util~get_struct_properties( EXPORTING iv_struct_name       = lv_mpa_struct_name
                                                                     IMPORTING et_struct_properties = lt_struct_properties ).

    "Then 2
    cl_abap_unit_assert=>assert_initial( act  = lt_struct_properties
                                         quit = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD convert_dec_Time.
    "given
    DATA(lo_cut) = NEW lcl_mpa_xlsx_parse_util( ).


    "when then
    cl_abap_unit_assert=>assert_equals( act  = lo_cut->convert_dec_time_to_hhmmss( iv_dec_time_string = '50.99' )
                                        exp =    '234536'
                                         quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD set_supported_separators.
    "given
    CLEAR mo_mpa_parse_util->mt_supported_separators .
    DATA lt_support_seperator TYPE TABLE OF mo_mpa_parse_util->ty_supported_sep WITH KEY separator .
    lt_support_seperator = VALUE #( ( separator = mo_mpa_parse_util->gc_sep_semicolon )
                                    ( separator = mo_mpa_parse_util->gc_sep_comma ) ).

    "when
    mo_mpa_parse_util->set_supported_separators( ).

    "then
    cl_abap_unit_assert=>assert_equals( act  = mo_mpa_parse_util->mt_supported_separators
                                    exp =   lt_support_seperator
                                     quit = if_aunit_constants=>quit-no
                                     msg = 'The sepeartors are not matching , and ;').


  ENDMETHOD.

  METHOD remove_comment_line_csv.

    DATA : lt_comment_lines       TYPE zcl_xlsx_parse_util=>gty_t_comment_line,
           lt_csv_lines_wo_coment TYPE string_table.

    lt_csv_lines_wo_coment = mt_csv_line_data.
    DELETE lt_csv_lines_wo_coment WHERE table_line CS '//'.

    "when
    mo_mpa_parse_util->remove_comment_line( IMPORTING  et_comment_lines = lt_comment_lines
                                            CHANGING   ct_csv_lines     = mt_csv_line_data  ).


    "then
    cl_abap_unit_assert=>assert_equals( act  = mt_csv_line_data
                                        exp =   lt_csv_lines_wo_coment
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'The comented lines where not deleted as expected').


  ENDMETHOD.

  METHOD find_separator.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mt_supported_separators = VALUE #( ( separator = mo_mpa_parse_util->gc_sep_semicolon )
                                                          ( separator = mo_mpa_parse_util->gc_sep_comma ) ).

    "when
    mo_mpa_parse_util->find_separator( EXPORTING it_linetab   = mt_csv_line_data
                                       IMPORTING ev_separator = DATA(lv_seperator) ).

    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_seperator
                                        exp =   mo_mpa_parse_util->gc_sep_semicolon
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'Issue in finding separator ! expected seperator is ;').

  ENDMETHOD.

  METHOD find_separator_invalid_file.

    "given
    DATA lt_linetab TYPE string_table.
    mo_mpa_parse_util->mt_supported_separators = VALUE #( ( separator = mo_mpa_parse_util->gc_sep_semicolon )
                                                          ( separator = mo_mpa_parse_util->gc_sep_comma ) ).

    "when
    TRY.
        mo_mpa_parse_util->find_separator( EXPORTING it_linetab   = lt_linetab
                                           IMPORTING ev_separator = DATA(lv_seperator) ).
      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no

    ).

  ENDMETHOD.

  METHOD find_separator_invalid_sep.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mt_supported_separators = VALUE #( ( separator = '$' ) ).

    "when
    TRY.
        mo_mpa_parse_util->find_separator( EXPORTING it_linetab   = mt_csv_line_data
                                           IMPORTING ev_separator = DATA(lv_seperator) ).
      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no

    ).

  ENDMETHOD.

  METHOD check_csv_layout.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.

    "when
    mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                   iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                         IMPORTING ev_status_flag = DATA(lv_status_flag) ).

    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_status_flag
                                        exp =   abap_false
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'There is status field in the asset data').



  ENDMETHOD.

  METHOD check_csv_layout_create.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mt_csv_line_data[ 1 ] =  `Asset Mass Create;;;;;;;;;;;;;;;;;;;;;;;;;;;;`.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).
      CATCH  cx_mpa_exception_handler  INTO DATA(lx_exception).
    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_status_flag
                                        exp =   abap_false
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'There is status field in the asset data').

  ENDMETHOD.

  METHOD check_csv_layout_retirement.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mt_csv_line_data[ 1 ] =  `Asset Mass Retirement;;;;;;;;;;;;;;;;;;;;;;;;;;;;`.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).
      CATCH  cx_mpa_exception_handler  INTO DATA(lx_exception).
    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_status_flag
                                        exp =   abap_false
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'There is status field in the asset data').

  ENDMETHOD.

  METHOD check_csv_layout_change.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mt_csv_line_data[ 1 ] =  `Asset Mass Change;;;;;;;;;;;;;;;;;;;;;;;;;;;;`.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).
      CATCH  cx_mpa_exception_handler  INTO DATA(lx_exception).
    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_status_flag
                                        exp =   abap_false
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'There is status field in the asset data').

  ENDMETHOD.

  METHOD check_csv_layout_adjustment.

    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mt_csv_line_data[ 1 ] =  `Asset Mass Adjustment;;;;;;;;;;;;;;;;;;;;;;;;;;;;`.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).
      CATCH  cx_mpa_exception_handler  INTO DATA(lx_exception).
    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_status_flag
                                        exp =   abap_false
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'There is status field in the asset data').

  ENDMETHOD.

  METHOD check_csv_layout_line_lt_3.
    "given
    mt_csv_line_data = VALUE #( ( `Asset Mass Create;;;;;;;;;;;;;;;;;;;;;;;;;;;;` ) ).

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).
      CATCH  cx_mpa_exception_handler  INTO DATA(lx_exception).
    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound( act              = lx_exception
                                       msg              = 'No exception thrown'
                                       quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD check_csv_layout_with_status.
    "given
    DATA(lt_csv_line)  = VALUE string_table( ( `Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                           ( `SLNO;STATUS;BLART;BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR` )
                                            ( `"*Row";"Status";"Document";"*Document";"*Posting Date";"*Asset Value";"Fiscal";"*Translation";"*Company";"KJ";"hui";"as";"aa";"qq";"qw";"yxff";"tt";"gh";"qw";"hh";"w";"q";"a";"y";"x";"c";"e";"f";"s";"g"` )
                                           ( `4;;;2021-01-01;2021-01-01;2021-01-01;;12;2020-10-25;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;` )
                                           ( `5;;AA;2021-01-01;2021-01-01;2021-01-01;;9;2020-06-15;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;` )  ).

    "when
    mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = lt_csv_line
                                                   iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                         IMPORTING ev_status_flag = DATA(lv_status_flag) ).

    "then
    cl_abap_unit_assert=>assert_equals( act  = lv_status_flag
                                        exp =   abap_true
                                       quit = if_aunit_constants=>quit-no
                                        msg = 'There is no status field in the asset data').



  ENDMETHOD.

  METHOD check_csv_exp_layout_not_ok.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lx_excp->if_t100_message~t100key
        exp                  = cx_mpa_exception_handler=>file_layout_not_ok
        msg                  = 'Exception type is not "file layout not ok"'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.


  METHOD check_csv_layout_hdr_not_ok.

    mt_csv_line_data[ 1 ] = `Asset Mass Transfer;;;;;;;asdasd;;;;;;;;;;;;;;;;;;;;;`.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD check_csv_layout_hdr_not_ok_2.

    mt_csv_line_data[ 1 ] = `;;Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;`.

    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = mt_csv_line_data
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD check_csv_layout_hdr_not_ok_3.

    "given
    DATA(lt_csv_line)  = VALUE string_table( ( `Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                           ( `SLNO;STATUS;BLART;BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR` )
                                            ( `"*Row";"Status";"Document";"*Document";"*Posting Date";"*Asset Value";"Fiscal";"*Translation";"*Company";"KJ";"hui";"as";"aa";"qq";"qw";"yxff";"tt";"gh";"qw";"hh";"w";"q";"a";"y";"x";"c";"e";"f";"s";"g"` )
                                           ( `;;;;;2021-01-01;;;2020-10-25;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;` )
                                           ( `;;AA;2021-01-01;2021-01-01;2021-01-01;;9;2020-06-15;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;` )  ).
    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = lt_csv_line
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD check_csv_layout_hdr_not_ok_4.

    "given
    DATA(lt_csv_line)  = VALUE string_table( ( `Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                           ( `BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR` )
                                            ( `"*Row";"Status";"Document";"*Document";"*Posting Date";"*Asset Value";"Fiscal";"*Translation";"*Company";"KJ";"hui";"as";"aa";"qq";"qw";"yxff";"tt";"gh";"qw";"hh";"w";"q";"a";"y";"x";"c";"e";"f";"s";"g"` )
                                           ( `4;;;2021-01-01;2021-01-01;2021-01-01;;;2020-10-25;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;` )
                                           ( `5;;AA;2021-01-01;2021-01-01;2021-01-01;;9;2020-06-15;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;` )  ).
    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = lt_csv_line
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD check_csv_layout_hdr_not_ok_5.

    "given
    DATA(lt_csv_line)  = VALUE string_table( ( `Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                           ( `SLNO;STATUS;BLART;BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR` )
                                            ( `"*Row";"Status";"Document";"*Document";"*Posting Date";"*Asset Value";"Fiscal";"*Translation";"*Company";"KJ";"hui";"as";"aa";"qq";"qw";"yxff";"tt";"gh";"qw";"hh";"w";"q";"a";"y";"x";"c";"e";"f";"s";"g"` )
                                           ( `4;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;` )
                                           ( `5;9;2020-06-15;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;` )  ).
    "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = lt_csv_line
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD check_csv_layout_more_cells.

    "given
    DATA(lt_csv_line)  = VALUE string_table( ( `Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                           ( `SLNO;STATUS;BLART;BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR` )
                                            ( `"*Row";"Status";"Document";"*Document";"*Posting Date";"*Asset Value";"Fiscal";"*Translation";"*Company";"KJ";"hui";"as";"aa";"qq";"qw";"yxff";"tt";"gh";"qw";"hh";"w";"q";"a";"y";"x";"c";"e";"f";"s";"g"` )
                                           ( `4;;;2021-01-01;2021-01-01;2021-01-01;;;2020-10-25;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;` )
                                           ( `5;;AA;2021-01-01;2021-01-01;2021-01-01;;9;2020-06-15;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;` )  ).
  "when
    TRY.
        mo_mpa_parse_util->check_csv_layout( EXPORTING it_line        = lt_csv_line
                                                       iv_separator   = CONV #( mo_mpa_parse_util->gc_sep_semicolon )
                                             IMPORTING ev_status_flag = DATA(lv_status_flag) ).

      CATCH cx_mpa_exception_handler INTO DATA(lx_excp).

    ENDTRY.
    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_excp
        msg              = 'Exception is not raised !'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD process_lines.

    lcl_function_module=>go_instance = NEW lcl_function_module_mock( ).
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-transfer.
    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mv_separator = mo_mpa_parse_util->gc_sep_semicolon .

    "when
    mo_mpa_parse_util->process_lines(  EXPORTING it_line        = mt_csv_line_data
                                                 iv_status_flag = abap_false
                                       IMPORTING ev_mpa_type    = DATA(lv_mpa_type)
                                                 et_asset_data  = DATA(lt_asset_data)
                                                 et_message     = DATA(lt_message) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lv_mpa_type
        exp                  = 'MT'
        msg                  = 'Asset data is not of Mass Transfer type !'
        quit                 = if_abap_unit_constant=>quit-no ).


    cl_abap_unit_assert=>assert_not_initial(
      EXPORTING
        act              = lt_asset_data
        msg              = 'Asset data is initial !'
        quit             = if_abap_unit_constant=>quit-no  ).

    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lines( lt_asset_data[ 1 ]-mass_transfer_data )
        exp                  = 2
        msg                  = 'Mass transfer data does not contain 2 lines as expected'
        quit                 = if_abap_unit_constant=>quit-no ).

    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lines( lt_asset_data[ 1 ]-mass_change_data )
        exp                  = 0
        msg                  = 'All mass asset data exept for transfer should be empty !'
        quit                 = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD process_lines_create.

    lcl_function_module=>go_instance = NEW lcl_function_module_mock( ).
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-create.
    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mv_separator = mo_mpa_parse_util->gc_sep_semicolon .

    "when
    mo_mpa_parse_util->process_lines(  EXPORTING it_line        = mt_csv_line_data
                                                 iv_status_flag = abap_false
                                       IMPORTING ev_mpa_type    = DATA(lv_mpa_type)
                                                 et_asset_data  = DATA(lt_asset_data)
                                                 et_message     = DATA(lt_message) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lv_mpa_type
        exp                  = 'CR'
        msg                  = 'Asset data is not of Mass Create type !'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD process_lines_change.

    lcl_function_module=>go_instance = NEW lcl_function_module_mock( ).
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-change.
    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mv_separator = mo_mpa_parse_util->gc_sep_semicolon .

    "when
    mo_mpa_parse_util->process_lines(  EXPORTING it_line        = mt_csv_line_data
                                                 iv_status_flag = abap_false
                                       IMPORTING ev_mpa_type    = DATA(lv_mpa_type)
                                                 et_asset_data  = DATA(lt_asset_data)
                                                 et_message     = DATA(lt_message) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lv_mpa_type
        exp                  = 'CH'
        msg                  = 'Asset data is not of Mass Create type !'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD process_lines_adjustment.

    lcl_function_module=>go_instance = NEW lcl_function_module_mock( ).
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-adjustment.
    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mv_separator = mo_mpa_parse_util->gc_sep_semicolon .

    "when
    mo_mpa_parse_util->process_lines(  EXPORTING it_line        = mt_csv_line_data
                                                 iv_status_flag = abap_false
                                       IMPORTING ev_mpa_type    = DATA(lv_mpa_type)
                                                 et_asset_data  = DATA(lt_asset_data)
                                                 et_message     = DATA(lt_message) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lv_mpa_type
        exp                  = 'MA'
        msg                  = 'Asset data is not of Mass Create type !'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD process_lines_retirement.

    lcl_function_module=>go_instance = NEW lcl_function_module_mock( ).
    zcl_xlsx_parse_util=>gv_template_type = if_mpa_output=>gc_mpa_temp-retirement.
    "given
    DELETE mt_csv_line_data WHERE table_line CS '//'.
    mo_mpa_parse_util->mv_separator = mo_mpa_parse_util->gc_sep_semicolon .

    "when
    mo_mpa_parse_util->process_lines(  EXPORTING it_line        = mt_csv_line_data
                                                 iv_status_flag = abap_false
                                       IMPORTING ev_mpa_type    = DATA(lv_mpa_type)
                                                 et_asset_data  = DATA(lt_asset_data)
                                                 et_message     = DATA(lt_message) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lv_mpa_type
        exp                  = 'RT'
        msg                  = 'Asset data is not of Mass Create type !'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD parse_csv_macro.

    "given
    lcl_function_module=>go_instance = NEW lcl_function_module_mock( ).

    "when
    mo_mpa_parse_util->if_mpa_xlsx_parse_util~parse_csv(
      EXPORTING
        ix_file     = '4173736574204D617373205472616E736665723B3B3B3B3B3B3B3B3B3B3B3B3B3'
      IMPORTING
*    ev_mpa_type =
*    et_asset    =
        et_message  = DATA(lt_message) ).

    cl_abap_unit_assert=>assert_not_initial(
      EXPORTING
        act              = lt_message
        msg              = 'Exception not thrown !'
        quit             = if_abap_unit_constant=>quit-no  ).

  ENDMETHOD.




  METHOD fill_table_dile_with_slno.

    DATA : ls_line           TYPE mpa_s_index_value_pair,
           ls_cell           TYPE mpa_s_index_value_pair,
           lt_mpa_asset_data TYPE mpa_t_index_value_pair.

    FIELD-SYMBOLS: <ft_linedata> TYPE mpa_t_index_value_pair,
                   <fs_celldata> TYPE any.


    "First record
    ls_line-index = gc_one.

    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    ls_cell-index = 'A'.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    <fs_celldata> = 'Asset Mass Transfer'.
*        IF iv_change_title_name IS NOT INITIAL.
    <fs_celldata> = 'Mass Transfer'.
*        ENDIF.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    "Second record
    ls_line-index = '2'.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = 'SLNO'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'B'.
    <fs_celldata> = 'BLDAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'C'.
    <fs_celldata> = 'BUDAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'D'.
    <fs_celldata> = 'BZDAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'E'.
    <fs_celldata> = 'SGTXT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'F'.
    <fs_celldata> = 'MONAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'G'.
    <fs_celldata> = 'WWERT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'H'.
    <fs_celldata> = 'BUKRS'.
*        IF iv_unwanted_field IS NOT INITIAL.
    <fs_celldata> = 'BUKRS1'.
*        ENDIF.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'I'.
    <fs_celldata> = 'ANLN1'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'J'.
    <fs_celldata> = 'ANLN2'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'K'.
    <fs_celldata> = 'ACC_PRINCIPLE'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'L'.
    <fs_celldata> = 'AFABER'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'M'.
    <fs_celldata> = 'PBUKRS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'N'.
    <fs_celldata> = 'PANL1'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'O'.
    <fs_celldata> = 'PANL2'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'P'.
    <fs_celldata> = 'ANLKL'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Q'.
    <fs_celldata> = 'KOSTL'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'R'.
    <fs_celldata> = 'TEXT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'S'.
    <fs_celldata> = 'TRAVA'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'T'.
    <fs_celldata> = 'ANBTR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'U'.
    <fs_celldata> = 'WAERS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'V'.
    <fs_celldata> = 'MENGE'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'W'.
    <fs_celldata> = 'MEINS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'X'.
    <fs_celldata> = 'PROZS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Y'.
    <fs_celldata> = 'XANEU'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Z'.
    <fs_celldata> = 'RECID'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'AA'.
    <fs_celldata> = 'XBLNR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'AB'.
    <fs_celldata> = 'DZUONR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    "Third record
*        IF iv_no_data IS INITIAL.
*          IF iv_clear_techdesc IS INITIAL.
    ls_line-index = '3'.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.
*          ENDIF.
    "Fourth record
    ls_line-index = '4'.
*          IF  iv_increase_row IS NOT INITIAL.
*            ls_line-index = '5'.
*          ENDIF.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = '4'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'B'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'C'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'D'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'G'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'H'.
    <fs_celldata> = gc_cc.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'I'.
    <fs_celldata> = '10000000073'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'J'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'N'.
    <fs_celldata> = '10000000074'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'O'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'T'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'U'.
    <fs_celldata> = 'EUR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'V'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'W'.
    <fs_celldata> = 'KG'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Y'.
    <fs_celldata> = 'X'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    "Fifth record
    ls_line-index = '5'.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = '4'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'B'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'C'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'D'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'G'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'H'.
    <fs_celldata> = '0002'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'I'.
    <fs_celldata> = '10000073'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'J'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'N'.
    <fs_celldata> = '10000000074'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'O'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'T'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'U'.
    <fs_celldata> = 'EUR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'V'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'W'.
    <fs_celldata> = 'KG'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Y'.
    <fs_celldata> = 'X'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    et_table = lt_mpa_asset_data.

  ENDMETHOD.

  METHOD fill_table_dile_with_long_date.

    DATA : ls_line           TYPE mpa_s_index_value_pair,
           ls_cell           TYPE mpa_s_index_value_pair,
           lt_mpa_asset_data TYPE mpa_t_index_value_pair.

    FIELD-SYMBOLS: <ft_linedata> TYPE mpa_t_index_value_pair,
                   <fs_celldata> TYPE any.


    "First record
    ls_line-index = gc_one.

    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    ls_cell-index = 'A'.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    <fs_celldata> = 'Asset Mass Transfer'.
*        IF iv_change_title_name IS NOT INITIAL.
    <fs_celldata> = 'Mass Transfer'.
*        ENDIF.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    "Second record
    ls_line-index = '2'.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = 'SLNO'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'B'.
    <fs_celldata> = 'BLDAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'C'.
    <fs_celldata> = 'BUDAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'D'.
    <fs_celldata> = 'BZDAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'E'.
    <fs_celldata> = 'SGTXT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'F'.
    <fs_celldata> = 'MONAT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'G'.
    <fs_celldata> = 'WWERT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'H'.
    <fs_celldata> = 'BUKRS'.
*        IF iv_unwanted_field IS NOT INITIAL.
    <fs_celldata> = 'BUKRS1'.
*        ENDIF.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'I'.
    <fs_celldata> = 'ANLN1'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'J'.
    <fs_celldata> = 'ANLN2'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'K'.
    <fs_celldata> = 'ACC_PRINCIPLE'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'L'.
    <fs_celldata> = 'AFABER'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'M'.
    <fs_celldata> = 'PBUKRS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'N'.
    <fs_celldata> = 'PANL1'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'O'.
    <fs_celldata> = 'PANL2'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'P'.
    <fs_celldata> = 'ANLKL'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Q'.
    <fs_celldata> = 'KOSTL'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'R'.
    <fs_celldata> = 'TEXT'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'S'.
    <fs_celldata> = 'TRAVA'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'T'.
    <fs_celldata> = 'ANBTR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'U'.
    <fs_celldata> = 'WAERS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'V'.
    <fs_celldata> = 'MENGE'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'W'.
    <fs_celldata> = 'MEINS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'X'.
    <fs_celldata> = 'PROZS'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Y'.
    <fs_celldata> = 'XANEU'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Z'.
    <fs_celldata> = 'RECID'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'AA'.
    <fs_celldata> = 'XBLNR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'AB'.
    <fs_celldata> = 'DZUONR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    "Third record
*        IF iv_no_data IS INITIAL.
*          IF iv_clear_techdesc IS INITIAL.
    ls_line-index = '3'.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.
*          ENDIF.
    "Fourth record
    ls_line-index = '4'.
*          IF  iv_increase_row IS NOT INITIAL.
*            ls_line-index = '5'.
*          ENDIF.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = '4'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'B'.
    <fs_celldata> = '2021070512345'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'C'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'D'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'G'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'H'.
    <fs_celldata> = gc_cc.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'I'.
    <fs_celldata> = '10000000073'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'J'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'N'.
    <fs_celldata> = '10000000074'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'O'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'T'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'U'.
    <fs_celldata> = 'EUR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'V'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'W'.
    <fs_celldata> = 'KG'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Y'.
    <fs_celldata> = 'X'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    "Fifth record
    ls_line-index = '5'.
    CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
    ASSIGN ls_line-value->* TO <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'A'.
    <fs_celldata> = '4'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'B'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'C'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'D'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'G'.
    <fs_celldata> = sy-datum.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'H'.
    <fs_celldata> = '0002'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'I'.
    <fs_celldata> = '10000073'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'J'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'N'.
    <fs_celldata> = '10000000074'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'O'.
    <fs_celldata> = ' '.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'T'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'U'.
    <fs_celldata> = 'EUR'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'V'.
    <fs_celldata> = gc_one.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'W'.
    <fs_celldata> = 'KG'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    CREATE DATA ls_cell-value TYPE string.
    ASSIGN ls_cell-value->* TO <fs_celldata>.
    ls_cell-index = 'Y'.
    <fs_celldata> = 'X'.
    INSERT ls_cell INTO TABLE <ft_linedata>.
    INSERT ls_line INTO TABLE lt_mpa_asset_data.

    et_table = lt_mpa_asset_data.

  ENDMETHOD.

METHOD get_value_count_initial.

"when
cl_abap_unit_assert=>assert_initial( mo_mpa_parse_util->get_value_count( it_cells_tab = VALUE #( ( `` ) ) ) ).

ENDMETHOD.

ENDCLASS.
