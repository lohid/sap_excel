**"* use this source file for your ABAP unit test classes
CLASS ltc_mpa_template_download_util DEFINITION DEFERRED.
CLASS ltc_download_util_unit_test DEFINITION DEFERRED.
CLASS zcl_template_download_util DEFINITION LOCAL FRIENDS ltc_mpa_template_download_util.
CLASS zcl_template_download_util DEFINITION LOCAL FRIENDS ltc_download_util_unit_test.


CLASS ltc_mpa_template_download_util DEFINITION FOR TESTING
  FINAL
  DURATION MEDIUM
  RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    CLASS-DATA:
      go_osql_environment       TYPE REF TO if_osql_test_environment,      "* Global instance of the OSQL test environment
      mo_template_download_util TYPE REF TO zcl_template_download_util.

    CLASS-METHODS:
      class_setup,                  "* Set up the OSQL test environment
      class_teardown.               "* Destroy the OSQL test environment

    CONSTANTS:
      gc_cc  TYPE bukrs VALUE '0001',
      gc_one TYPE char1 VALUE '1'.

    METHODS:

      "! Prepare the run before starting
      setup,

      "! Clear objects
      teardown,

      " Fill tables in OSQL environment
      fill_mock_tables,

      "Fill excel type data in the hash table
      fill_table_file_data
        IMPORTING
          iv_scen_type TYPE mpa_template_type
        EXPORTING
          et_table     TYPE mpa_t_index_value_pair,

      "! Test get instance method
      get_instance            FOR TESTING RAISING cx_static_check,

      "!Mass Transfer excel dowmload
      test_mtr_xlsx_download  FOR TESTING RAISING cx_static_check,

      "!Mass Transfer CSVC dowmload
      test_mtr_csvc_download  FOR TESTING RAISING cx_static_check,

      "!Mass Transfer CSVS dowmload
      test_mtr_csvs_download  FOR TESTING RAISING cx_static_check,

      "!Mass Creation excel dowmload
      test_mcr_xlsx_download  FOR TESTING RAISING cx_static_check,

      "!Mass Creation csvc dowmload
      test_mcr_csvc_download  FOR TESTING RAISING cx_static_check,

      "!Mass Creation csvs dowmload
      test_mcr_csvs_download  FOR TESTING RAISING cx_static_check,

      "!Mass Change excel dowmload
      test_mch_xlsx_download  FOR TESTING RAISING cx_static_check,

      "!Mass Change csvc dowmload
      test_mch_csvc_download  FOR TESTING RAISING cx_static_check,

      "!Mass Change csvs dowmload
      test_mch_csvs_download  FOR TESTING RAISING cx_static_check,

      "! Test mass transfer excel file download with data
      test_mtr_file_download  FOR TESTING RAISING cx_static_check,

      "! Test mass creation excel file download with data
      test_mcr_file_download  FOR TESTING RAISING cx_static_check,

      "! Test mass change excel file download with data
      test_mch_file_download  FOR TESTING RAISING cx_static_check,

      "! test generate excel data block
      test_generate_excel_data_block FOR TESTING RAISING cx_static_check,

      "!test generate excel template
      test_generate_excel_template   FOR TESTING RAISING cx_static_check,

      "!test generate csv template
      test_generate_csv_template     FOR TESTING RAISING cx_static_check,

      "!Test get template type method
      test_get_tmplt_type     FOR TESTING RAISING cx_static_check,

      "!Concatenate field line method test
      concatenate_field_line  FOR TESTING RAISING cx_static_check,

      "!Test render header and render title method
      test_render_header_title FOR TESTING RAISING cx_static_check,

      "! test fill label method
      test_fillin_label FOR TESTING RAISING cx_static_check,

      "!Test fill the field information method
      test_fillin_others           FOR TESTING RAISING cx_static_check,

      get_full_fields_mapping FOR TESTING RAISING cx_static_check,

      get_trans_mapping FOR TESTING RAISING cx_static_check,

      generate_csvc_template FOR TESTING RAISING cx_static_check,

      test_save_file FOR TESTING RAISING cx_static_check,

      download_csv_file FOR TESTING RAISING cx_static_check,

      save_csv_result FOR TESTING RAISING cx_static_check,

      concatenate_data_line_create FOR TESTING RAISING cx_static_check,

      concatenate_data_line_change FOR TESTING RAISING cx_static_check,

      concatenate_data_line_transfer FOR TESTING RAISING cx_static_check,

      concatenate_data_line_adj FOR TESTING RAISING cx_static_check,

      concatenate_data_line_ret FOR TESTING RAISING cx_static_check,

      get_full_fields_mapping_mr FOR TESTING RAISING cx_static_check,

      test_mch_file_download_csv FOR TESTING RAISING cx_static_check,

      test_save_file_excel FOR TESTING RAISING cx_static_check,

      test_save_file_csv_cr FOR TESTING RAISING cx_static_check,

      test_save_file_csv_ch FOR TESTING RAISING cx_static_check,

      test_save_file_csv_ma FOR TESTING RAISING cx_static_check,

      test_save_file_csv_rt FOR TESTING RAISING cx_static_check,

      assemble_field_mapping_neg FOR TESTING RAISING cx_static_check,

      format_cell_for_download_file FOR TESTING RAISING cx_static_check.

ENDCLASS.

CLASS ltc_download_util_unit_test DEFINITION FOR TESTING
RISK LEVEL HARMLESS
DURATION SHORT.

  PRIVATE SECTION .

    DATA :
      "! code under test
      mo_cut        TYPE REF TO zcl_template_download_util.

    METHODS :
      setup,
      format_cell_for_download_file FOR TESTING RAISING cx_static_check.

ENDCLASS.

CLASS ltc_download_util_unit_test IMPLEMENTATION.

  METHOD setup.

    mo_cut = NEW zcl_template_download_util( ).

  ENDMETHOD.

  METHOD format_cell_for_download_file.

*    DATA xml TYPE xstring.
*    xml =
*    cl_abap_codepage=>convert_to(
**      `<text>` &&
**      `<line>aaaa</line>` &&
**      `<line>bbbb</line>` &&
**      `<line>cccc</line>` &&
**      `</text>` ).
*            `<sheetData>` &&
*      `<c r="A6" s="8" t="s">` &&
*                `<v>61</v>` &&
*                `</c>` &&
*               `<c r="B6" s="9" t="s">` &&
*                `<v>61</v>` &&
*                `</c>` &&
*                `</sheetData>` ) .
*
*    "Access iXML-Library
*
**    DATA ixml TYPE REF TO if_ixml.
**    DATA stream_factory TYPE REF TO if_ixml_stream_factory.
*    DATA document TYPE REF TO cl_xml_document.
**    ixml = cl_ixml=>create( ).
**
**    stream_factory = ixml->create_stream_factory( ).
**
**    document = ixml->create_document( ).
*
***********************************************************************
*
*  document->parse_xstring( xml ).
*
*    DATA lo_custom_node TYPE REF TO if_ixml_node.
*
*    lo_custom_node = document->find_node( name = 'sheetData' ).
*
***********************************************************************
*
*
**    "Parse xml-data into dom
***    IF
**    ixml->create_parser(  document = document
**                             stream_factory = stream_factory
**                             istream = stream_factory->create_istream_xstring( string = xml ) )->parse( ) ."<> 0.
***      RETURN.
***    ENDIF.
**
**    "Iterate DOM and modify text elements
**    DATA iterator TYPE REF TO if_ixml_node_iterator.
**    iterator = document->create_iterator( ).
**    DATA node TYPE REF TO if_ixml_node.
***    DO.
**    node = iterator->get_next( ).
***      IF node IS INITIAL.
***        EXIT.
***      ENDIF.
***      IF node->get_type( ) = if_ixml_node=>co_node_text.
***        node->set_value( to_upper( node->get_value( ) ) ).
***      ENDIF.
***    ENDDO.
*
*
*    DATA : ls_block      TYPE if_salv_export_appendix=>ys_block,
*           lt_char_col   TYPE cl_mpa_template_download_util=>gty_tt_col_name,
*           lt_date_col   TYPE cl_mpa_template_download_util=>gty_tt_col_name,
*           lo_col_node   TYPE REF TO if_ixml_node,
*           lv_style_text TYPE string,
*           lv_style_date TYPE string,
*           lv_row_num    TYPE i.
**
*    DATA: lo_xlsx_doc           TYPE REF TO cl_xlsx_document,
*          lo_workbookpart       TYPE REF TO cl_xlsx_workbookpart,
*          lo_wordsheetparts     TYPE REF TO cl_openxml_partcollection,
*          lo_wordsheetpart      TYPE REF TO cl_openxml_part,
*          lo_sheet_content      TYPE xstring,
*          lo_xml_document       TYPE REF TO cl_xml_document,
*          lo_node               TYPE REF TO if_ixml_node,
*          lo_node_attr          TYPE REF TO if_ixml_node,
*          lo_node_rows          TYPE REF TO if_ixml_node_list,
*          lo_attrs_map          TYPE REF TO if_ixml_named_node_map,
*          lo_uri                TYPE REF TO cl_openxml_parturi,
*          lo_formarted          TYPE xstring,
*          lo_doc_parts          TYPE REF TO cl_openxml_partcollection,
*          lo_stylepart          TYPE REF TO cl_openxml_part,
*          lo_node_first_font    TYPE REF TO if_ixml_node,
*          lo_node_last_font     TYPE REF TO if_ixml_node,
*          lo_node_last_fill     TYPE REF TO if_ixml_node,
*          lo_node_first_element TYPE REF TO if_ixml_node,
*          lo_node_last_element  TYPE REF TO if_ixml_node,
*          lo_node_first_style   TYPE REF TO if_ixml_node,
*          lo_node_last_style    TYPE REF TO if_ixml_node,
*          lo_node_first_col     TYPE REF TO if_ixml_node,
*          lo_node_col           TYPE REF TO if_ixml_node,
*          iv_source_doc         TYPE xstring.
*
**lo_col_node = new cl_ixml_node( ).
**
**    lo_xlsx_doc       = new cl_xlsx_document( ).
**
**    lo_uri           = cl_openxml_parturi=>create_from_filename( iv_filename = '/xl/worksheets/sheet1.xml' ).
**    lo_wordsheetpart = lo_xlsx_doc->get_part_by_uri( ir_parturi = lo_uri ). "lo_wordsheetparts->get_part( 0 ).
**    lo_sheet_content = lo_wordsheetpart->get_data( ).
**
**    CREATE OBJECT lo_xml_document.
**    lo_xml_document->parse_xstring( iv_source_doc ).
**
**    " remove frozen setting
**    lo_node = lo_xml_document->find_node( name = 'selection' ).
***
*    mo_cut->format_cell_for_download_file(
*      EXPORTING
*        is_block      = ls_block
*        it_char_col   = lt_char_col
*        it_date_col   = lt_date_col
*        io_col_node   = lo_custom_node
*        iv_style_text = lv_style_text
*        iv_style_date = lv_style_date
*        iv_row_num = lv_row_num
*    ).

  ENDMETHOD.

ENDCLASS.

CLASS ltc_mpa_template_download_util IMPLEMENTATION.

  METHOD class_setup.
    DATA lt_tables TYPE if_osql_test_environment=>ty_t_sobjnames.

    " List of tables that will be abstracted by OSQL
    lt_tables = VALUE #( ( 'MPA_ASSET_DATA' )
                         ( 'DD04T' )
                         ( 'DD04L' )
                         ( 'DD07V' )
                         ( 'T002' )
     ).

    " Register tables with OSQL environment
    ltc_mpa_template_download_util=>go_osql_environment = cl_osql_test_environment=>create( lt_tables ).

    mo_template_download_util = NEW zcl_template_download_util( ).
    mo_template_download_util->mo_lcl_template_dl = NEW lcl_mpa_tmplt_dl_util_mock( ).

  ENDMETHOD.


  METHOD class_teardown.
    ltc_mpa_template_download_util=>go_osql_environment->disable_double_redirection( ).
    ltc_mpa_template_download_util=>go_osql_environment->destroy( ).
  ENDMETHOD.


  METHOD setup.
    fill_mock_tables( ).
    mo_template_download_util->mo_function_module = NEW ltd_function_module( ).
  ENDMETHOD.


  METHOD teardown.
    "" Delete all test data generated by the test methods
    go_osql_environment->clear_doubles( ).
  ENDMETHOD.


  METHOD get_instance.

    DATA(lr_temp_dl) = zcl_template_download_util=>get_instance( ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_temp_dl
                                             msg   = 'Error in creating object of the class'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

    zcl_template_download_util=>go_instance = lr_temp_dl.
    lr_temp_dl = zcl_template_download_util=>get_instance( ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_temp_dl
                                             msg   = 'Error in creating object of the class'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.


  "**================================ Mass Transfer template Download=============================**
  METHOD test_mtr_xlsx_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.


    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'XLSX' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'MT' )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.xlsx' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Transfer template is not created in .XLSX format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mtr_csvc_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'MT' )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.csv' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Transfer template is not created in .CSVC format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mtr_csvs_download .

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVS' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'MT' )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.csv' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Transfer template is not created in .CSVS format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  "**================================ Mass Creation template Download=============================**
  METHOD test_mcr_xlsx_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'XLSX' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CR' )
                                                       ( name = 'FileName'   value = 'AssetMassCreate_Template.xlsx' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Creation template is not created in .XLSX format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mcr_csvc_download .

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CR' )
                                                       ( name = 'FileName'   value = 'AssetMassCreate_Template.csv' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Creation template is not created in .CSVC format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mcr_csvs_download .

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVS' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CR' )
                                                       ( name = 'FileName'   value = 'AssetMassCreate_Template.csv' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Creation template is not created in .CSVS format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  "**================================ Mass Change template Download=============================**
  METHOD test_mch_xlsx_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'XLSX' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CH' )
                                                       ( name = 'FileName'   value = 'AssetMassChange_Template.xlsx' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Change template is not created in .XLSX format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mch_csvc_download .

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CH' )
                                                       ( name = 'FileName'   value = 'AssetMassChange_Template.csv' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Change template is not created in .CSVC format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mch_csvs_download .

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVS' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CH' )
                                                       ( name = 'FileName'   value = 'AssetMassChange_Template.csv' )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~generate_template(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
                                             msg   = 'Mass Change template is not created in .CSVS format'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.


  "*========================================= Test Uploaded file downloading=================================================*
  METHOD test_mtr_file_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'FileId'     value = '0000000168' )
                                                       ( name = 'TemplateId' value = 'MT'         )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.XLSX'  )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~download_file(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'Mass Transfer File is not downloaded in .XLSX format'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mcr_file_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'FileId'     value = '0000000169' )
                                                       ( name = 'TemplateId' value = 'CR'         )
                                                       ( name = 'FileName'   value = 'AssetMassCreate_Template.XLSX'  )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~download_file(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'Mass Transfer File is not downloaded in .XLSX format'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mch_file_download.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'FileId'     value = '0000000170' )
                                                       ( name = 'TemplateId' value = 'CH'         )
                                                       ( name = 'FileName'   value = 'AssetMassChange_Template.XLSX'  )
                                                     ).

    mo_template_download_util->if_mpa_template_download_util~download_file(
           EXPORTING
             it_key_tab               = lt_key_tab
             io_tech_request_context  = lo_tech_request_context
           IMPORTING
             er_stream                = lr_stream
             ev_filename              = lv_filename ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'Mass Transfer File is not downloaded in .XLSX format'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_mch_file_download_csv.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'FileId'     value = '0000000171' )
                                                       ( name = 'TemplateId' value = 'CH'         )
                                                       ( name = 'FileName'   value = 'AssetMassChange_Template.CSV'  )
                                                     ).

    TRY.
        mo_template_download_util->if_mpa_template_download_util~download_file(
               EXPORTING
                 it_key_tab               = lt_key_tab
                 io_tech_request_context  = lo_tech_request_context
               IMPORTING
                 er_stream                = lr_stream
                 ev_filename              = lv_filename ).

      CATCH cx_root INTO DATA(lx_exptn).


    ENDTRY.

    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_exptn
        msg              = 'No exception thrown !'
*    level            = if_abap_unit_constant=>severity-medium
        quit             = if_abap_unit_constant=>quit-no
*  RECEIVING
*    assertion_failed =
    ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'Mass Transfer File is not downloaded in .XLSX format'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.


  METHOD test_generate_excel_data_block.

    DATA: lv_file_name       TYPE string,
          lx_file            TYPE xstring,
          lv_scen_type       TYPE mpa_template_type,
          lt_line_index_data TYPE mpa_t_excel_doc_index.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'FileId'     value = '0000000168' )
                                                       ( name = 'TemplateId' value = 'MT'         )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.XLSX'  )
                                                     ).

    DATA(lt_line_index) = VALUE mpa_t_line_index( ( '4' ) ( '5' ) ).

    lt_line_index_data = VALUE mpa_t_excel_doc_index( ( begin_symbol = '2' header_techn = '2' data = lt_line_index )
                                                      ( begin_symbol = '3' header_techn = '2' data = lt_line_index )
                                                     ).

    fill_table_file_data( EXPORTING iv_scen_type = 'MT' IMPORTING et_table =  DATA(lt_excel_rows) ).

    mo_template_download_util->generate_excel_data_block(
                                 EXPORTING
                                   it_excel_rows = lt_excel_rows
                                   it_line_index = lt_line_index_data
                                 IMPORTING
                                   es_block = DATA(ls_block) ).

    cl_abap_unit_assert=>assert_not_initial( act  = ls_block
                                             quit = if_aunit_constants=>quit-no  ).

    CLEAR: lt_excel_rows, ls_block.
    fill_table_file_data( EXPORTING iv_scen_type = 'CR' IMPORTING et_table =  lt_excel_rows ).

    mo_template_download_util->generate_excel_data_block(
                                 EXPORTING
                                   it_excel_rows = lt_excel_rows
                                   it_line_index = lt_line_index_data
                                 IMPORTING
                                   es_block = ls_block ).

    cl_abap_unit_assert=>assert_not_initial( act  = ls_block
                                             quit = if_aunit_constants=>quit-no  ).


  ENDMETHOD.


  METHOD test_generate_excel_template.

    DATA: ls_block TYPE if_salv_export_appendix=>ys_block.

    ls_block = VALUE #( ordinal_number = 3 location = 'TOP' name = 'Template_Data'
                        cells = VALUE #( ( row_index = 0  column_index = 1 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  'JVU1')
                                         ( row_index = 0  column_index = 2 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  'TEST ASSET')
                                         ( row_index = 0  column_index = 3 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  'KG')
                                         ( row_index = 0  column_index = 4 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '155555555')
                                         ( row_index = 0  column_index = 5 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '316421')
                                         ( row_index = 0  column_index = 6 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '31641')
                                         ( row_index = 0  column_index = 7 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '100')
                                         ( row_index = 0  column_index = 8 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '200')
                                         ( row_index = 0  column_index = 9 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '0L')
                                         ( row_index = 0  column_index = 10 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '0L')
                                         ( row_index = 0  column_index = 11 row_span = 1 column_span = 0 content_type = 'text' formatting  = '' value =  '0L')
                                       ) ).

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'XLSX' )
                                                       ( name = 'Language'   value = ' ')
                                                       ( name = 'TemplateId' value = 'CR' )
                                                       ( name = 'FileName'   value = ' ' )
                                                     ).

    mo_template_download_util->get_trans_mapping(
      IMPORTING
        et_full_fields_mapping   = DATA(lt_field_mapping) ).

    mo_template_download_util->generate_excel_template( EXPORTING
                                                          it_field_mapping = lt_field_mapping
                                                          is_block         = ls_block
                                                        IMPORTING
                                                          er_stream        = DATA(lr_stream)
                                                          ev_filename      = DATA(lv_filename) ).

  ENDMETHOD.


  METHOD test_generate_csv_template.

    DATA lt_field_mapping TYPE gty_t_field_name_mappings.

    mo_template_download_util->generate_csv_template( EXPORTING
                                                        it_field_mapping = lt_field_mapping
                                                      IMPORTING
                                                        er_stream        = DATA(lr_stream)
                                                        ev_filename      = DATA(lv_filename) ).

    cl_abap_unit_assert=>assert_not_initial( act  = lr_stream
                                             quit = if_aunit_constants=>quit-no  ).

    cl_abap_unit_assert=>assert_equals( act  = lv_filename
                                        quit = if_aunit_constants=>quit-no
                                        exp  = 'Template.csv' ).
  ENDMETHOD.

  METHOD test_render_header_title.

    DATA: lt_fields_header TYPE gty_t_field_name_mappings,
          lt_blocks        TYPE if_salv_export_appendix=>yts_block.

    mo_template_download_util->render_header( EXPORTING
                                                it_fields_header = lt_fields_header
                                              CHANGING
                                                ct_blocks = lt_blocks ).

    CALL METHOD mo_template_download_util->render_title
      CHANGING
        ct_blocks = lt_blocks.

    cl_abap_unit_assert=>assert_not_initial( act   = lt_blocks
                                             msg   = 'Excel block is not created'
                                             level = if_abap_unit_constant=>severity-medium
                                             quit  = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_get_tmplt_type.

    DATA lt_template_type TYPE cl_mpa_asset_process_mpc=>tt_templatetype.
    lt_template_type = VALUE #(
                               ( templateid = 'MT' templatedesc = 'Mass Transfer' )
                               ( templateid = 'CR' templatedesc = 'Mass Creation' )
                               ( templateid = 'CH' templatedesc = 'Mass Change' )
                               ( templateid = 'RT' templatedesc = 'Mass Retirement' )
                              ).

    mo_template_download_util->if_mpa_template_download_util~get_template_type( IMPORTING et_entityset = DATA(lt_entityset) ).

    cl_abap_unit_assert=>assert_equals(
      act  = lt_entityset
      exp  = lt_template_type
      quit = if_aunit_constants=>quit-no
    ).
  ENDMETHOD.


  METHOD concatenate_field_line.

    DATA : lv_title          TYPE string,
           lt_fields_mapping TYPE gty_t_field_name_mappings,
           lv_delimiter      TYPE string VALUE  ','.

    mo_template_download_util->concatenate_field_line( EXPORTING
                                                         iv_title          = lv_title
                                                         it_fields_mapping = lt_fields_mapping
                                                         iv_delimiter      = lv_delimiter
                                                       IMPORTING
                                                         ev_line           = DATA(lv_header_str) ).

    cl_abap_unit_assert=>assert_not_initial( act  = lv_header_str
                                             quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD test_fillin_others.

    DATA: lt_required_fields_mapping TYPE gty_t_field_name_mappings.

    mo_template_download_util->fillin_others( CHANGING ct_fields_mapping = lt_required_fields_mapping ).

    cl_abap_unit_assert=>assert_initial( act  = lt_required_fields_mapping
                                         quit = if_aunit_constants=>quit-no ).

    lt_required_fields_mapping = VALUE #( ( position   = '2' label      = 'Company Code' stru_name  = 'COMP_CODE'  stru_type = 'BUKRS'  )
                                          ( position   = '2' label      = 'Company Code' stru_name  = 'COMP_CODE3' stru_type = 'BUKRS3' ) ).

    mo_template_download_util->fillin_others( CHANGING ct_fields_mapping = lt_required_fields_mapping ).

    LOOP AT lt_required_fields_mapping INTO DATA(ls_mapping).
      IF sy-tabix = 1.
        cl_abap_unit_assert=>assert_equals( act  = ls_mapping-length
                                            quit = if_aunit_constants=>quit-no
                                            exp = '0004' ).

        cl_abap_unit_assert=>assert_equals( act  = ls_mapping-data_type
                                            quit = if_aunit_constants=>quit-no
                                            exp = 'CHAR' ).
      ELSE.
        cl_abap_unit_assert=>assert_initial( act  = ls_mapping-length
                                            quit = if_aunit_constants=>quit-no ).

        cl_abap_unit_assert=>assert_initial( act  = ls_mapping-data_type
                                            quit = if_aunit_constants=>quit-no ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD test_fillin_label.

    DATA: lt_key_text_element TYPE mpa_t_textpool,
          lt_fields_mapping   TYPE gty_t_field_name_mappings,
          lt_exp_mapping      TYPE gty_t_field_name_mappings.

    lt_fields_mapping = VALUE #( ( stru_name = 'BUKRS' stru_type = 'BUKRS' label_type = 'BUKRS' )
                                 ( stru_name = 'ANLN1' stru_type = 'ANLN1' label_type = ' ' )
                                 ( stru_name = 'ANLN2' stru_type = 'ANLN2' label_type = 'HELL' )
                               ).

    lt_exp_mapping =  VALUE #( ( label = 'Company Code' stru_name = 'BUKRS' stru_type = 'BUKRS' label_type = 'BUKRS' )
                              ( label = '' stru_name = 'ANLN1' stru_type = 'ANLN1' label_type = ' ' )
                              ( label = '' stru_name = 'ANLN2' stru_type = 'ANLN2' label_type = 'HELL' )
                             ).

    mo_template_download_util->fillin_label(
                                 EXPORTING
                                   iv_language         = 'E'
                                 CHANGING
                                   ct_fields_mapping   = lt_fields_mapping ).

    cl_abap_unit_assert=>assert_equals( act  = lt_fields_mapping
                                        quit = if_aunit_constants=>quit-no
                                        exp = lt_exp_mapping ).

  ENDMETHOD.

  METHOD test_save_file.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    DATA: lt_mpa_transfer   TYPE TABLE OF mpa_s_asset_transfer,
          ls_mpa_asset_data TYPE mpa_asset_data,
          lv_csv_delimiter  TYPE string,
          lv_new_fileid     TYPE char10.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'MT' )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.csv' )
                                                     ).

    "Prapare the processed file data into required format
    zcl_template_download_util=>get_instance( )->save_result_file( EXPORTING it_table          = lt_mpa_transfer
                                                                                is_mpa_asset_data = ls_mpa_asset_data
                                                                                iv_csv_delimiter  = lv_csv_delimiter
                                                                      IMPORTING ev_fileid         = lv_new_fileid ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'File is not saved'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD test_save_file_csv_cr.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    DATA: lt_mpa_create     TYPE TABLE OF mpa_s_asset_create,
          ls_mpa_asset_data TYPE mpa_asset_data,
          lv_csv_delimiter  TYPE string,
          lv_new_fileid     TYPE char10.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CR' )
                                                       ( name = 'FileName'   value = 'AssetMassCreate_Template.csv' )
                                                     ).
    mo_template_download_util->gv_template_type = 'CR'.
    ls_mpa_asset_data-file_name = 'AssetMassCreate_Template.csv'.
    ls_mpa_asset_data-scen_type = 'CR'.

    "Prapare the processed file data into required format
    zcl_template_download_util=>get_instance( )->save_result_file( EXPORTING it_table          = lt_mpa_create
                                                                                is_mpa_asset_data = ls_mpa_asset_data
                                                                                iv_csv_delimiter  = lv_csv_delimiter
                                                                      IMPORTING ev_fileid         = lv_new_fileid ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'File is not saved'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD test_save_file_csv_ch.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    DATA: lt_mpa_change     TYPE TABLE OF mpa_s_asset_change,
          ls_mpa_asset_data TYPE mpa_asset_data,
          lv_csv_delimiter  TYPE string,
          lv_new_fileid     TYPE char10.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'CH' )
                                                       ( name = 'FileName'   value = 'AssetMassChange_Template.csv' )
                                                     ).
    mo_template_download_util->gv_template_type = 'CH'.
    ls_mpa_asset_data-file_name = 'AssetMassChange_Template.csv'.
    ls_mpa_asset_data-scen_type = 'CH'.

    "Prapare the processed file data into required format
    zcl_template_download_util=>get_instance( )->save_result_file( EXPORTING it_table          = lt_mpa_change
                                                                                is_mpa_asset_data = ls_mpa_asset_data
                                                                                iv_csv_delimiter  = lv_csv_delimiter
                                                                      IMPORTING ev_fileid         = lv_new_fileid ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'File is not saved'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD test_save_file_csv_ma.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    DATA: lt_mpa_adjustment TYPE TABLE OF mpa_s_asset_adjustment,
          ls_mpa_asset_data TYPE mpa_asset_data,
          lv_csv_delimiter  TYPE string,
          lv_new_fileid     TYPE char10.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'MA' )
                                                       ( name = 'FileName'   value = 'AssetMassAdjustment_Template.csv' )
                                                     ).
    mo_template_download_util->gv_template_type = 'MA'.
    ls_mpa_asset_data-file_name = 'AssetMassAdjustment_Template.csv'.
    ls_mpa_asset_data-scen_type = 'MA'.

    "Prapare the processed file data into required format
    zcl_template_download_util=>get_instance( )->save_result_file( EXPORTING it_table          = lt_mpa_adjustment
                                                                                is_mpa_asset_data = ls_mpa_asset_data
                                                                                iv_csv_delimiter  = lv_csv_delimiter
                                                                      IMPORTING ev_fileid         = lv_new_fileid ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'File is not saved'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD test_save_file_csv_rt.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    DATA: lt_mpa_retirement TYPE TABLE OF mpa_s_asset_retirement,
          ls_mpa_asset_data TYPE mpa_asset_data,
          lv_csv_delimiter  TYPE string,
          lv_new_fileid     TYPE char10.

    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'CSVC' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'TR' )
                                                       ( name = 'FileName'   value = 'AssetMassRetirement_Template.csv' )
                                                     ).
    mo_template_download_util->gv_template_type = 'RT'.
    ls_mpa_asset_data-file_name = 'AssetMassRetirement_Template.csv'.
    ls_mpa_asset_data-scen_type = 'RT'.

    "Prapare the processed file data into required format
    zcl_template_download_util=>get_instance( )->save_result_file( EXPORTING it_table          = lt_mpa_retirement
                                                                                is_mpa_asset_data = ls_mpa_asset_data
                                                                                iv_csv_delimiter  = lv_csv_delimiter
                                                                      IMPORTING ev_fileid         = lv_new_fileid ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'File is not saved'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.

  METHOD test_save_file_excel.

    DATA: lt_key_tab               TYPE /iwbep/t_mgw_name_value_pair,
          lo_tech_request_context  TYPE REF TO /iwbep/if_mgw_req_entity,
          lv_service_version       TYPE numc4,
          lv_service_document_name TYPE string,
          lr_stream                TYPE REF TO data,
          lv_filename              TYPE bapidocid.

    DATA: lt_mpa_transfer   TYPE TABLE OF mpa_s_asset_transfer,
          ls_mpa_asset_data TYPE mpa_asset_data,
          lv_csv_delimiter  TYPE string,
          lv_new_fileid     TYPE char10.

    "given
    mo_template_download_util->gv_template_type = 'MT'.
    mo_template_download_util->mt_param_tab = VALUE #( ( name = 'IsTemplate' value = 'X')
                                                       ( name = 'Mimetype'   value = 'XLSX' )
                                                       ( name = 'Language'   value = 'EN')
                                                       ( name = 'TemplateId' value = 'MT' )
                                                       ( name = 'FileName'   value = 'AssetMassTransfer_Template.xlsx' )
                                                     ).

    ls_mpa_asset_data-file_name = 'AssetMassTransfer_Template.xlsx'.
    ls_mpa_asset_data-scen_type = 'MT'.

    lt_mpa_transfer = VALUE #( ( slno = '01'
                                 bukrs = 'JVU1'
                                 anln1 = '123'
                                 anlkl = '300') ).

    "Prapare the processed file data into required format
    mo_template_download_util->if_mpa_template_download_util~save_result_file( EXPORTING it_table          = lt_mpa_transfer
                                                                                         is_mpa_asset_data = ls_mpa_asset_data
                                                                                         iv_csv_delimiter  = lv_csv_delimiter
                                                                               IMPORTING ev_fileid         = lv_new_fileid ).

*    cl_abap_unit_assert=>assert_not_initial( act   = lr_stream
*                                             msg   = 'File is not saved'
*                                             level = if_abap_unit_constant=>severity-medium
*                                             quit  = if_aunit_constants=>quit-no ).
  ENDMETHOD.


  "*======================================Filling data for unit testing================================================*
  "--------------------------------------------------------------------------------------------------------------------*
  METHOD fill_table_file_data.

    DATA : ls_line           TYPE mpa_s_index_value_pair,
           ls_cell           TYPE mpa_s_index_value_pair,
           lt_mpa_asset_data TYPE mpa_t_index_value_pair.

    FIELD-SYMBOLS: <ft_linedata> TYPE mpa_t_index_value_pair,
                   <fs_celldata> TYPE any.

    CASE iv_scen_type.
      WHEN 'MT'.

        "First record
        ls_line-index = gc_one.

        CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
        ASSIGN ls_line-value->* TO <ft_linedata>.
        ls_cell-index = 'A'.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        <fs_celldata> = 'Asset Mass Transfer'.
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
        <fs_celldata> = 'AFABE_POST'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'M'.
        <fs_celldata> = 'BUKRS1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'N'.
        <fs_celldata> = 'BF_ANLN1'.
        INSERT ls_cell INTO TABLE <ft_linedata>.
        CREATE DATA ls_cell-value TYPE string.
        ASSIGN ls_cell-value->* TO <fs_celldata>.
        ls_cell-index = 'O'.
        <fs_celldata> = 'BF_ANLN2'.
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

      WHEN 'CR'.

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

    ENDCASE.

    et_table = lt_mpa_asset_data.

  ENDMETHOD.


  METHOD fill_mock_tables.

    DATA: lt_dd04t TYPE TABLE OF dd04t.   "R/3 DD: Data element texts

    "R/3 DD: Data element texts
    lt_dd04t = VALUE #(
                        ( rollname = 'MPA_SLNO' ddlanguage = 'E' as4local = 'A' ddtext = 'Row Number' reptext = 'Mass Processing of Asset Row Number' scrtext_s = 'Row Num' scrtext_m = 'Row Number' scrtext_l = 'Row Number' )
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
                                ( mandt = sy-mandt ernam = 'C5232603' erdat = '20200325' erzet = '163151' file_id = '0000000168' file_name = 'AssetMassTransfer_Template.XLSX' scen_type = 'MT' file_status = '1')
                                ( mandt = sy-mandt ernam = 'C5232000' erdat = '20200325' erzet = '163151' file_id = '0000000169' file_name = 'AssetMassCreate_Template.XLSX'   scen_type = 'CR' file_status = '1')
                                ( mandt = sy-mandt ernam = 'C5232603' erdat = '20200325' erzet = '163151' file_id = '0000000170' file_name = 'AssetMassChange_Template.XLSX'   scen_type = 'CH' file_status = '1')
                                ( mandt = sy-mandt ernam = 'C5232603' erdat = '20200325' erzet = '163151' file_id = '0000000171' file_name = 'AssetMassChange_Template.CSV'   scen_type = 'CH' file_status = '1')
                              ).

    go_osql_environment->insert_test_data( lt_mpa_asset_data ).

    DATA: lt_dd04l TYPE TABLE OF dd04l.   "Data elements

    "Data elements
    lt_dd04l = VALUE #(
                        ( rollname = 'MPA_SLNO' as4local = 'A' domname = ' ' memoryid = ' ' logflag = ' ' headlen = '55' scrlen1 = '10' scrlen2 = '13' scrlen3 = '38' applclass = ' ' as4user = 'LISSITSYNA' as4date = '20200213' as4time = '142320'
                        dtelmaster = 'E' shlpname = ' ' shlpfield = ' ' deffdname = ' ' datatype = 'INT4' leng = '000010' outputlen = '000011' entitytab = ' ' refkind = ' ' )
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
  "*====================================================*====================================================*


  METHOD get_full_fields_mapping.

    DATA ls_asset_adjustment TYPE mpa_s_asset_adjustment.

    mo_template_download_util->gv_template_type = 'MA'.

    mo_template_download_util->get_full_fields_mapping(
      EXPORTING
        is_struct             = ls_asset_adjustment
      IMPORTING
        et_full_field_mapping = DATA(lt_field_mapping)
    ).

    cl_abap_unit_assert=>assert_not_initial( act  = lt_field_mapping
                                                 quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD get_full_fields_mapping_mr.

    DATA ls_asset_retirement TYPE mpa_s_asset_retirement.

    mo_template_download_util->gv_template_type = 'RT'.

    mo_template_download_util->get_full_fields_mapping(
      EXPORTING
        is_struct             = ls_asset_retirement
      IMPORTING
        et_full_field_mapping = DATA(lt_field_mapping)
    ).

    cl_abap_unit_assert=>assert_not_initial( act  = lt_field_mapping
                                                 quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD get_trans_mapping.

    mo_template_download_util->gv_template_type = 'MA'.

    mo_template_download_util->get_trans_mapping(
      IMPORTING
        et_full_fields_mapping = DATA(lt_field_mapping)
    ).

    cl_abap_unit_assert=>assert_not_initial( act  = lt_field_mapping
                                                 quit = if_aunit_constants=>quit-no ).

  ENDMETHOD.

  METHOD generate_csvc_template.


  ENDMETHOD.

  METHOD download_csv_file.

    "given
    DATA : ls_uploaded_file TYPE mpa_asset_data,
           lo_mpa_parse     TYPE REF TO if_mpa_xlsx_parse_util,
           lt_asset         TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data.

    lt_asset = VALUE #( ( mass_transfer_data = VALUE #( ( slno = '1' )  ) ) ).

    lo_mpa_parse ?= cl_abap_testdouble=>create( 'IF_MPA_XLSX_PARSE_UTIL' ).

    DATA(lo_td_config) = cl_abap_testdouble=>configure_call( double = lo_mpa_parse )->ignore_all_parameters( ).

    lo_td_config->set_parameter( name          = 'ev_mpa_type'
                                 value         = 'MT' ).

    lo_td_config->set_parameter( name          = 'et_asset'
                                 value         = lt_asset ).

    lo_td_config->set_parameter( name          = 'ev_seperator'
                                 value         = ';' ).

    lo_td_config->set_parameter( name          = 'ev_status_flag'
                                 value         = abap_true ).

    lo_mpa_parse->parse_csv( ix_file     = ls_uploaded_file-file_data ).


    "when
    mo_template_download_util->get_csv_file_with_data(
      EXPORTING
        io_mpa_pasrs     = lo_mpa_parse
        is_uploaded_file = ls_uploaded_file
  IMPORTING
    ev_filename      = DATA(lv_filename) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = lv_filename
        exp                  = 'Template.csv'
        msg                  = 'Incorrect File name'
        quit                 = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

  METHOD save_csv_result.

    "given
    DATA(lt_transfer) = VALUE mpa_t_asset_transfer( (  ) ).
    DATA(ls_mpa_asset_data) = VALUE mpa_asset_data(
                                   mandt = sy-mandt
                                   file_id = '0000001421'
                                   ernam = 'C5232603'
                                   erdat = '20200828'
                                   erzet = '090143'
                                   timestamp = '20200828070144'
                                   file_name = 'AssetMassTransfer_TemplateXLSX.csv'
                                   scen_type = 'MT'
                                   file_status = '1'
                                   file_data = '504B03041400060008000000210041378'
                                 ).



    "when
    mo_template_download_util->if_mpa_template_download_util~save_result_file(
      EXPORTING
        it_table          = lt_transfer
        is_mpa_asset_data = ls_mpa_asset_data
        iv_csv_delimiter  = ';'
      IMPORTING
    ev_fileid         = DATA(file_id) ).

    "then
    cl_abap_unit_assert=>assert_equals(
      EXPORTING
        act                  = file_id
        exp                  = '1234567890'
        msg                  = 'Incorrect File id'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD concatenate_data_line_create.

    "given
    DATA(ls_mpa_data) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( mass_create_data = VALUE #( ( slno = '123'
                                                                                                          bukrs = 'JVU1'
                                                                                                          anln1 = '1000'  ) ) ).

    "when
    mo_template_download_util->concatenate_data_line(
      EXPORTING
        is_mpa_data  = ls_mpa_data
        iv_delimiter = ','
        iv_mpa_type  = 'CR'
      IMPORTING
        ev_data_line = DATA(lv_data_line)
    ).

    "then
    cl_abap_unit_assert=>assert_char_cp(
      EXPORTING
        act                  = lv_data_line
        exp                  =
'123 ,,JVU1,,,1000,,,,,,0000-00-00,,,0000-00-00,,,,,,,,,,,,,,,0.00 ,,,,0000-00-00,0000,0.00 ,,,0000-00-00,,0000-00-00,,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00'
        msg                  = 'issue with concatination for create'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD concatenate_data_line_change.

    "given
    DATA(ls_mpa_data) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( mass_change_data = VALUE #( ( slno = '123'
                                                                                                          bukrs = 'JVU1'
                                                                                                          anln1 = '1000'  ) ) ).

    "when
    mo_template_download_util->concatenate_data_line(
      EXPORTING
        is_mpa_data  = ls_mpa_data
        iv_delimiter = ','
        iv_mpa_type  = 'CH'
      IMPORTING
        ev_data_line = DATA(lv_data_line)
    ).

    "then
    cl_abap_unit_assert=>assert_char_cp(
      EXPORTING
        act                  = lv_data_line
        exp                  =
'123 ,,JVU1,1000,,,,,,,0000-00-00,,,0000-00-00,,,,,,,,,,,,,,0.00 ,,,,0000-00-00,0000,0.00 ,,,0000-00-00,,0000-00-00,,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00,00,,000,000,0000-00-00'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD concatenate_data_line_transfer.

    "given
    DATA(ls_mpa_data) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( mass_transfer_data = VALUE #( ( slno = '123'
                                                                                                            bukrs = 'JVU1'
                                                                                                            anln1 = '1000'  ) ) ).

    "when
    mo_template_download_util->concatenate_data_line(
      EXPORTING
        is_mpa_data  = ls_mpa_data
        iv_delimiter = ','
        iv_mpa_type  = 'MT'
      IMPORTING
        ev_data_line = DATA(lv_data_line)
    ).

    "then
    cl_abap_unit_assert=>assert_char_cp(
      EXPORTING
        act                  = lv_data_line
        exp                  = '123 ,,,0000-00-00,0000-00-00,0000-00-00,,00,0000-00-00,JVU1,1000,,,00,,,,,,,,0.00 ,,0.000 ,,0.00 ,,,,'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD concatenate_data_line_adj.

    "given
    DATA(ls_mpa_data) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( mass_adjustment_data = VALUE #( ( slno = '123'
                                                                                                            bukrs = 'JVU1'
                                                                                                            anln1 = '1000'  ) ) ).

    "when
    mo_template_download_util->concatenate_data_line(
      EXPORTING
        is_mpa_data  = ls_mpa_data
        iv_delimiter = ','
        iv_mpa_type  = 'MA'
      IMPORTING
        ev_data_line = DATA(lv_data_line)
    ).

    "then
    cl_abap_unit_assert=>assert_char_cp(
      EXPORTING
        act                  = lv_data_line
        exp                  = '123 ,,JVU1,1000,,,,00,0000-00-00,0000-00-00,0000-00-00,0.00 ,,,,,,,,'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD concatenate_data_line_ret.

    "given
    DATA(ls_mpa_data) = VALUE cl_mpa_asset_process_dpc_ext=>ty_s_file_data( mass_retirement_data = VALUE #( ( slno = '123'
                                                                                                            bukrs = 'JVU1'
                                                                                                            anln1 = '1000'  ) ) ).

    "when
    mo_template_download_util->concatenate_data_line(
      EXPORTING
        is_mpa_data  = ls_mpa_data
        iv_delimiter = ','
        iv_mpa_type  = 'RT'
      IMPORTING
        ev_data_line = DATA(lv_data_line)
    ).

    "then
    cl_abap_unit_assert=>assert_char_cp(
      EXPORTING
        act                  = lv_data_line
        exp                  = '123 ,,JVU1,1000,,,,00,,,0000-00-00,0000-00-00,00,0000-00-00,,0.00 ,0.0000 ,0.0000 ,,0.000 ,,0.00 ,,00,0000-00-00,,,,00000000,,,,,,,,,,,'
        quit                 = if_abap_unit_constant=>quit-no ).


  ENDMETHOD.

  METHOD assemble_field_mapping_neg.

    "given
    DATA ls_mpa_struct TYPE mpa_s_asset_create.


    "when
    mo_template_download_util->assemble_fields_mapping(
      EXPORTING
        is_params                  = VALUE #( bukrs = 'JVU1' langu = 'EN')
        is_struct                  = ls_mpa_struct
        it_full_fields_mapping     = VALUE #( (  ) )
*    it_excepted_fields_mapping =
*    it_expected_fields_mapping =
      IMPORTING
        et_required_fields_mapping = DATA(lt_act)
    ).
*CATCH /iwbep/cx_mgw_med_exception.


    "then
    cl_abap_unit_assert=>assert_initial(
      EXPORTING
        act              = lt_act
        msg              = 'Field mapping not initial as expected !'
        quit             = if_abap_unit_constant=>quit-no
    ).

  ENDMETHOD.

  METHOD format_cell_for_download_file.

    "given
    DATA : ls_block    TYPE if_salv_export_appendix=>ys_block,
           lt_char_col TYPE zcl_template_download_util=>gty_tt_col_name,
           lt_date_col TYPE zcl_template_download_util=>gty_tt_col_name,
           lo_col_node TYPE REF TO if_ixml_node.


    "when
    TRY.
        mo_template_download_util->format_cell_for_download_file(
          EXPORTING
            is_block      = ls_block
            it_char_col   = lt_char_col
            it_date_col   = lt_date_col
            io_col_node   = lo_col_node
            iv_style_text = ''
            iv_style_date = ''
*    iv_row_num    = 0
        ).

      CATCH cx_root INTO DATA(lx_exp).
    ENDTRY.

    "then
    cl_abap_unit_assert=>assert_bound(
      EXPORTING
        act              = lx_exp
        msg              = 'No exception thrown'
        quit             = if_abap_unit_constant=>quit-no ).

  ENDMETHOD.

ENDCLASS.
