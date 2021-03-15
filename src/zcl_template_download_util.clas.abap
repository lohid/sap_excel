CLASS zcl_template_download_util DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES if_mpa_template_download_util .

    TYPES:
      BEGIN OF gty_s_struct_properties,
        field_name   TYPE  name_feld,
        data_element TYPE rollname,
        data_type    TYPE datatype_d,
        length       TYPE ddleng,
      END OF gty_s_struct_properties .
    TYPES:
      gty_t_struct_properties TYPE STANDARD TABLE OF gty_s_struct_properties WITH NON-UNIQUE DEFAULT KEY .

    CLASS-METHODS get_instance
      RETURNING
        VALUE(ro_instance) TYPE REF TO if_mpa_template_download_util .
    METHODS constructor .

  PROTECTED SECTION.

  PRIVATE SECTION.

    TYPES gty_s_mpa_transfer TYPE mpa_s_asset_transfer .
    TYPES:
      BEGIN OF gty_st_col_name,
        name TYPE string,
      END OF gty_st_col_name .
    TYPES:
      gty_t_mpa_transfer TYPE STANDARD TABLE OF gty_s_mpa_transfer,
      gty_tt_col_name    TYPE STANDARD TABLE OF gty_st_col_name WITH DEFAULT KEY.

    METHODS format_cell_for_download_file
      IMPORTING
        is_block      TYPE if_salv_export_appendix=>ys_block
        it_char_col   TYPE gty_tt_col_name
        it_date_col   TYPE gty_tt_col_name
        io_col_node   TYPE REF TO if_ixml_node
        iv_style_text TYPE string
        iv_style_date TYPE string
        iv_row_num    TYPE i DEFAULT 0.
*    METHODS format_doc_download_file
*      IMPORTING
*        !iv_source_doc    TYPE xstring
*        !it_fields_header TYPE gty_t_field_name_mappings OPTIONAL
*        is_block          TYPE if_salv_export_appendix=>ys_block OPTIONAL
*      EXPORTING
*        !ev_target_doc    TYPE xstring
*      RAISING
*        cx_openxml_format
*        cx_openxml_not_found .

    DATA:
      gv_template_type TYPE c LENGTH 2 .
    DATA gv_title TYPE string .
    DATA mt_param_tab TYPE /iwbep/t_mgw_name_value_pair .
    DATA mo_lcl_template_dl TYPE REF TO lif_mpa_template_download_util .
    DATA mo_function_module TYPE REF TO lif_function_module.
    CLASS-DATA go_instance TYPE REF TO if_mpa_template_download_util .
    CONSTANTS gc_template TYPE dd07v-domname VALUE 'MPA_TEMPLATE_TYPE' ##NO_TEXT.
    CONSTANTS gc_comment_symbol TYPE string VALUE '//' ##NO_TEXT.
    CONSTANTS:
      "! Supported file formats and file extensions
      BEGIN OF gc_file_format,
        xlsx  TYPE string VALUE '.XLSX',
        xls   TYPE string VALUE '.XLS',
        excel TYPE string VALUE 'XLSX',
        csv   TYPE string VALUE 'CSV',
        csvc  TYPE string VALUE 'CSVC',
        csvs  TYPE string VALUE 'CSVS',
      END OF gc_file_format .
    CONSTANTS:
      "! UI Parameter from io_tech_request_context object
      BEGIN OF gc_param_name,
        fileid       TYPE string VALUE 'FileId',
        filename     TYPE string VALUE 'FileName',
        templateid   TYPE string VALUE 'TemplateId',
        language     TYPE string VALUE 'Language',
        mimetype     TYPE string VALUE 'Mimetype',
        templatetype TYPE string VALUE 'TemplateType',
      END OF gc_param_name .
    CONSTANTS:
      "! Excel blocks for title, tables header and data
      BEGIN OF gc_excel_blocks,
        template_title  TYPE string VALUE 'Template_Title',
        template_header TYPE string VALUE 'Template_Header',
        template_data   TYPE string VALUE 'Template_Data',
      END OF gc_excel_blocks .
    CONSTANTS:
      "! Mime types for file download
      BEGIN OF gc_mime_type,
        app_excel TYPE string VALUE 'APPLICATION/XLSX',
        app_csv   TYPE string VALUE 'APPLICATION/CSV',
      END OF gc_mime_type .
    CONSTANTS:
      "! File names if not passed from UI
      BEGIN OF gc_file_name,
        excel_file TYPE bapidocid VALUE 'Template.xlsx',
        csv_file   TYPE bapidocid VALUE 'Template.csv',
      END OF gc_file_name .

    "! get the label of each field
    METHODS fillin_label
      IMPORTING
        !iv_language       TYPE lang
      CHANGING
        !ct_fields_mapping TYPE gty_t_field_name_mappings .
    "! Fill field related other information like length of fields
    METHODS fillin_others
      CHANGING
        !ct_fields_mapping TYPE gty_t_field_name_mappings .
    "! Fill the data type of the field
    METHODS find_field_type
      IMPORTING
        !is_stru           TYPE any
      CHANGING
        !ct_fields_mapping TYPE gty_t_field_name_mappings .
    "! Format excel template, adjust row height, column length, column formatting
    METHODS format_doc
      IMPORTING
        !iv_source_doc       TYPE xstring
        !it_fields_header    TYPE gty_t_field_name_mappings OPTIONAL
        is_block             TYPE if_salv_export_appendix=>ys_block OPTIONAL
        iv_file_download_ind TYPE abap_bool OPTIONAL
      EXPORTING
        !ev_target_doc       TYPE xstring
      RAISING
        cx_openxml_format
        cx_openxml_not_found .
    "! Format excel template label
    METHODS format_label
      CHANGING
        !ct_fields_mapping TYPE gty_t_field_name_mappings .
    "! Generate excel template using XML
    METHODS generate_excel_template
      IMPORTING
        !it_field_mapping    TYPE gty_t_field_name_mappings
        it_asset_data        TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data   OPTIONAL
        iv_file_download_ind TYPE abap_bool OPTIONAL
      EXPORTING
        !er_stream           TYPE REF TO data
        !ev_filename         TYPE bapidocid
      RAISING
        cx_openxml_format
        cx_openxml_not_found
        cx_salv_export_error .
    "! Attach the excel template reference to download
    METHODS copy_data_to_ref
      IMPORTING
        !is_data TYPE any
      CHANGING
        !cr_data TYPE REF TO data .
    "! Get the field mapping for the template
    METHODS get_trans_mapping
      IMPORTING
        iv_status_slno_flag     TYPE abap_bool OPTIONAL
      EXPORTING
        !et_full_fields_mapping TYPE gty_t_field_name_mappings
      RAISING
        /iwbep/cx_mgw_med_exception .
    "! Assemble the structure of all required field in template based on scenario
    METHODS assemble_fields_mapping
      IMPORTING
        !is_params                  TYPE mpa_s_filter_ui
        !is_struct                  TYPE any
        !it_full_fields_mapping     TYPE gty_t_field_name_mappings
        !it_excepted_fields_mapping TYPE string_table OPTIONAL
        !it_expected_fields_mapping TYPE string_table OPTIONAL
      EXPORTING
        !et_required_fields_mapping TYPE gty_t_field_name_mappings
      RAISING
        /iwbep/cx_mgw_med_exception .
    "! Render the title of the excel template
    METHODS render_title
      CHANGING
        !ct_blocks TYPE if_salv_export_appendix=>yts_block .
    "! Render the technical field name and label of the excel template
    METHODS render_header
      IMPORTING
        !it_fields_header TYPE gty_t_field_name_mappings
      CHANGING
        VALUE(ct_blocks)  TYPE if_salv_export_appendix=>yts_block .
    "! Get the fields list which is not needed in the excel template but there in the structure
    METHODS get_excepted_fields
      RETURNING
        VALUE(rt_excepted_fields) TYPE string_table .
    "! Generate the CSV template ( Separator ',' and ';')
    METHODS generate_csv_template
      IMPORTING
        !it_field_mapping TYPE gty_t_field_name_mappings
        !iv_delimiter     TYPE string DEFAULT ','
      EXPORTING
        !er_stream        TYPE REF TO data
        !ev_filename      TYPE bapidocid .
    "! Get the basic settings for fields in template ( Mandatory field settings based on scenario)
    METHODS get_full_fields_mapping
      IMPORTING
        !is_struct             TYPE any
      EXPORTING
        !et_full_field_mapping TYPE gty_t_field_name_mappings .
    "! Concatenate the text for CSV format
    METHODS concatenate_field_line
      IMPORTING
        !iv_title          TYPE string OPTIONAL
        !it_fields_mapping TYPE gty_t_field_name_mappings
        !iv_delimiter      TYPE string DEFAULT ','
      EXPORTING
        !ev_line           TYPE string .
    "! Get the structure details
    METHODS get_struct_properties
      IMPORTING
        !iv_struct_name       TYPE any
      EXPORTING
        !et_struct_properties TYPE gty_t_struct_properties .
    "! get the commented text for excel template
    METHODS get_comment_text
      EXPORTING
        !et_comment TYPE string_table .
    "! Generate the excel data block to download the already uploaded file
    METHODS generate_excel_data_block
      IMPORTING
        it_excel_rows TYPE mpa_t_index_value_pair
        it_line_index TYPE mpa_t_excel_doc_index
      EXPORTING
        es_block      TYPE if_salv_export_appendix=>ys_block .
    METHODS concatenate_data_line
      IMPORTING
        is_mpa_data  TYPE cl_mpa_asset_process_dpc_ext=>ty_s_file_data
        iv_delimiter TYPE string
        iv_mpa_type  TYPE mpa_template_type
      EXPORTING
        ev_data_line TYPE string.
    METHODS generate_csv_download_file
      IMPORTING
        it_field_mapping TYPE gty_t_field_name_mappings
        iv_delimiter     TYPE string
        is_data          TYPE cl_mpa_asset_process_dpc_ext=>ty_s_file_data
        iv_mpa_type      TYPE mpa_template_type
      EXPORTING
        er_stream        TYPE REF TO data
        ev_filename      TYPE bapidocid .

    METHODS get_csv_file_with_data
      IMPORTING
        io_mpa_pasrs     TYPE REF TO if_mpa_xlsx_parse_util
        is_uploaded_file TYPE mpa_asset_data
      EXPORTING
        ev_filename      TYPE bapidocid
        er_stream        TYPE REF TO data
      RAISING
        /iwbep/cx_mgw_med_exception.
**********************************************************************
    METHODS zcreate_info_for_new_sheet
      IMPORTING
                iv_sheet_name        TYPE string
                io_worksheet_part    TYPE REF TO cl_xlsx_worksheetpart
                io_workbookpart      TYPE REF TO cl_xlsx_workbookpart
      RETURNING VALUE(rs_sheet_info) TYPE cl_ehfnd_xlsx=>gty_s_sheet_info
      RAISING   cx_dynamic_check.
    METHODS zupdate_wb_xml_after_add_sheet
      IMPORTING
        is_sheet_info   TYPE cl_ehfnd_xlsx=>gty_s_sheet_info
        io_workbookpart TYPE REF TO cl_xlsx_workbookpart.
    METHODS set_stylesxml
      IMPORTING
        io_xlsx_doc         TYPE REF TO cl_xlsx_document
        io_xml_document     TYPE REF TO cl_xml_document
      RETURNING
        VALUE(ro_stylepart) TYPE REF TO cl_openxml_part
      RAISING
        cx_openxml_format
        cx_openxml_not_found.
    METHODS create_excel_with_sheets
      IMPORTING it_fields_header     TYPE gty_t_field_name_mappings
                it_asset_data        TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data
      RETURNING VALUE(rv_excel_file) TYPE xstring
      RAISING
                cx_dynamic_check
                cx_openxml_format
                cx_openxml_not_allowed
                cx_openxml_not_found.
    METHODS get_asset_data
      IMPORTING
        ix_file         TYPE xstring
      RETURNING
        VALUE(r_result) TYPE cl_mpa_asset_process_dpc_ext=>ty_t_file_data.
ENDCLASS.



CLASS zcl_template_download_util IMPLEMENTATION.


  METHOD render_title.

    DATA: ls_block          TYPE if_salv_export_appendix=>ys_block,
          ls_cells          TYPE TABLE OF if_salv_export_appendix=>ys_cell,
          ls_cell           LIKE LINE OF ls_block-cells,
          ls_formatting     TYPE if_salv_export_appendix=>ys_cell_formatting,
          ls_formatting_cmt TYPE if_salv_export_appendix=>ys_cell_formatting,
          lt_comment        TYPE string_table,
          lv_row_index      TYPE i VALUE 1.

    " Set title formatting
    ls_formatting = VALUE #( is_bold              = abap_true
                             horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-forced_left
                             vertical_alignment   = if_salv_export_appendix=>cs_vertical_alignment-center
                             background_color     = 'FFFFC000'
                           ).

    " Set Comments formatting
    ls_formatting_cmt = VALUE #( is_bold              = abap_false
                                 horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-forced_left
                                 vertical_alignment   = if_salv_export_appendix=>cs_vertical_alignment-center
                               ).

    " create title of excel template
    ls_cells = VALUE #( ( row_index    = lv_row_index
                          column_index = 1
                          row_span     = 0
                          column_span  = 0
                          content_type = if_salv_export_appendix=>cs_cell_content_type-text
                          formatting   = ls_formatting
                          value        = gv_title
                       ) ).

    "Add 1 for next row
    lv_row_index = lv_row_index + 1.

    " get the commented text for excel template
    get_comment_text( IMPORTING et_comment = lt_comment ).

    LOOP AT lt_comment INTO DATA(lv_comment).
      CONCATENATE gc_comment_symbol lv_comment INTO DATA(lv_remark) SEPARATED BY space.

      " Create cell for comments
      ls_cell = VALUE #( row_index    = lv_row_index
                         column_index = 1
                         row_span     = 0
                         column_span  = 0
                         content_type = if_salv_export_appendix=>cs_cell_content_type-text
                         formatting   = ls_formatting_cmt
                         value        = lv_remark ).

      lv_row_index = lv_row_index + 1.
      APPEND ls_cell TO ls_cells.
    ENDLOOP.

    " Append cells to block on top of file
    ls_block = VALUE #( ordinal_number = 1
                        location       = if_salv_export_appendix=>cs_appendix_location-top
                        name           = gc_excel_blocks-template_title
                        cells          = ls_cells
                      ).

    INSERT ls_block INTO ct_blocks INDEX 1.

  ENDMETHOD.


  METHOD render_header.

    DATA: ls_block             TYPE if_salv_export_appendix=>ys_block,
          ls_cells             TYPE TABLE OF if_salv_export_appendix=>ys_cell,
          ls_cell              LIKE LINE OF ls_block-cells,
          ls_struct_formatting TYPE if_salv_export_appendix=>ys_cell_formatting,
          ls_label_formatting  TYPE if_salv_export_appendix=>ys_cell_formatting,
          ls_value_formatting  TYPE if_salv_export_appendix=>ys_cell_formatting,
          ls_field_header      TYPE gty_s_field_name_mapping,
          lv_column_index      TYPE i VALUE 1,
          lv_row_index         TYPE i VALUE 0,
          lv_value_row_index   TYPE i.

    " Technical field name formatting
    ls_struct_formatting = VALUE #( is_bold              = abap_false
                                    horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-center
                                    vertical_alignment   = if_salv_export_appendix=>cs_vertical_alignment-center
                                    foreground_color     = 'FF000000'
                                    background_color     = 'FFC6E2FF'  ).

    " Label formatting
    ls_label_formatting = VALUE #( is_bold              = abap_true
                                   horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-center
                                   vertical_alignment   = if_salv_export_appendix=>cs_vertical_alignment-center
                                   foreground_color     = 'FF000000'
                                   background_color     = 'FFC6E2FF'
*                                   is_text_wrapping_enabled = abap_true
                                 ).

    " Block for template header
    ls_block = VALUE #( ordinal_number = 2
                        location       = if_salv_export_appendix=>cs_appendix_location-top
                        name           = gc_excel_blocks-template_header
                        cells          = ls_cells
                      ).

    " Create the technical field row for the template
    LOOP AT it_fields_header INTO ls_field_header.

      IF ls_field_header-stru_name IS NOT INITIAL.
        ls_cell = VALUE #( row_index    = lv_row_index
                           row_span     = 1
                           column_index = lv_column_index
                           column_span  = 0
                           content_type = if_salv_export_appendix=>cs_cell_content_type-text
                           formatting   = ls_struct_formatting
                           value        = ls_field_header-stru_name
                         ).
        APPEND ls_cell TO ls_block-cells.
      ENDIF.
      lv_column_index = lv_column_index + 1.       "Move to next field

    ENDLOOP.


    lv_column_index = 1.

    "Create the label row for the template
    LOOP AT it_fields_header INTO ls_field_header.

      IF ls_field_header-f_label IS NOT INITIAL.
        ls_cell = VALUE #( row_index    = lv_row_index + 1
                           column_index = lv_column_index
                           row_span     = 1
                           column_span  = 0
                           content_type = if_salv_export_appendix=>cs_cell_content_type-text
                           formatting   = ls_label_formatting
                           value        = ls_field_header-f_label ).

        APPEND ls_cell TO ls_block-cells.
      ENDIF.

      lv_column_index = lv_column_index + 1.       " Move to next field
    ENDLOOP.

    INSERT ls_block INTO ct_blocks INDEX 2.

  ENDMETHOD.


  METHOD get_trans_mapping.

    DATA: ls_params              TYPE mpa_s_filter_ui,
          ls_field_mapping       TYPE gty_s_field_name_mapping,
          lt_key_element         TYPE cl_mpa_asset_process_mpc=>tt_text_elements,
          lt_selected_element    TYPE string_table,
          lt_key_text_element    TYPE mpa_t_textpool,
          ls_mass_transfer       TYPE mpa_s_asset_transfer,
          ls_mass_create         TYPE mpa_s_asset_create,
          ls_mass_change         TYPE mpa_s_asset_change,
          ls_mass_adjustment     TYPE mpa_s_asset_adjustment,
          ls_mass_retirement     TYPE mpa_s_asset_retirement,
          lt_full_fields_mapping TYPE gty_t_field_name_mappings,
          lt_excepted_fields     TYPE string_table.

    FIELD-SYMBOLS: <fs_field_mapping> TYPE gty_s_field_name_mapping.


    " Assign the scenario type to global variable for further uses
    READ TABLE mt_param_tab INTO DATA(ls_template_type) WITH KEY name = gc_param_name-templateid.
    IF sy-subrc IS INITIAL.
      CLEAR: gv_template_type.
      gv_template_type = ls_template_type-value.
    ENDIF.

    READ TABLE mt_param_tab INTO DATA(ls_language) WITH KEY name = gc_param_name-language.
    IF sy-subrc IS INITIAL.
      ls_params-langu = sy-langu.
    ELSE.
      TRANSLATE ls_language-value TO UPPER CASE.
      SELECT SINGLE spras INTO ls_params-langu FROM t002 WHERE laiso = ls_language-value . "#EC CI_NOORDER
    ENDIF.

    SET LANGUAGE ls_params-langu.

    lt_excepted_fields = COND #( WHEN iv_status_slno_flag = abap_false
                                 THEN get_excepted_fields( ) ).

    CASE gv_template_type.

      WHEN if_mpa_output=>gc_mpa_scen-transfer.  "Mass Transfer Scenario

        " Get full fields with output order
        get_full_fields_mapping( EXPORTING
                                   is_struct             = ls_mass_transfer
                                 IMPORTING
                                   et_full_field_mapping = lt_full_fields_mapping ).

        " Find required fields for Mass Asset Transfer
        assemble_fields_mapping( EXPORTING
                                   is_params                  = ls_params
                                   is_struct                  = ls_mass_transfer
                                   it_full_fields_mapping     = lt_full_fields_mapping
                                   it_excepted_fields_mapping = lt_excepted_fields
                                 IMPORTING
                                   et_required_fields_mapping = et_full_fields_mapping ).

      WHEN if_mpa_output=>gc_mpa_scen-create.   "Mass Create Scenario

        " Get full fields with output order
        get_full_fields_mapping( EXPORTING
                                   is_struct             = ls_mass_create
                                 IMPORTING
                                   et_full_field_mapping = lt_full_fields_mapping ).

        " Find required fields for Mass Asset Creation
        assemble_fields_mapping( EXPORTING
                                   is_params                  = ls_params
                                   is_struct                  = ls_mass_create
                                   it_full_fields_mapping     = lt_full_fields_mapping
                                   it_excepted_fields_mapping = lt_excepted_fields
                                 IMPORTING
                                   et_required_fields_mapping = et_full_fields_mapping ).

      WHEN if_mpa_output=>gc_mpa_scen-change.       "Mass Create Change

        " Get full fields with output order
        get_full_fields_mapping( EXPORTING
                                   is_struct             = ls_mass_change
                                 IMPORTING
                                   et_full_field_mapping = lt_full_fields_mapping ).

        " Find required fields for Mass Asset Change
        assemble_fields_mapping( EXPORTING
                                   is_params                  = ls_params
                                   is_struct                  = ls_mass_change
                                   it_full_fields_mapping     = lt_full_fields_mapping
                                   it_excepted_fields_mapping = lt_excepted_fields
                                 IMPORTING
                                   et_required_fields_mapping = et_full_fields_mapping ).

      WHEN if_mpa_output=>gc_mpa_scen-adjustment.  "Mass Create Adjustment

        " Get full fields with output order
        get_full_fields_mapping( EXPORTING
                                   is_struct             = ls_mass_adjustment
                                 IMPORTING
                                   et_full_field_mapping = lt_full_fields_mapping ).

        " Find required fields for Mass Asset Adjustment
        assemble_fields_mapping( EXPORTING
                                   is_params                  = ls_params
                                   is_struct                  = ls_mass_adjustment
                                   it_full_fields_mapping     = lt_full_fields_mapping
                                   it_excepted_fields_mapping = lt_excepted_fields
                                 IMPORTING
                                   et_required_fields_mapping = et_full_fields_mapping ).

      WHEN if_mpa_output=>gc_mpa_scen-retirement.  "Mass Create retirement

        " Get full fields with output order
        get_full_fields_mapping( EXPORTING
                                   is_struct             = ls_mass_retirement
                                 IMPORTING
                                   et_full_field_mapping = lt_full_fields_mapping ).

        " Find required fields for Mass Asset retirement
        assemble_fields_mapping( EXPORTING
                                   is_params                  = ls_params
                                   is_struct                  = ls_mass_retirement
                                   it_full_fields_mapping     = lt_full_fields_mapping
                                   it_excepted_fields_mapping = lt_excepted_fields
                                 IMPORTING
                                   et_required_fields_mapping = et_full_fields_mapping ).

    ENDCASE.

    " Sort both tables by position
    SORT et_full_fields_mapping BY position.

  ENDMETHOD.


  METHOD get_full_fields_mapping.

    DATA: ls_field_mapping  TYPE gty_s_field_name_mapping,
          ls_components     TYPE abap_compdescr,
          lo_strucdescr     TYPE REF TO cl_abap_structdescr,
          ls_asset_transfer TYPE mpa_s_asset_transfer.

    lo_strucdescr ?= cl_abap_typedescr=>describe_by_data( is_struct ).

    CASE gv_template_type.
      WHEN if_mpa_output=>gc_mpa_scen-transfer.
        " Set the title of the document
        gv_title  = if_mpa_output=>gc_mpa_temp-transfer.
        LOOP AT lo_strucdescr->components INTO ls_components.

          ls_field_mapping-position  = sy-tabix.
          ls_field_mapping-stru_name = ls_components-name.

          CASE ls_field_mapping-stru_name.
            WHEN 'SLNO' OR 'BLDAT' OR 'BUDAT' OR 'BZDAT' OR 'WWERT' OR 'BUKRS' OR 'ANLN1'  OR 'XANEU'.
              ls_field_mapping-mandatory = abap_true.
          ENDCASE.

          INSERT ls_field_mapping INTO TABLE et_full_field_mapping.
          CLEAR ls_field_mapping.

        ENDLOOP.

      WHEN if_mpa_output=>gc_mpa_scen-create.
        "Title of the template
        gv_title = if_mpa_output=>gc_mpa_temp-create.

        LOOP AT lo_strucdescr->components INTO ls_components.
          ls_field_mapping-position  = sy-tabix.
          ls_field_mapping-stru_name = ls_components-name.

          CASE ls_field_mapping-stru_name.
            WHEN 'SLNO' OR 'BUKRS' OR 'ANLKL' OR 'TXA50_ANLT' OR 'PRCTR'.
              ls_field_mapping-mandatory = abap_true.
          ENDCASE.

          INSERT ls_field_mapping INTO TABLE et_full_field_mapping.
          CLEAR ls_field_mapping.

        ENDLOOP.

      WHEN if_mpa_output=>gc_mpa_scen-change.
        "Title of the template
        gv_title = if_mpa_output=>gc_mpa_temp-change.

        LOOP AT lo_strucdescr->components INTO ls_components.
          ls_field_mapping-position  = sy-tabix.
          ls_field_mapping-stru_name = ls_components-name.

          CASE ls_field_mapping-stru_name.
            WHEN 'SLNO' OR 'BUKRS' OR 'ANLN1' OR 'ANLN2' .
              ls_field_mapping-mandatory = abap_true.
          ENDCASE.

          INSERT ls_field_mapping INTO TABLE et_full_field_mapping.
          CLEAR ls_field_mapping.
        ENDLOOP.

      WHEN if_mpa_output=>gc_mpa_scen-adjustment.
        "Title of the template
        gv_title = if_mpa_output=>gc_mpa_temp-adjustment.

        LOOP AT lo_strucdescr->components INTO ls_components.
          ls_field_mapping-position  = sy-tabix.
          ls_field_mapping-stru_name = ls_components-name.

          CASE ls_field_mapping-stru_name.
            WHEN 'SLNO' OR 'BLDAT' OR 'BUDAT' OR 'BZDAT' OR 'BUKRS' OR 'ANLN1' OR 'BWASL'.
              ls_field_mapping-mandatory = abap_true.
          ENDCASE.

          INSERT ls_field_mapping INTO TABLE et_full_field_mapping.
          CLEAR ls_field_mapping.
        ENDLOOP.

      WHEN if_mpa_output=>gc_mpa_scen-retirement.
        "Title of the template
        gv_title = if_mpa_output=>gc_mpa_temp-retirement.

        LOOP AT lo_strucdescr->components INTO ls_components.
          ls_field_mapping-position  = sy-tabix.
          ls_field_mapping-stru_name = ls_components-name.

          CASE ls_field_mapping-stru_name.
            WHEN 'SLNO' OR 'BLDAT' OR 'BUDAT' OR 'BZDAT' OR 'BUKRS' OR 'ANLN1' OR 'BWASL'.
              ls_field_mapping-mandatory = abap_true.
          ENDCASE.

          INSERT ls_field_mapping INTO TABLE et_full_field_mapping.
          CLEAR ls_field_mapping.
        ENDLOOP.

    ENDCASE.

  ENDMETHOD.


  METHOD generate_excel_template.


    DATA: ls_stream               TYPE /iwbep/if_mgw_core_srv_runtime=>ty_s_media_resource,
          lv_content              TYPE xstring,
          lo_tool_xls             TYPE REF TO cl_salv_export_tool_ats,
          lt_mpa                  TYPE gty_t_mpa_transfer,
          lr_mpa                  TYPE REF TO data,
          lt_blocks               TYPE if_salv_export_appendix=>yts_block,
          lr_appendix             TYPE REF TO if_salv_export_appendix=>yts_block,
          lv_template_doc         TYPE xstring,
          lo_table_row_descriptor TYPE REF TO cl_abap_structdescr,
          lo_source_table_descr   TYPE REF TO cl_abap_tabledescr.

    GET REFERENCE OF lt_mpa INTO lr_mpa.
    DATA(lo_itab_services) = cl_salv_itab_services=>create_for_table_ref( lr_mpa ).

    lo_source_table_descr   ?= cl_abap_tabledescr=>describe_by_data_ref( lr_mpa ).
    lo_table_row_descriptor ?= lo_source_table_descr->get_table_line_type( ).

    lo_tool_xls = cl_salv_export_tool_ats_xls=>create_for_excel_from_ats(
                                                 io_itab_services       = lo_itab_services
                                                 io_source_struct_descr = lo_table_row_descriptor
                                                 it_aggregation_rules   = VALUE if_salv_service_types=>yt_aggregation_rule( )
                                                 it_grouping_rules      = VALUE if_salv_service_types=>yt_grouping_rule( )
                                               ).

    DATA(lo_config) = lo_tool_xls->configuration( ).

    " Generate title of template
    render_title( CHANGING ct_blocks = lt_blocks ).

    " Generate header of table
    render_header( EXPORTING it_fields_header = it_field_mapping
                   CHANGING ct_blocks        = lt_blocks ).

    " Append the block for file download with data
*    IF is_block IS NOT INITIAL.
*      INSERT is_block INTO lt_blocks INDEX 3.
*    ENDIF.

    GET REFERENCE OF lt_blocks INTO lr_appendix.
    lo_config->if_salv_export_appendix~set_blocks( lr_appendix ).

    lo_tool_xls->read_result( IMPORTING content = lv_content  ).

*    add_worksheet( CHANGING cv_doc = lv_content ).

    DATA(lv_excel_xstring) = create_excel_with_sheets( it_fields_header = it_field_mapping
                                                       it_asset_data = it_asset_data ).

    " Format excel (hide structure fields, delete freeze line, add font)
    format_doc( EXPORTING
                  iv_source_doc        = lv_excel_xstring
                  it_fields_header     = it_field_mapping
*                  is_block             = is_block
                  iv_file_download_ind = iv_file_download_ind
                IMPORTING
                  ev_target_doc        = lv_template_doc ).

    " Set file content
    ls_stream-value     = lv_template_doc.
    ls_stream-mime_type = gc_mime_type-app_excel.

    copy_data_to_ref( EXPORTING
                        is_data = ls_stream
                      CHANGING
                        cr_data = er_stream ).

    " Set file name that get from request URL
    READ TABLE mt_param_tab INTO DATA(ls_filename) WITH KEY name = gc_param_name-filename.
    IF sy-subrc IS INITIAL.
      ev_filename = ls_filename-value.
    ELSE.
      ev_filename = gc_file_name-excel_file       ##NO_TEXT.
    ENDIF.

  ENDMETHOD.

  METHOD create_excel_with_sheets.

    DATA: lo_doc               TYPE REF TO zif_mpa_xlsx_doc,
          lt_sheet_info        TYPE zcl_mpa_xlsx=>gty_th_sheet_info,
          lo_sheet             TYPE REF TO zif_mpa_xlsx_sheet,
          lo_sheet_2           TYPE REF TO zif_mpa_xlsx_sheet,
          lo_sheet_3           TYPE REF TO zif_mpa_xlsx_sheet,
          lt_file              TYPE cpt_x255,
          lv_filename          TYPE localfile,
          lv_bytes_transferred TYPE i.

    DATA(lo_xlsx) = zcl_mpa_xlsx=>get_instance( ).
    lo_doc = lo_xlsx->create_doc( ).



    lt_sheet_info = lo_doc->get_sheets( ).
    lo_sheet = lo_doc->get_sheet_by_id( lt_sheet_info[ 1 ]-sheet_id ).
    lo_sheet->change_sheet_name( iv_new_name = 'Data' ). "not working lohid
    lo_sheet->set_cell_content( iv_row = 1 iv_column = 1 iv_value = gv_title ).

    DATA lt_comment TYPE string_table.
    " get the commented text for excel template
    get_comment_text( IMPORTING et_comment = lt_comment ).

    DATA lv_row_num TYPE i VALUE 2.
    LOOP AT lt_comment INTO DATA(lv_comment).
      CONCATENATE gc_comment_symbol lv_comment INTO DATA(lv_remark) SEPARATED BY space.

      " Create cell for comments
      lo_sheet->set_cell_content( iv_row = lv_row_num iv_column = 1 iv_value = lv_remark ).
      lv_row_num += 1.

    ENDLOOP.

    DATA lv_col_num TYPE i VALUE 1.
    " Create the technical field row for the template
    LOOP AT it_fields_header INTO DATA(ls_field_header).

      IF ls_field_header-stru_name IS NOT INITIAL.
        lo_sheet->set_cell_content( iv_row = lv_row_num iv_column = lv_col_num iv_value = ls_field_header-stru_name ).
      ENDIF.
      lv_col_num += 1.
    ENDLOOP.


    lv_col_num = 1.
    lv_row_num += 1.
    "Create the label row for the template
    LOOP AT it_fields_header INTO ls_field_header.

      IF ls_field_header-f_label IS NOT INITIAL.
        lo_sheet->set_cell_content( iv_row = lv_row_num iv_column = lv_col_num iv_value = ls_field_header-f_label ).
      ENDIF.
      lv_col_num += 1.
    ENDLOOP.

    IF it_asset_data IS NOT INITIAL.
      ASSIGN it_asset_data[ 1 ] TO FIELD-SYMBOL(<ls_asset_data>).
      DATA lv_comp_count TYPE i.


      LOOP AT <ls_asset_data>-mass_create_data ASSIGNING FIELD-SYMBOL(<ls_create_data>).
        lv_col_num = 1.
        CLEAR lv_comp_count.
        DO.
          ADD 1 TO lv_comp_count.
          ASSIGN COMPONENT lv_comp_count OF STRUCTURE <ls_create_data> TO FIELD-SYMBOL(<fs_comp>).
          IF sy-subrc NE 0.
            EXIT.
          ENDIF.
          lo_sheet->set_cell_content( iv_row = lv_row_num iv_column = lv_col_num iv_value = <fs_comp> ).
          lv_col_num += 1.
        ENDDO.

        lv_row_num += 1.

      ENDLOOP.

    ENDIF.


    lo_doc->add_new_sheet( iv_sheet_name = 'Field_List'  ).
*CATCH cx_openxml_format.
*CATCH cx_openxml_not_allowed.
*CATCH cx_dynamic_check.
    lo_sheet_2 = lo_doc->get_sheet_by_id( 2 ).

    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 1 iv_value = 'TagID_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 2 iv_value = 'AmountDate_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 3 iv_value = 'AmountTime_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 4 iv_value = 'AmountValue_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 5 iv_value = 'AmountUnit_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 6 iv_value = 'FaultInd_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 7 iv_value = 'CalibrationInd_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 8 iv_value = 'OutOfPrecisenessOperator_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 9 iv_value = 'NotAvailableInd_2' ).
    lo_sheet_2->set_cell_content( iv_row = 1 iv_column = 10 iv_value = 'Remark_2' ).

    lo_doc->add_new_sheet( iv_sheet_name = 'Introduction'  ).
*CATCH cx_openxml_format.
*CATCH cx_openxml_not_allowed.
*CATCH cx_dynamic_check.
    lo_sheet_3 = lo_doc->get_sheet_by_id( 3 ).

    "longtext
    DATA: lv_id          TYPE doku_id VALUE 'TX',
          lv_object      TYPE doku_obj VALUE 'ZMPA_DOCU_TEST',
          lv_langu       TYPE syst_langu,
          lt_line        TYPE TABLE OF tline,
          lt_new_comment TYPE string_table.

    lv_langu =  sy-langu.

    CALL FUNCTION 'DOCU_GET'
      EXPORTING
        id       = lv_id
        langu    = lv_langu
        object   = lv_object
      TABLES
        line     = lt_line
      EXCEPTIONS
        ret_code = 01
        OTHERS   = 99.

    CALL FUNCTION 'CONVERT_ITF_TO_STREAM_TEXT'
      EXPORTING
        lf           = 'X'
      IMPORTING
        stream_lines = lt_comment
      TABLES
        itf_text     = lt_line.

    lv_col_num = 1.
    lv_row_num = 1.
    "Create the label row for the template
    LOOP AT lt_comment ASSIGNING FIELD-SYMBOL(<ls_comment>).

      lo_sheet->set_cell_content( iv_row = lv_row_num iv_column = lv_col_num iv_value = <ls_comment> ).
      lv_col_num += 1.
      lv_row_num += 1 .
    ENDLOOP.

    rv_excel_file = lo_doc->save( ).

  ENDMETHOD.




  METHOD generate_csv_template.

    DATA: lv_batchid_str    TYPE string,
          lv_lineitem_str   TYPE string,
          lv_header_str     TYPE string,
          lv_header_str2    TYPE string,
          lv_template_str   TYPE string,
          ls_stream         TYPE /iwbep/if_mgw_core_srv_runtime=>ty_s_media_resource,
          lv_content        TYPE xstring,
          lv_header         TYPE string,
          lv_header2        TYPE string,
          lv_lineitem_title TYPE string,
          lt_comment        TYPE string_table,
          lv_remark_1       TYPE string,
          lv_remark_2       TYPE string,
          lv_remark_3       TYPE string.

    " Get Commented text
    get_comment_text( IMPORTING et_comment = lt_comment ).

    lv_remark_1 = |"{ cl_fac_xlsx_parse_utils=>gc_comment_symbol } { lt_comment[ 1 ] }"| .
    lv_remark_2 = |"{ cl_fac_xlsx_parse_utils=>gc_comment_symbol } { lt_comment[ 2 ] }"|.

    concatenate_field_line( EXPORTING it_fields_mapping = it_field_mapping
                                      iv_delimiter      = iv_delimiter
                            IMPORTING ev_line           = lv_header_str ).

    CONCATENATE gv_title
                lv_remark_1
                lv_remark_2
                lv_header_str
           INTO lv_template_str SEPARATED BY cl_abap_char_utilities=>cr_lf.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text     = lv_template_str
        mimetype = 'CSV'
      IMPORTING
        buffer   = lv_content.

    " Set file content
    ls_stream-value     = lv_content.
    ls_stream-mime_type = gc_mime_type-app_csv.

    copy_data_to_ref( EXPORTING
                        is_data = ls_stream
                      CHANGING
                        cr_data = er_stream ).

    " Set file name that get from request URL
    READ TABLE mt_param_tab INTO DATA(ls_filename) WITH KEY name = gc_param_name-filename.
    IF sy-subrc IS INITIAL.
      ev_filename = ls_filename-value.
    ELSE.
      ev_filename = gc_file_name-csv_file       ##NO_TEXT.
    ENDIF.

  ENDMETHOD.


  METHOD format_label.

    DATA: lv_leng           TYPE i,
          lv_leng_char      TYPE string,
          lv_formated_label TYPE string,
          ls_usr01          TYPE usr01,
          lv_curr_format    TYPE c LENGTH 15.

    CONSTANTS: lc_date_format TYPE c LENGTH 10 VALUE 'YYYY-MM-DD',
               lc_curr_format TYPE c LENGTH 15 VALUE '1.234.567,89'.

*    select single * from usr01 into ls_usr01 where bname eq sy-uname.
*    if sy-subrc is initial.
*      select single ddtext from dd07t into lv_date_format where domname = 'XUDATFM' and ddlanguage = sy-langu and as4local = 'A' and as4vers  = '0000' and domvalue_l = ls_usr01-datfm.
*      if sy-subrc is not initial.
*        lv_date_format = lc_date_format.
*      endif.
*      select single ddtext from dd07t into lv_curr_format where domname = 'XUDCPFM' and ddlanguage = sy-langu and as4local = 'A' and as4vers  = '0000' and domvalue_l = ls_usr01-dcpfm.
*      if sy-subrc is not initial.
*        lv_curr_format = lc_curr_format.
*      endif.
*    endif.

    LOOP AT ct_fields_mapping ASSIGNING FIELD-SYMBOL(<ls_field_mapping>).
      IF <ls_field_mapping>-data_type IS NOT INITIAL AND <ls_field_mapping>-length IS NOT INITIAL.

        " Type is date or currency, don't display length
        IF <ls_field_mapping>-data_type NE 'DATS' AND <ls_field_mapping>-data_type NE 'CURR'.
          lv_leng = <ls_field_mapping>-length.
          lv_leng_char = lv_leng.
          CONDENSE lv_leng_char NO-GAPS.
          CONCATENATE <ls_field_mapping>-label ' (' lv_leng_char ')' INTO  lv_formated_label.
        ELSE.
          lv_formated_label = <ls_field_mapping>-label.
        ENDIF.
        " Mandatory mark
        IF <ls_field_mapping>-mandatory EQ abap_true.
          CONCATENATE '*' lv_formated_label INTO  lv_formated_label.
        ENDIF.

        " To add descrption for date fields
        IF <ls_field_mapping>-data_type EQ 'DATS'.
          CONCATENATE <ls_field_mapping>-label ' (' lc_date_format ')' INTO  lv_formated_label.
        ENDIF.
*
*        " To add descrption for currency fields
*        if <ls_field_mapping>-data_type eq 'CURR'.
*          concatenate <ls_field_mapping>-label ' (' lv_curr_format ')' into  lv_formated_label.
*        endif.

      ENDIF.

      <ls_field_mapping>-f_label = lv_formated_label.

*      "longtext
*      DATA: lv_id      TYPE doku_id VALUE 'TX',
*            lv_object  TYPE doku_obj VALUE 'ZMPA_DOCU_TEST',
*            lv_langu   TYPE syst_langu,
*            lt_line    TYPE TABLE OF tline,
*            lt_comment TYPE string_table.
*
*      lv_langu =  sy-langu.
*
*      CALL FUNCTION 'DOCU_GET'
*        EXPORTING
*          id       = lv_id
*          langu    = lv_langu
*          object   = lv_object
*        TABLES
*          line     = lt_line
*        EXCEPTIONS
*          ret_code = 01
*          OTHERS   = 99.
*
*      CALL FUNCTION 'CONVERT_ITF_TO_STREAM_TEXT'
*        EXPORTING
*          lf           = 'X'
*        IMPORTING
*          stream_lines = lt_comment
*        TABLES
*          itf_text     = lt_line.
*
*      <ls_field_mapping>-f_label = |{ lv_formated_label } \n { lt_comment[ 1 ] } |.

    ENDLOOP.

  ENDMETHOD.


  METHOD format_doc.

    TYPES:
      BEGIN OF lty_st_col_type_index,
        type   TYPE char4,
        column TYPE i,
      END OF lty_st_col_type_index .

    DATA: lo_xlsx_doc           TYPE REF TO cl_xlsx_document,
          lo_workbookpart       TYPE REF TO cl_xlsx_workbookpart,
          lo_wordsheetparts     TYPE REF TO cl_openxml_partcollection,
          lo_wordsheetpart      TYPE REF TO cl_openxml_part,
          lo_sheet_content      TYPE xstring,
          lo_xml_document       TYPE REF TO cl_xml_document,
          lo_node               TYPE REF TO if_ixml_node,
          lo_node_attr          TYPE REF TO if_ixml_node,
          lo_node_rows          TYPE REF TO if_ixml_node_list,
          lo_attrs_map          TYPE REF TO if_ixml_named_node_map,
          lo_uri                TYPE REF TO cl_openxml_parturi,
          lo_formarted          TYPE xstring,
          lo_doc_parts          TYPE REF TO cl_openxml_partcollection,
          lo_stylepart          TYPE REF TO cl_openxml_part,
          lo_node_first_font    TYPE REF TO if_ixml_node,
          lo_node_last_font     TYPE REF TO if_ixml_node,
          lo_node_last_fill     TYPE REF TO if_ixml_node,
          lo_node_first_element TYPE REF TO if_ixml_node,
          lo_node_last_element  TYPE REF TO if_ixml_node,
          lo_node_first_style   TYPE REF TO if_ixml_node,
          lo_node_last_style    TYPE REF TO if_ixml_node,
          lo_node_first_col     TYPE REF TO if_ixml_node,
          lo_node_col           TYPE REF TO if_ixml_node,

          lv_col_index          TYPE i VALUE 1,          " column index starts from 1
          lt_text_col_index     TYPE TABLE OF i,
          lt_date_col_index     TYPE TABLE OF i,
          lt_column_type_index  TYPE TABLE OF lty_st_col_type_index,
          lo_node_col_per_row   TYPE REF TO if_ixml_node_list,

          lt_char_col           TYPE gty_tt_col_name,
          lt_date_col           TYPE gty_tt_col_name,
          lo_attribute          TYPE REF TO if_ixml_attribute,
          lo_col_node           TYPE REF TO if_ixml_node.

    CONSTANTS : lc_data_type_char TYPE char4 VALUE 'CHAR',
                lc_data_type_dats TYPE char4 VALUE 'DATS',
                lc_data_type_numc TYPE char4 VALUE 'NUMC',
                lc_data_type_curr TYPE char4 VALUE 'CURR'.

    FIELD-SYMBOLS <fs_val> TYPE string.

    DATA(lv_style_text) = COND string( WHEN iv_file_download_ind = abap_true
                                       THEN '8'
                                       ELSE '7' ).
    DATA(lv_style_date) = COND string( WHEN iv_file_download_ind = abap_true
                                       THEN '9'
                                       ELSE '8' ).

    lo_xlsx_doc       = cl_xlsx_document=>load_document( iv_source_doc ).
    lo_workbookpart   = lo_xlsx_doc->get_workbookpart( ).
    "sheet
**********************************************************************


*    DATA: lo_worksheet_part TYPE REF TO cl_xlsx_worksheetpart.
*    "Create a new Worksheet part
*    lo_worksheet_part = lo_workbookpart->add_worksheetpart( ).
*
*    "Create Sheet info for the new Sheet
*    DATA(ls_sheet_info) = zcreate_info_for_new_sheet( iv_sheet_name     = 'Sheet_2'
*                                                      io_worksheet_part = lo_worksheet_part
*                                                      io_workbookpart = lo_workbookpart ).
*
*
**   update the Workbook XML
**    zupdate_wb_xml_after_add_sheet( io_workbookpart = lo_workbookpart
**                                    is_sheet_info = ls_sheet_info ).
*
*    DATA lv_workbook_xml  TYPE xstring.
*    lv_workbook_xml =  lo_workbookpart->get_data( ).
*
*    CALL TRANSFORMATION xl_mpa_insert_sheet
*             PARAMETERS active_sheet = 2
*                        sheet_name   = ls_sheet_info-name
*                        sheet_id     = ls_sheet_info-sheet_id
*                        sheet_rid    = ls_sheet_info-rid
*             SOURCE XML lv_workbook_xml
*             RESULT XML lv_workbook_xml.



*    lo_xlsx_doc       = cl_xlsx_document=>load_document( iv_source_doc ).
*    lo_workbookpart   = lo_xlsx_doc->get_workbookpart( ).

*    data(lo_new_xlsx_doc) = cl_xlsx_document=>load_document( iv_source_doc ).
*    data(lo_new_workbookpart)   = lo_new_xlsx_doc->get_workbookpart( ).
*
*    data(lo_new_wordsheetpart) = lo_new_xlsx_doc->get_part_by_uri( ir_parturi = cl_openxml_parturi=>create_from_filename( iv_filename = '/xl/worksheets/sheet1.xml' ) ).

*
*    TRY.
*        DATA(lo_new_worksheet) = lo_workbookpart->add_worksheetpart( ).
*      CATCH cx_openxml_not_allowed.
*        "handle exception
*    ENDTRY.


*    CATCH cx_openxml_not_allowed.
*    DATA(lv_filename) = lo_new_worksheet->get_uri( ).
*    try.
*        lo_new_worksheet->add_part( ir_part = lo_new_wordsheetpart ).
*      catch cx_openxml_not_allowed.
*        "handle exception
*    endtry.
*    CATCH cx_openxml_not_allowed.
*lo_new_worksheet->feed_data( iv_data = iv_source_doc ).
************************************************************************
    lo_wordsheetparts = lo_workbookpart->get_worksheetparts( ).

**********************************************************************
    DATA(count) = lo_wordsheetparts->get_count( ).
**********************************************************************

    lo_uri           = cl_openxml_parturi=>create_from_filename( iv_filename = '/xl/worksheets/sheet1.xml' ).
    lo_wordsheetpart = lo_xlsx_doc->get_part_by_uri( ir_parturi = lo_uri ). "lo_wordsheetparts->get_part( 0 ).
    lo_sheet_content = lo_wordsheetpart->get_data( ).



    CREATE OBJECT lo_xml_document.
    lo_xml_document->parse_xstring( lo_sheet_content ).
*
*    " remove frozen setting
*    lo_node = lo_xml_document->find_node( name = 'selection' ).
*    lo_attrs_map = lo_node->get_attributes( ).
*    lo_attrs_map->remove_named_item( name = 'pane' ).
*    lo_node = lo_node->get_prev( ).
*    IF lo_node IS NOT INITIAL.
*      lo_node->remove_node( ).
*    ENDIF.

    " adjustment column width: fix the first cell width
*    lo_node           = lo_xml_document->find_node( name = 'cols' ).
*    lo_node_first_col = lo_node->get_first_child( ).
*    lo_attrs_map      = lo_node_first_col->get_attributes( ).
*    lo_attrs_map->get_named_item_ns( name = 'width' )->set_value( '20' ).  "length of column A
*
*    IF it_fields_header IS NOT INITIAL.
*
*      IF it_fields_header[ 1 ]-data_type = lc_data_type_char.
*        "adjust column style for the first column
*        lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'width' )->clone( ).
*        lo_node_attr->set_name( 'style' ).
*        lo_node_attr->set_value( lv_style_text ).
*        lo_attrs_map->set_named_item_ns( node = lo_node_attr ).
*      ENDIF.
*
*      lv_col_index = 2.
*
*      "adjust column style based on the type of data (from column 2)
*      lo_node_first_col = lo_node_first_col->get_next( ).
*      WHILE lo_node_first_col IS BOUND.
**        lo_attrs_map->get_named_item_ns( name = 'width' )->set_value( '20' ).  "length of column "longtext
*
*        IF it_fields_header[ lv_col_index ]-data_type = lc_data_type_char OR it_fields_header[ lv_col_index ]-data_type = lc_data_type_dats.
*
*          APPEND VALUE #( type = lc_data_type_char column = lv_col_index ) TO lt_column_type_index.
*
*          lo_attrs_map = lo_node_first_col->get_attributes( ).
*          lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'width' )->clone( ).
*          lo_node_attr->set_name( 'style' ).
*          lo_node_attr->set_value( lv_style_text ).
*          lo_attrs_map->set_named_item_ns( node = lo_node_attr ).
*
**        ELSEIF it_fields_header[ lv_col_index ]-data_type = lc_data_type_dats.
**
**          APPEND VALUE #( type = lc_data_type_dats column = lv_col_index )  TO lt_column_type_index.
**
**          lo_attrs_map = lo_node_first_col->get_attributes( ).
**          lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'width' )->clone( ).
**          lo_node_attr->set_name( 'style' ).
**          lo_node_attr->set_value( lv_style_date ).
**          lo_attrs_map->set_named_item_ns( node = lo_node_attr ).
*
*        ELSE.
*          APPEND VALUE #( type = lc_data_type_curr column = lv_col_index )  TO lt_column_type_index.
*        ENDIF.
*
*        lo_node_first_col = lo_node_first_col->get_next( ).
*        lv_col_index += 1.
*
*      ENDWHILE.
*    ENDIF.

    " row index start from 0, if input nothing, the row will be ignored from index
    lo_node      = lo_xml_document->find_node( name = 'sheetData' ).
    lo_node_rows = lo_node->get_children( ).

    " set height of row
    " adjust the height of the first header
    lo_node      = lo_node_rows->get_item( 0 ).
    lo_attrs_map = lo_node->get_attributes( ).

    " reference node 'r'
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'r' )->clone( ).
    lo_node_attr->set_name( 'ht' ).
    lo_node_attr->set_value( '25.5' )." adjust the height of 1st row
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).

    " reference node 'r'
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'r' )->clone( ).
    lo_node_attr->set_name( 'customHeight' ).
    lo_node_attr->set_value( '1' ).
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).

    " reference node 'r'
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'r' )->clone( ).
    lo_node_attr->set_name( 'customFormat' ).
    lo_node_attr->set_value( '1' ).
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).

    " reference node 'r'
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'r' )->clone( ).
    lo_node_attr->set_name( 's' ).
    lo_node_attr->set_value( lv_style_text ).
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).

    "*============================ Format excel with Data =============================**

*    "set style and type for rows with data, for downloaded file
*    DATA(lo_node_iterator) = lo_node_rows->create_iterator( ).
*    "get the row with data - from row 6 in excel, while downloading the file
*
*    DO 5 TIMES.
*      lo_node = lo_node_iterator->get_next( ).
*    ENDDO.
*
*    lo_node_col_per_row = lo_node->get_children( ).
*
*    LOOP AT lt_column_type_index REFERENCE INTO DATA(lr_column_ind).
*
*      lo_col_node = lo_node_col_per_row->get_item( index = lr_column_ind->column - 1 ).
*      CHECK lo_col_node IS NOT INITIAL.
*
*      lo_attrs_map = lo_col_node->get_attributes( ).
*      lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'r' ).
*
*      lo_attribute ?= lo_node_attr->query_interface( ixml_iid_attribute ).
*      DATA(lv_cell_index) = lo_attribute->get_value( ).
*
*      DATA(lv_length) = strlen( lv_cell_index ) - 1 .
*
*      IF lr_column_ind->type = lc_data_type_char OR lr_column_ind->type = lc_data_type_dats.
*        APPEND VALUE #( name = lv_cell_index(lv_length) ) TO lt_char_col.
**      ELSEIF lr_column_ind->type = lc_data_type_dats.
**        APPEND VALUE #( name = lv_cell_index(lv_length) ) TO lt_date_col.
*      ENDIF.
*
*    ENDLOOP.
*
*    DATA lv_row_num TYPE i VALUE 0.
*
*    "Process from row 6: data is filled from row 6 onwards
*    lo_node = lo_node_iterator->get_next( ).
*
*    WHILE lo_node IS NOT INITIAL.
*
*      "1st time would be row 6 then incrimented with one each time ...
*      "column is alphabet b
*      lo_node_col_per_row = lo_node->get_children( ).
*      DATA(lo_node_itr)   = lo_node_col_per_row->create_iterator( ).
*      lo_col_node         = lo_node_itr->get_next( ).
*
*      WHILE lo_col_node IS NOT INITIAL.
*
*        format_cell_for_download_file( is_block = is_block
*                                       it_char_col = lt_char_col
*                                       it_date_col = lt_date_col
*                                       io_col_node = lo_col_node
*                                       iv_style_text = lv_style_text
*                                       iv_style_date = lv_style_date
*                                       iv_row_num = lv_row_num ).
*
*        lo_col_node = lo_node_itr->get_next( ).
*
*      ENDWHILE.
*      lo_node = lo_node_iterator->get_next( ).
*      lv_row_num += 1.
*    ENDWHILE.

    "*======================= End of Format excel with Data ===============**



    lo_xml_document->render_2_xstring( IMPORTING stream = lo_formarted ).
    lo_wordsheetpart->feed_data( lo_formarted ).

***********************************************************************
*    lo_uri           = cl_openxml_parturi=>create_from_filename( iv_filename = '/xl/worksheets/sheet1.xml' ).
*    lo_wordsheetpart = lo_xlsx_doc->get_part_by_uri( ir_parturi = lo_uri ). "lo_wordsheetparts->get_part( 0 ).
*    lo_sheet_content = lo_wordsheetpart->get_data( ).


*    lo_new_worksheet->feed_data( iv_data = lo_sheet_content ).

**********************************************************************
*    DATA: lo_worksheet_part TYPE REF TO cl_xlsx_worksheetpart.
*    "Create a new Worksheet part
*    lo_worksheet_part = lo_workbookpart->add_worksheetpart( ).
*
*    "Create Sheet info for the new Sheet
*    DATA(ls_sheet_info) = zcreate_info_for_new_sheet( iv_sheet_name     = 'Sheet_2'
*                                                      io_worksheet_part = lo_worksheet_part
*                                                      io_workbookpart = lo_workbookpart ).
*
*
**   update the Workbook XML
**    zupdate_wb_xml_after_add_sheet( io_workbookpart = lo_workbookpart
**                                    is_sheet_info = ls_sheet_info ).
*
*    DATA lv_workbook_xml  TYPE xstring.
*    lv_workbook_xml =  lo_workbookpart->get_data( ).
*
*    CALL TRANSFORMATION xl_mpa_insert_sheet
*             PARAMETERS active_sheet = 1
*                        sheet_name   = ls_sheet_info-name
*                        sheet_id     = ls_sheet_info-sheet_id
*                        sheet_rid    = ls_sheet_info-rid
*             SOURCE XML lv_workbook_xml
*             RESULT XML lv_workbook_xml.
*
*
*    lo_uri           = cl_openxml_parturi=>create_from_filename( iv_filename = '/xl/worksheets/sheet2.xml' ).
*    lo_wordsheetpart = lo_xlsx_doc->get_part_by_uri( ir_parturi = lo_uri ). "lo_wordsheetparts->get_part( 0 ).
*    lo_sheet_content = lo_wordsheetpart->get_data( ).
*
*
*
*    lo_xml_document->render_2_xstring( IMPORTING stream = lo_formarted ).
*    lo_wordsheetpart->feed_data( lo_formarted ).


**********************************************************************
    "*------------------------------ style.xml -----------------------------*"
    " 1. Modify or add font styles
    " 2. Add cell style for text fields
    " 3. Add cell style for date fields
    " 4. Format other remaining cells
    "*----------------------------------------------------------------------*"
*
*    lo_stylepart = set_stylesxml( io_xlsx_doc = lo_xlsx_doc
*                                  io_xml_document = lo_xml_document ).


** Temp code
*
*    lo_node      = lo_xml_document->find_node( name = 'cellStyleXfs').
*    lo_attrs_map = lo_node->get_attributes( ).
*    lo_attrs_map->get_named_item_ns( name = 'count' )->set_value( '1' ).
*
**    lo_node_first_style = lo_node->get_first_child( )->clone( ).
**    lo_attrs_map        = lo_node_first_style->get_attributes( ).
**    lo_node->remove_child( old_child = lo_node_first_style ).
**    lo_attrs_map->get_named_item_ns( name = 'fontId' )->set_value( '0' ).
**    lo_node->append_child( lo_node_first_style ).
*
*    lo_node_first_style = lo_node->get_last_child( )->clone( ).
*    lo_node->remove_child( old_child = lo_node_first_style ).

*    lo_node      = lo_xml_document->find_node( name = 'cellStyles').
*    lo_attrs_map = lo_node->get_attributes( ).
*    lo_attrs_map->get_named_item_ns( name = 'count' )->set_value( '1' ).
*
*    lo_node_first_style = lo_node->get_first_child( )->clone( ).
*    lo_node_first_style->remove_node( ).

    "*----------------------------- End of style.xml ----------------------------*
    "*---------------------------------------------------------------------------*
    "sheet
**********************************************************************
    "get workbook
*  lo_doc_parts = lo_xlsx_doc->get_parts( ).
*    lo_uri       = cl_openxml_parturi=>create_from_filename( iv_filename = 'xl/workbook.xml' ).
*    lo_stylepart = lo_xlsx_doc->get_part_by_uri( ir_parturi = lo_uri ).
*    lo_xml_document->parse_xstring( lo_stylepart->get_data( ) ).
*
*    lo_node      = lo_xml_document->find_node( name = 'sheets').
*
*    lo_node_first_font    = lo_node->get_first_child( )->clone( )."sheet
**    lo_node_first_element = lo_node_first_font->get_first_child( ).
*    lo_attrs_map          = lo_node_first_font->get_attributes( ).
**     lo_attrs_map->get_named_item_ns( name = 'sheet' ).
*    lo_attrs_map->get_named_item_ns( name = 'id' )->set_value( 'rId4' ).
*    lo_attrs_map->get_named_item_ns( name = 'sheetId' )->set_value( '2' ).
*    lo_attrs_map->get_named_item_ns( name = 'name' )->set_value( 'Intro' ).
*    lo_node->append_child( lo_node_first_font ).

**********************************************************************



*    lo_xml_document->render_2_xstring( IMPORTING stream = lo_formarted ).
*    lo_stylepart->feed_data( lo_formarted ).

    ev_target_doc = lo_xlsx_doc->get_package_data( ).

  ENDMETHOD.

  METHOD set_stylesxml.

    DATA lo_node TYPE REF TO if_ixml_node.
    DATA lo_node_attr TYPE REF TO if_ixml_node.
    DATA lo_attrs_map TYPE REF TO if_ixml_named_node_map.
    DATA lo_uri TYPE REF TO cl_openxml_parturi.
    DATA lo_doc_parts TYPE REF TO cl_openxml_partcollection.
    DATA lo_node_first_font TYPE REF TO if_ixml_node.
    DATA lo_node_first_element TYPE REF TO if_ixml_node.
    DATA lo_node_first_style TYPE REF TO if_ixml_node.

    "Get the reference of styles.xml
    lo_doc_parts = io_xlsx_doc->get_parts( ).
    lo_uri       = cl_openxml_parturi=>create_from_filename( iv_filename = 'xl/styles.xml' ).
    ro_stylepart = io_xlsx_doc->get_part_by_uri( ir_parturi = lo_uri ).
    io_xml_document->parse_xstring( ro_stylepart->get_data( ) ).

    "Add font style, Modify the count of fonts
    lo_node      = io_xml_document->find_node( name = 'fonts').
    lo_attrs_map = lo_node->get_attributes( ).
    lo_attrs_map->get_named_item_ns( name = 'count' )->set_value( '6' ).

    "Set the styles for values
    lo_node_first_font    = lo_node->get_first_child( )->clone( ).
    lo_node_first_element = lo_node_first_font->get_first_child( ).
    lo_attrs_map          = lo_node_first_element->get_attributes( ).
    lo_attrs_map->get_named_item_ns( name = 'val' )->set_value( '11' ).
    lo_node->append_child( lo_node_first_font ).

    "Add cell style
    lo_node      = io_xml_document->find_node( name = 'cellXfs').
    lo_attrs_map = lo_node->get_attributes( ).
    lo_attrs_map->get_named_item_ns( name = 'count' )->set_value( '8' ).

    "Add font reference
    lo_node_first_style = lo_node->get_first_child( )->clone( ).
    lo_attrs_map        = lo_node_first_style->get_attributes( ).
    lo_attrs_map->get_named_item_ns( name = 'fontId' )->set_value( '5' ).
    lo_node->append_child( lo_node_first_style ).

    "Add fill style reference
    lo_node_first_style = lo_node->get_first_child( )->clone( ).
    lo_node_first_style->get_first_child( )->remove_node( ).
    lo_attrs_map        = lo_node_first_style->get_attributes( ).

    "Reference node 'xfId'
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'xfId' )->clone( ).
    lo_node_attr->set_name( 'applyFill' ).
    lo_node_attr->set_value( '0' ).
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).
    lo_attrs_map->get_named_item_ns( name = 'fillId' )->set_value( '0' ). "5
    lo_attrs_map->get_named_item_ns( name = 'applyNumberFormat' )->set_value( '1' ).
    lo_attrs_map->get_named_item_ns( name = 'applyAlignment' )->set_value( '1' ).
    lo_node->append_child( lo_node_first_style ).

    " Add font reference(Same code)
    lo_node_first_style = lo_node->get_first_child( )->clone( ).
    lo_attrs_map = lo_node_first_style->get_attributes( ).
    lo_attrs_map->get_named_item_ns( name = 'fontId' )->set_value( '5' ).
    lo_node->append_child( lo_node_first_style ).

    "Reference node 'xfId' for text field styling
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'xfId' )->clone( ).
    lo_node_attr->set_name( 'applyFill' ).
    lo_node_attr->set_value( '0' ).
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).
    lo_attrs_map->get_named_item_ns( name = 'fillId' )->set_value( '0' ). "2
    lo_attrs_map->get_named_item_ns( name = 'applyNumberFormat' )->set_value( '1' ).
    lo_attrs_map->get_named_item_ns( name = 'applyAlignment' )->set_value( '1' ).
    lo_attrs_map->get_named_item_ns( name = 'numFmtId' )->set_value( '49' ). "text
    lo_node->append_child( lo_node_first_style ).

    " add font reference(Same code)
    lo_node_first_style = lo_node->get_first_child( )->clone( ).
    lo_attrs_map        = lo_node_first_style->get_attributes( ).
    lo_attrs_map->get_named_item_ns( name = 'fontId' )->set_value( '5' ).
    lo_node->append_child( lo_node_first_style ).

    "Reference node 'xfId' for date field styling
    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'xfId' )->clone( ).
    lo_node_attr->set_name( 'applyFill' ).
    lo_node_attr->set_value( '0' ).
    lo_attrs_map->set_named_item_ns( node = lo_node_attr ).
    lo_attrs_map->get_named_item_ns( name = 'fillId' )->set_value( '0' ).
    lo_attrs_map->get_named_item_ns( name = 'applyNumberFormat' )->set_value( '1' ).
    lo_attrs_map->get_named_item_ns( name = 'applyAlignment' )->set_value( '1' ).
    lo_attrs_map->get_named_item_ns( name = 'numFmtId' )->set_value( '14' )."date
    lo_node->append_child( lo_node_first_style ).

  ENDMETHOD.




  METHOD format_cell_for_download_file.

    TYPES name TYPE string.
    DATA lo_node TYPE REF TO if_ixml_node.
    DATA lo_node_attr TYPE REF TO if_ixml_node.
    DATA lo_attrs_map TYPE REF TO if_ixml_named_node_map.
    DATA lo_attribute TYPE REF TO if_ixml_attribute.
    DATA lv_cell_index TYPE string.


*        data(lv_value) = lo_col_node->get_value( ).
    lo_attrs_map   = io_col_node->get_attributes( ).


    lo_node_attr = lo_attrs_map->get_named_item_ns( name = 'r' ).
    DATA(lv_val) = lo_node_attr->get_value( ).

    lo_attribute ?= lo_node_attr->query_interface( ixml_iid_attribute ).
    lv_cell_index = lo_attribute->get_value( ).


    DATA(lv_col_str_length) = strlen( lv_cell_index ) - strlen( CONV string( iv_row_num + 6 ) ) + 1. "get row name and column from cell

    DATA(lv_column_name) = lv_cell_index(lv_col_str_length).
    FIND lv_column_name IN sy-abcde MATCH OFFSET DATA(lv_com_num).


    IF line_exists( it_char_col[ name = lv_cell_index(lv_col_str_length) ] ) .

      lo_node_attr = lo_attrs_map->get_named_item_ns( name = 's' ).
      lo_node_attr->set_value( iv_style_text ).


*    ELSEIF line_exists( it_date_col[ name = lv_cell_index(lv_col_str_length) ] ) .
*
*      lo_node_attr = lo_attrs_map->get_named_item_ns( name = 's' ).
*      lo_node_attr->set_value( iv_style_date ).
*
*
*      lo_attrs_map->remove_named_item_ns( name = 't' ).
*
*      DATA(lo_child) = io_col_node->get_children( ).
*      lo_node      = lo_child->get_item( 0 ).
*
*      DATA(lv_date_from_db) = is_block-cells[ column_index = lv_com_num + 1 row_index = iv_row_num  ]-value.
*
*      IF strlen( lv_date_from_db ) > 8 .
*
*        DATA : lv_sap_date       TYPE d,
*               lv_sap_start_date TYPE d VALUE '18991231',
*               lv_excel_date     TYPE i.
*
*        CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
*          EXPORTING
*            date_external = lv_date_from_db
*          IMPORTING
*            date_internal = lv_sap_date.
*
*        lv_excel_date = lv_sap_date - lv_sap_start_date + 1.
*
*        lo_node->set_value( value = CONV string( lv_excel_date ) ).
*
*      ELSE.
*        lo_node->set_value( value = is_block-cells[ column_index = lv_com_num + 1 row_index = iv_row_num  ]-value ).
*      ENDIF.
*
    ENDIF.

  ENDMETHOD.


  METHOD find_field_type.

    DATA: lo_strucdescr TYPE REF TO cl_abap_structdescr,
          lo_datadescr  TYPE REF TO cl_abap_datadescr.

    FIELD-SYMBOLS: <ls_field>   TYPE gty_s_field_name_mapping.

    lo_strucdescr ?= cl_abap_typedescr=>describe_by_data( is_stru ).
    LOOP AT ct_fields_mapping ASSIGNING <ls_field>.
      lo_datadescr = lo_strucdescr->get_component_type( p_name = <ls_field>-stru_name ).
      <ls_field>-stru_type = lo_datadescr->get_relative_name( ).
    ENDLOOP.
  ENDMETHOD.


  METHOD fillin_others.

    DATA lt_fields_mapping TYPE gty_t_field_name_mappings.
    FIELD-SYMBOLS: <fs_field_mapping> TYPE gty_s_field_name_mapping.

    " Fill in label by data element first
    SELECT leng AS length, datatype AS data_type, rollname AS stru_type FROM dd04l
      INTO CORRESPONDING FIELDS OF TABLE @lt_fields_mapping
      FOR ALL ENTRIES IN @ct_fields_mapping
       WHERE rollname = @ct_fields_mapping-stru_type AND as4local   = 'A'.

    " Merge length and data type into target internal table
    LOOP AT ct_fields_mapping ASSIGNING FIELD-SYMBOL(<ls_fields_mapping>).
      READ TABLE lt_fields_mapping INTO DATA(ls_fields_mapping_len) WITH KEY stru_type = <ls_fields_mapping>-stru_type.
      IF sy-subrc = 0.
        <ls_fields_mapping>-length = ls_fields_mapping_len-length.
        <ls_fields_mapping>-data_type = ls_fields_mapping_len-data_type.

      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD fillin_label.

    DATA lt_fields_mapping TYPE gty_t_field_name_mappings.
    FIELD-SYMBOLS <fs_field_mapping> TYPE gty_s_field_name_mapping.

    " Fill in label by data element first
    SELECT ddtext AS label, rollname AS stru_type FROM dd04t
      INTO CORRESPONDING FIELDS OF TABLE @lt_fields_mapping
       FOR ALL ENTRIES IN @ct_fields_mapping
       WHERE rollname = @ct_fields_mapping-stru_type AND ddlanguage = @iv_language AND as4local   = 'A'.

    " Merge label into target internal table
    LOOP AT ct_fields_mapping ASSIGNING FIELD-SYMBOL(<ls_fields_mapping>).
      READ TABLE lt_fields_mapping INTO DATA(ls_fields_mapping_label) WITH KEY stru_type = <ls_fields_mapping>-stru_type.
      IF sy-subrc = 0.
        <ls_fields_mapping>-label = ls_fields_mapping_label-label.
      ENDIF.
    ENDLOOP.

    " Label_type override stru_type label, if label_type is not initial
    LOOP AT ct_fields_mapping ASSIGNING <fs_field_mapping> WHERE label_type IS NOT INITIAL.
      SELECT SINGLE FROM dd04t                          "#EC CI_NOORDER
          FIELDS ddtext
         WHERE rollname = @<fs_field_mapping>-label_type AND ddlanguage = @iv_language AND as4local   = 'A'
        INTO @DATA(lv_default_label) ##WARN_OK.

      IF sy-subrc = 0.
        <fs_field_mapping>-label = lv_default_label.
      ENDIF.
    ENDLOOP.

    " If label is empty yet, english label as default label will be set
    LOOP AT ct_fields_mapping ASSIGNING <fs_field_mapping> WHERE label IS INITIAL.
      SELECT SINGLE FROM dd04t FIELDS ddtext            "#EC CI_NOORDER
       WHERE rollname = @<fs_field_mapping>-stru_type AND ddlanguage = 'E' AND as4local   = 'A'
        INTO @lv_default_label ##WARN_OK.

      IF sy-subrc = 0.
        <fs_field_mapping>-label = lv_default_label.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD copy_data_to_ref.

    FIELD-SYMBOLS: <ls_data> TYPE any.

    CREATE DATA cr_data LIKE is_data.
    ASSIGN cr_data->* TO <ls_data>.
    <ls_data> = is_data.

  ENDMETHOD.


  METHOD concatenate_field_line.

    DATA: lt_fields TYPE string_table,
          lv_line_f TYPE string,
          lv_line_s TYPE string.

    CLEAR lt_fields.
    LOOP AT it_fields_mapping INTO DATA(ls_fields_mapping).
      APPEND ls_fields_mapping-prefix && ls_fields_mapping-stru_name TO lt_fields.
    ENDLOOP.
    CONCATENATE LINES OF lt_fields INTO lv_line_s SEPARATED BY iv_delimiter.

    CLEAR lt_fields.
    LOOP AT it_fields_mapping INTO ls_fields_mapping.
      ls_fields_mapping-f_label = |"{ ls_fields_mapping-f_label }"|.
      APPEND ls_fields_mapping-f_label TO lt_fields.
    ENDLOOP.
    CONCATENATE LINES OF lt_fields INTO lv_line_f SEPARATED BY iv_delimiter.
    CONCATENATE lv_line_s lv_line_f INTO ev_line SEPARATED BY cl_abap_char_utilities=>cr_lf.   "iv_title

  ENDMETHOD.


  METHOD assemble_fields_mapping.

    DATA: ls_field_mapping    TYPE gty_s_field_name_mapping,
          lt_selected_element TYPE string_table,
          lt_key_text_element TYPE mpa_t_textpool,
          lt_properties       TYPE gty_t_struct_properties,
          ls_properties       TYPE gty_s_struct_properties,
          ls_mass_transfer    TYPE mpa_s_asset_transfer.

    CLEAR: lt_selected_element.

    get_struct_properties( EXPORTING
                             iv_struct_name       = is_struct
                           IMPORTING
                             et_struct_properties = lt_properties ).

    " Find required fields for header
    LOOP AT lt_properties INTO ls_properties.

      " Find the required field against excepted fields list
      IF it_excepted_fields_mapping IS SUPPLIED.    "Excepted Field Table
        IF NOT line_exists( it_excepted_fields_mapping[ table_line = ls_properties-field_name ] ).

          ls_field_mapping-stru_name = ls_properties-field_name.

          " If don't find field in full mapping table, will set position into 999(last)
          READ TABLE it_full_fields_mapping INTO DATA(ls_full_field) WITH KEY stru_name = ls_properties-field_name.
          IF sy-subrc IS INITIAL.
            ls_field_mapping-position   = ls_full_field-position.
            ls_field_mapping-mandatory  = ls_full_field-mandatory.
            ls_field_mapping-label_type = ls_full_field-label_type.
            ls_field_mapping-prefix     = ls_full_field-prefix.
          ENDIF.
          APPEND ls_field_mapping TO et_required_fields_mapping.
          CLEAR ls_field_mapping.

        ENDIF.
      ENDIF.
    ENDLOOP.

    " Find data element from structure name
    find_field_type( EXPORTING
                       is_stru           = is_struct
                     CHANGING
                       ct_fields_mapping = et_required_fields_mapping ).

    " Find label and fill into mapping
    fillin_label( EXPORTING
                    iv_language       = is_params-langu
                  CHANGING
                    ct_fields_mapping = et_required_fields_mapping ).

    " Find the max length of field and mandatory
    fillin_others( CHANGING
                     ct_fields_mapping = et_required_fields_mapping ).

    " Format label like *Transaction Currency (5)
    format_label( CHANGING
                    ct_fields_mapping = et_required_fields_mapping ).

  ENDMETHOD.


  METHOD constructor.
    mo_lcl_template_dl = NEW lcl_mpa_template_download_util( ).
    mo_function_module = NEW lcl_function_module( ).
  ENDMETHOD.


  METHOD get_struct_properties.

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

    SELECT leng AS length, datatype AS data_type, rollname AS data_element INTO TABLE @DATA(lt_data) FROM dd04l
       FOR ALL ENTRIES IN @lt_data_element
        WHERE rollname = @lt_data_element-data_element AND as4local   = 'A'.

    LOOP AT lt_data_element INTO ls_data_element.

      READ TABLE lt_data INTO DATA(ls_data) WITH KEY data_element = ls_data_element-data_element.
      IF sy-subrc IS INITIAL.
        INSERT VALUE #( length       = ls_data-length
                        data_type    = ls_data-data_type
                        field_name   = ls_data_element-field_name
                        data_element = ls_data_element-data_element )
             INTO TABLE et_struct_properties.

      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD get_instance.
    IF go_instance IS NOT BOUND.
      ro_instance = NEW zcl_template_download_util( ).
    ELSE.
      ro_instance = go_instance.
    ENDIF.
  ENDMETHOD.


  METHOD if_mpa_template_download_util~generate_template.

    DATA: lt_field_mapping TYPE gty_t_field_name_mappings,
          ls_mimetype      TYPE /iwbep/s_mgw_name_value_pair.

    TRY.

*        if io_tech_request_context is not initial and it_key_tab is not initial.

        "** Set parameters for generate template document
        mo_lcl_template_dl->set_parameter( EXPORTING
                                             io_tech_request_context = io_tech_request_context
                                             it_key_tab              = it_key_tab
                                           CHANGING
                                             ct_param_tab            = mt_param_tab ).

        READ TABLE mt_param_tab INTO ls_mimetype WITH KEY name = gc_param_name-mimetype ##NO_TEXT.
        IF sy-subrc IS INITIAL.
          TRANSLATE ls_mimetype-value TO UPPER CASE.
        ENDIF.
*        else.
*          ls_mimetype-value = 'XLSX'.
*        endif.

        " Get required field name and translated label
        get_trans_mapping( IMPORTING et_full_fields_mapping = lt_field_mapping ).

        IF lt_field_mapping IS NOT INITIAL.

          " Bases on mime type, call the appropriate method for generating template
          CASE ls_mimetype-value.
            WHEN gc_file_format-excel.
              generate_excel_template( EXPORTING
                                         it_field_mapping = lt_field_mapping
                                       IMPORTING
                                         er_stream        = er_stream
                                         ev_filename      = ev_filename ).

            WHEN gc_file_format-csvc.
              generate_csv_template( EXPORTING
                                       it_field_mapping = lt_field_mapping
                                     IMPORTING
                                       er_stream        = er_stream
                                       ev_filename      = ev_filename ).
            WHEN gc_file_format-csvs.
              generate_csv_template( EXPORTING
                                       it_field_mapping = lt_field_mapping
                                       iv_delimiter     = ';'
                                     IMPORTING
                                       er_stream        = er_stream
                                       ev_filename      = ev_filename ).

            WHEN OTHERS.  "TODO

          ENDCASE.
        ENDIF.

      CATCH cx_static_check INTO DATA(lo_exception).
        "assert 1 = 2. "TODO

    ENDTRY.

  ENDMETHOD.


  METHOD if_mpa_template_download_util~get_template_type.

    TYPES: BEGIN OF lty_s_template,
             valpos     TYPE valpos,
             domvalue_l TYPE c LENGTH 15,
             ddtext     TYPE c LENGTH 30,
           END OF lty_s_template .

    DATA: lt_template TYPE STANDARD TABLE OF lty_s_template,
          ls_entity   TYPE cl_mpa_asset_process_mpc=>ts_templatetype.

    SELECT valpos domvalue_l ddtext FROM dd07v INTO TABLE lt_template WHERE domname = gc_template AND ddlanguage = sy-langu.
    IF sy-subrc IS INITIAL.
      SORT lt_template BY valpos.
      LOOP AT lt_template ASSIGNING FIELD-SYMBOL(<fs>).
        MOVE <fs>-domvalue_l TO ls_entity-templateid.
        MOVE <fs>-ddtext     TO ls_entity-templatedesc.
        APPEND ls_entity TO et_entityset.
        CLEAR ls_entity.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.


  METHOD get_comment_text.

    DATA: ls_comment TYPE string.

    "  Add comments for template header
    ls_comment = TEXT-001.
    APPEND ls_comment TO et_comment.
    ls_comment = TEXT-002.
    APPEND ls_comment TO et_comment.

  ENDMETHOD.


  METHOD get_excepted_fields.

    CLEAR rt_excepted_fields.
    APPEND 'STATUS' TO rt_excepted_fields.


  ENDMETHOD.


  METHOD if_mpa_template_download_util~download_file.

    DATA: ls_uploaded_file TYPE mpa_asset_data,
          lv_mime_type     TYPE string,
          ls_param         TYPE /iwbep/s_mgw_name_value_pair,
          lo_mpa_pasrs     TYPE REF TO if_mpa_xlsx_parse_util.

    lo_mpa_pasrs = zcl_xlsx_parse_util=>get_instance( ).

    "-------Set parameters for generate template document------"
    mo_lcl_template_dl->set_parameter( EXPORTING
                                         io_tech_request_context = io_tech_request_context
                                         it_key_tab              = it_key_tab
                                       CHANGING
                                         ct_param_tab            = mt_param_tab ).

    CLEAR ls_param.
    READ TABLE mt_param_tab INTO ls_param WITH KEY name = gc_param_name-fileid.
    IF sy-subrc IS INITIAL.

      "-------Get the table data based on file id----------------"
      SELECT SINGLE * FROM mpa_asset_data INTO ls_uploaded_file WHERE file_id = ls_param-value. "#EC CI_NOORDER
      IF sy-subrc IS INITIAL.

        CLEAR: ls_param.
        ls_param-name  = gc_param_name-templateid.
        ls_param-value = ls_uploaded_file-scen_type.
        APPEND ls_param TO mt_param_tab.

        CLEAR: ls_param.
        ls_param-name  = gc_param_name-filename.
        ls_param-value = ls_uploaded_file-file_name.
        APPEND ls_param TO mt_param_tab.

        IF ls_uploaded_file-file_name CS gc_file_format-xlsx." OR ls_uploaded_file-file_name CS gc_file_format-xlsx.
          lv_mime_type = gc_file_format-excel.
        ELSEIF ls_uploaded_file-file_name CS gc_file_format-csv.
          lv_mime_type = gc_file_format-csv.
        ENDIF.

        TRY.


            CASE lv_mime_type.
              WHEN gc_file_format-excel.

                "Get the file data into table
                zcl_xlsx_parse_util=>get_instance( )->transform_xstring_2_tab(
                   EXPORTING
                     ix_file         = ls_uploaded_file-file_data
                     iv_download_file_ind = abap_true
                   IMPORTING
                     et_table        = DATA(lt_excel_rows) ).

                "Get the excel data index
                zcl_xlsx_parse_util=>get_instance( )->check_excel_layout(
                   IMPORTING
                    et_line_index = DATA(lt_line_index)
                    ev_slno_status_flag = DATA(lv_status_flag)
                   CHANGING
                     ct_table     = lt_excel_rows ).

                "Get the data into excel blocks
                generate_excel_data_block( EXPORTING
                                              it_excel_rows = lt_excel_rows
                                              it_line_index = lt_line_index
                                            IMPORTING
                                              es_block = DATA(ls_block) ).

                DATA(lt_asset_data) = get_asset_data( EXPORTING ix_file = ls_uploaded_file-file_data ).

                "Get required field name and translated label
                get_trans_mapping(
                  EXPORTING
                    iv_status_slno_flag = lv_status_flag
                    IMPORTING
                    et_full_fields_mapping   = DATA(lt_field_mapping) ).

                IF lt_field_mapping IS NOT INITIAL.

                  "Generate the excel template with data
                  generate_excel_template( EXPORTING
                                             it_field_mapping = lt_field_mapping
                                             it_asset_data         = lt_asset_data
                                             iv_file_download_ind = abap_true
                                           IMPORTING
                                             er_stream        = er_stream
                                             ev_filename      = ev_filename ).
                ENDIF.

              WHEN gc_file_format-csv .

                get_csv_file_with_data( EXPORTING io_mpa_pasrs = lo_mpa_pasrs
                                                  is_uploaded_file = ls_uploaded_file
                                        IMPORTING ev_filename = ev_filename
                                                  er_stream = er_stream ).

            ENDCASE.


          CATCH cx_mpa_exception_handler cx_openxml_not_found cx_openxml_format cx_salv_export_error /iwbep/cx_mgw_med_exception INTO DATA(lo_exception).   "##NO_HANDLER
            "Call similer method to handle the exception like handle_exception
            "or
            "insert value bapiret2( type   = if_mpa_output=>gc_msg_type-error
            "             id     = if_mpa_output=>gc_msgid-mpa
            "             number = if_mpa_output=>gc_msgno_mpa-chgd_tmpl ) into table et_message.

        ENDTRY.
      ENDIF.

    ELSE.
      "Invalid file
    ENDIF.

  ENDMETHOD.

  METHOD get_csv_file_with_data.

    DATA lv_status_flag TYPE abap_bool.

    io_mpa_pasrs->parse_csv(
      EXPORTING
        ix_file     = is_uploaded_file-file_data
      IMPORTING
        ev_mpa_type = DATA(lv_mpa_type)
        et_asset    = DATA(lt_asset)
        ev_seperator = DATA(lv_delimiter)
        ev_status_flag = lv_status_flag ).

    "Get required field name and translated label
    get_trans_mapping( EXPORTING iv_status_slno_flag = abap_true
                       IMPORTING et_full_fields_mapping   = DATA(lt_csv_field_mapping) ).



    generate_csv_download_file( EXPORTING it_field_mapping = lt_csv_field_mapping
                                          iv_delimiter = lv_delimiter
                                          is_data =  lt_asset[ 1 ]
                                          iv_mpa_type = lv_mpa_type
                                IMPORTING er_stream        = er_stream
                                          ev_filename      = ev_filename ).

  ENDMETHOD.




  METHOD generate_excel_data_block.

    TYPES: BEGIN OF lty_st_field_name_mapping,
             cell_name TYPE string,
             cell_posi TYPE string,
             stru_name TYPE string,
             value     TYPE string,
           END OF lty_st_field_name_mapping,

           lty_tt_field_name_mapping TYPE HASHED TABLE OF lty_st_field_name_mapping
                                WITH UNIQUE KEY cell_name cell_posi,

           BEGIN OF lty_st_value_pair,
             index TYPE  char100,
             value TYPE REF TO data,
           END   OF lty_st_value_pair.

    DATA : ls_field_mapping    TYPE lty_st_field_name_mapping,
           lt_field_mapping    TYPE lty_tt_field_name_mapping,
           lv_mpa_type         TYPE mpa_template_type,
           ls_cells            TYPE TABLE OF if_salv_export_appendix=>ys_cell,
           ls_cell             LIKE LINE OF es_block-cells,
           ls_value_formatting TYPE if_salv_export_appendix=>ys_cell_formatting,
           lv_column_index     TYPE i VALUE 1,
           lv_row_index        TYPE i VALUE 0,
           lv_date_internal    TYPE dats.


    FIELD-SYMBOLS : <ft_line>   TYPE mpa_t_index_value_pair,
                    <fs_cell>   TYPE lty_st_value_pair,
                    <fs_value>  TYPE string,
                    <dyn_value> TYPE any.

    "-----------Value format----------"
    ls_value_formatting = VALUE #( is_bold              = abap_false
                                   horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-center
                                   vertical_alignment   = if_salv_export_appendix=>cs_vertical_alignment-center
                                   foreground_color     = 'FF000000'
                                 ).

    "----------Excel data block--------"
    es_block = VALUE #( ordinal_number = 3
                        location       = if_salv_export_appendix=>cs_appendix_location-top
                        name           = gc_excel_blocks-template_data
                        cells          = ls_cells ).

    "-----Read the uploaded file data based on index--------"
    READ TABLE it_line_index INTO DATA(ls_index) INDEX 1.
    IF sy-subrc IS INITIAL.
      READ TABLE it_excel_rows ASSIGNING FIELD-SYMBOL(<fs_header_techn>) WITH TABLE KEY index = ls_index-header_techn.
      IF sy-subrc IS INITIAL.

        ASSIGN <fs_header_techn>-value->* TO <ft_line>.
        LOOP AT <ft_line> ASSIGNING <fs_cell>.
          ASSIGN <fs_cell>-value->* TO <fs_value>.

          INSERT VALUE #( cell_name = <fs_value>
                          cell_posi = <fs_cell>-index
                          stru_name = <fs_value> ) INTO TABLE lt_field_mapping.
        ENDLOOP.

        "*--------Parse data lines base on field-column mapping--------*
        LOOP AT ls_index-data INTO DATA(lv_header_index).

          READ TABLE it_excel_rows ASSIGNING FIELD-SYMBOL(<fs_header_data>) WITH TABLE KEY index = lv_header_index.
          IF sy-subrc IS INITIAL.
            ASSIGN <fs_header_data>-value->* TO <ft_line>.

            LOOP AT lt_field_mapping INTO DATA(ls_mapping).

              READ TABLE <ft_line> ASSIGNING <fs_cell> WITH KEY index = ls_mapping-cell_posi.
              IF sy-subrc IS INITIAL.
                ASSIGN <fs_cell>-value->* TO  <fs_value>.

                ls_value_formatting-horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-begin_of_line.

                "Create the data cells for excel sheet
                ls_cell = VALUE #( row_index    = lv_row_index
                                   column_index = lv_column_index
                                   row_span     = 1
                                   column_span  = 0
                                   content_type = if_salv_export_appendix=>cs_cell_content_type-text
                                   formatting   = ls_value_formatting
                                   value        = <fs_value> ).


                APPEND ls_cell TO es_block-cells.
                lv_column_index = lv_column_index + 1.       "Move to next field
              ELSE.
                lv_column_index = lv_column_index + 1.
              ENDIF.

            ENDLOOP.
            lv_row_index    = lv_row_index + 1.
            lv_column_index = 1.

          ENDIF.
        ENDLOOP.

      ENDIF.
    ENDIF.

  ENDMETHOD.


  METHOD if_mpa_template_download_util~save_result_file.

    DATA: ls_block            TYPE if_salv_export_appendix=>ys_block,
          lt_block            TYPE if_salv_export_appendix=>yts_block,
          ls_field_header     TYPE gty_s_field_name_mapping,
          lv_column_index     TYPE i VALUE 1,
          lv_row_index        TYPE i VALUE 1,
          lv_value_row_index  TYPE i,
          ls_cells            TYPE TABLE OF if_salv_export_appendix=>ys_cell,
          ls_cell             LIKE LINE OF ls_block-cells,
          ls_value_formatting TYPE if_salv_export_appendix=>ys_cell_formatting.

    DATA: ls_field_mapping  TYPE gty_s_field_name_mapping,
          lt_field_mapping  TYPE gty_t_field_name_mappings,
          ls_components     TYPE abap_compdescr,
          lo_strucdescr     TYPE REF TO cl_abap_structdescr,
          lo_datadescr      TYPE REF TO cl_abap_datadescr,
          ls_asset_transfer TYPE mpa_s_asset_transfer.

    CLEAR mt_param_tab.

    SPLIT is_mpa_asset_data-file_name AT '.' INTO DATA(lv1) DATA(lv2).
    mt_param_tab = VALUE #( ( name = 'Language'   value = 'EN' )
                            ( name = 'Mimetype'   value = lv2 )
                            ( name = 'TemplateId' value = is_mpa_asset_data-scen_type ) ).

    TRY.
        get_trans_mapping(
          EXPORTING
            iv_status_slno_flag    = abap_true
          IMPORTING
            et_full_fields_mapping = lt_field_mapping ).
      CATCH /iwbep/cx_mgw_med_exception.
    ENDTRY.


    IF lt_field_mapping IS NOT INITIAL.

      READ TABLE mt_param_tab INTO DATA(ls_mimetype) WITH KEY name = gc_param_name-mimetype ##NO_TEXT.
      IF sy-subrc IS INITIAL.
        TRANSLATE ls_mimetype-value TO UPPER CASE.
      ENDIF.

      CASE ls_mimetype-value.
        WHEN gc_file_format-excel.
*          generate_excel_template( exporting
*                                     it_field_mapping = lt_field_mapping
*                                   importing
*                                     er_stream        = er_stream
*                                     ev_filename      = ev_filename ).

          " Generate title of template
          render_title( CHANGING
                          ct_blocks = lt_block ).

          " Generate header of table
          render_header( EXPORTING
                           it_fields_header = lt_field_mapping
                         CHANGING
                           ct_blocks        = lt_block ).

          DATA: lv_table_descr  TYPE REF TO cl_abap_tabledescr,
                lv_struct_descr TYPE REF TO cl_abap_structdescr,
                lt_columns      TYPE abap_compdescr_tab.


          FIELD-SYMBOLS: <fs_column> TYPE abap_compdescr,
                         <fs_value>  TYPE string,
                         <dyn_value> TYPE any.

          lv_table_descr  ?= cl_abap_typedescr=>describe_by_data( it_table ).
          lv_struct_descr ?= lv_table_descr->get_table_line_type( ).
          lt_columns      =  lv_struct_descr->components.


          "----------Excel data block--------"
          ls_block = VALUE #( ordinal_number = 3
                              location       = if_salv_export_appendix=>cs_appendix_location-top
                              name           = 'Template_data'
                              cells          = ls_cells ).


          "-----------Value format----------"
          ls_value_formatting = VALUE #( is_bold              = abap_false
                                         horizontal_alignment = if_salv_export_appendix=>cs_horizontal_alignment-center
                                         vertical_alignment   = if_salv_export_appendix=>cs_vertical_alignment-center
                                         foreground_color     = 'FF000000'
                                       ).

          LOOP AT it_table ASSIGNING FIELD-SYMBOL(<fs>).
            lv_column_index = 1.

            LOOP AT lt_columns ASSIGNING <fs_column>.
              ASSIGN COMPONENT <fs_column>-name OF STRUCTURE <fs> TO <dyn_value>.


              "Create the data cells for excel sheet
              ls_cell = VALUE #( row_index    = lv_row_index
                                 column_index = lv_column_index
                                 row_span     = 1
                                 column_span  = 0
                                 content_type = if_salv_export_appendix=>cs_cell_content_type-text
                                 formatting   = ls_value_formatting
                                 value        = <dyn_value> ).

              APPEND ls_cell TO ls_block-cells.
              lv_column_index = lv_column_index + 1.       "Move to next field

            ENDLOOP.
            lv_row_index    = lv_row_index + 1.
            lv_column_index = 1.

          ENDLOOP.
          INSERT ls_block INTO lt_block INDEX 3.
          CLEAR:ls_block,lv_row_index,lv_column_index, ls_cell.

          "**====================================Create excel=======================
          DATA: ls_stream               TYPE /iwbep/if_mgw_core_srv_runtime=>ty_s_media_resource,
                lv_content              TYPE xstring,
                lo_tool_xls             TYPE REF TO cl_salv_export_tool_ats,
                lt_mpa                  TYPE gty_t_mpa_transfer,
                lr_mpa                  TYPE REF TO data,
                lt_blocks               TYPE if_salv_export_appendix=>yts_block,
                lr_appendix             TYPE REF TO if_salv_export_appendix=>yts_block,
                lv_template_doc         TYPE xstring,
                lo_table_row_descriptor TYPE REF TO cl_abap_structdescr,
                lo_source_table_descr   TYPE REF TO cl_abap_tabledescr.

          GET REFERENCE OF lt_mpa INTO lr_mpa.
          DATA(lo_itab_services) = cl_salv_itab_services=>create_for_table_ref( lr_mpa ).

          TRY.

              lo_source_table_descr   ?= cl_abap_tabledescr=>describe_by_data_ref( lr_mpa ).
              lo_table_row_descriptor ?= lo_source_table_descr->get_table_line_type( ).

              lo_tool_xls = cl_salv_export_tool_ats_xls=>create_for_excel_from_ats(
                                                           io_itab_services       = lo_itab_services
                                                           io_source_struct_descr = lo_table_row_descriptor
                                                           it_aggregation_rules   = VALUE if_salv_service_types=>yt_aggregation_rule( )
                                                           it_grouping_rules      = VALUE if_salv_service_types=>yt_grouping_rule( ) ).

              DATA(lo_config) = lo_tool_xls->configuration( ).

              GET REFERENCE OF lt_block INTO lr_appendix.
              lo_config->if_salv_export_appendix~set_blocks( lr_appendix ).
              lo_tool_xls->read_result(  IMPORTING content  = lv_content  ).

            CATCH cx_salv_export_error.

          ENDTRY.

        WHEN gc_file_format-csv.

          DATA ls_result TYPE cl_mpa_asset_process_dpc_ext=>ty_s_file_data.

          CASE is_mpa_asset_data-scen_type.

            WHEN if_mpa_output=>gc_mpa_scen-transfer.
              ls_result-mass_transfer_data = it_table.
            WHEN if_mpa_output=>gc_mpa_scen-create.
              ls_result-mass_create_data = it_table.
            WHEN if_mpa_output=>gc_mpa_scen-adjustment.
              ls_result-mass_adjustment_data = it_table.
            WHEN if_mpa_output=>gc_mpa_scen-change.
              ls_result-mass_change_data = it_table.
            WHEN if_mpa_output=>gc_mpa_scen-retirement.
              ls_result-mass_retirement_data = it_table.
          ENDCASE.

          DATA: lv_header_str   TYPE string,
                lv_template_str TYPE string,
                lt_comment      TYPE string_table,
                lv_remark_1     TYPE string,
                lv_remark_2     TYPE string.

          " Get Commented text
          get_comment_text( IMPORTING et_comment = lt_comment ).

          lv_remark_1 = |"{ cl_fac_xlsx_parse_utils=>gc_comment_symbol } { lt_comment[ 1 ] }"| .
          lv_remark_2 = |"{ cl_fac_xlsx_parse_utils=>gc_comment_symbol } { lt_comment[ 2 ] }"|.

          concatenate_field_line( EXPORTING it_fields_mapping = lt_field_mapping
                                            iv_delimiter      = iv_csv_delimiter
                                  IMPORTING ev_line           = lv_header_str ).

          concatenate_data_line( EXPORTING is_mpa_data =  ls_result
                                           iv_mpa_Type = is_mpa_asset_data-scen_type
                                           iv_delimiter = iv_csv_delimiter
                                 IMPORTING ev_data_line = DATA(lv_data_str) ).

          CONCATENATE gv_title
                      lv_remark_1
                      lv_remark_2
                      lv_header_str
                      lv_data_str
                 INTO lv_template_str SEPARATED BY cl_abap_char_utilities=>cr_lf.

          mo_function_module->scms_string_to_xstring( EXPORTING text     = lv_template_str
                                                                mimetype = 'CSV'
                                                     IMPORTING  buffer   = lv_content ).

      ENDCASE.
    ENDIF.


    " *=======================================================================================*

    DATA: lv_returncode TYPE inri-returncode,
          lv_file_id    TYPE mpa_fileid.

    mo_function_module->number_get_next( EXPORTING nr_range_nr             = '01'
                                                   object                  = 'MPA_FILEID'
                                         IMPORTING number                  = lv_file_id
                                                   returncode              = lv_returncode
                                        EXCEPTIONS interval_not_found      = 1
                                                   number_range_not_intern = 2
                                                   object_not_found        = 3
                                                   quantity_is_0           = 4
                                                   quantity_is_not_1       = 5
                                                   interval_overflow       = 6
                                                   buffer_overflow         = 7
                                                   OTHERS                  = 8 ).

    IF sy-subrc <> 0 OR lv_returncode <> ' '.
      "message id sy-msgid type sy-msgty number sy-msgno with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 into data(lv_message).   "TODO Export the message
    ENDIF.

    IF lv_file_id IS NOT INITIAL.
      ev_fileid = lv_file_id.
      GET TIME STAMP FIELD DATA(lv_timestamp).
      CONCATENATE 'Result_' is_mpa_asset_data-file_name INTO DATA(lv_filename).
      DATA(ls_file_data) = VALUE mpa_asset_data( file_id     = lv_file_id
                                                 ernam       = sy-uname
                                                 erdat       = sy-datum
                                                 erzet       = sy-uzeit
                                                 timestamp   = lv_timestamp
                                                 scen_type   = is_mpa_asset_data-scen_type
                                                 file_status = '3'
                                                 file_name   = lv_filename
                                                 file_data   = lv_content
                                                 ref_file_id = is_mpa_asset_data-file_id ).

      "Insert file data into MPA_ASSET_DATA table
      INSERT mpa_asset_data FROM ls_file_data.
      IF sy-subrc IS INITIAL.
        COMMIT WORK.
      ENDIF.
    ENDIF.

  ENDMETHOD.

  METHOD concatenate_data_line.

    TYPES:
      BEGIN OF ty_objs,
        id   TYPE char10,
        data TYPE REF TO data,
      END OF ty_objs.

    TYPES lty_tt_fieldnames TYPE TABLE OF fieldname.

    DATA: lt_act                  TYPE STANDARD TABLE OF ty_objs,
          ls_act                  LIKE LINE OF lt_act,
          lt_exp                  TYPE STANDARD TABLE OF ty_objs,
          ls_exp                  LIKE LINE OF lt_exp,
          lo_table_descr          TYPE REF TO cl_abap_tabledescr,
          lo_line                 TYPE REF TO cl_abap_datadescr,
          lo_struct               TYPE REF TO cl_abap_structdescr,
          lt_fieldname            TYPE lty_tt_fieldnames,
          lt_range_date_fieldname TYPE RANGE OF fieldname,
          lv_line_count           TYPE i VALUE 1,
          lv_line_type            TYPE string,
          lv_line_s               TYPE string,
          lt_line_table           TYPE string_table.

    DATA lt_fields TYPE string_table.

    FIELD-SYMBOLS : <lt_asset_data> TYPE ANY TABLE.

    CASE iv_mpa_type.

      WHEN if_mpa_output=>gc_mpa_scen-transfer.
        ASSIGN  is_mpa_data-mass_transfer_data TO <lt_asset_data>.
      WHEN if_mpa_output=>gc_mpa_scen-create.
        ASSIGN  is_mpa_data-mass_create_data TO <lt_asset_data>.
      WHEN if_mpa_output=>gc_mpa_scen-change.
        ASSIGN  is_mpa_data-mass_change_data TO <lt_asset_data>.
      WHEN if_mpa_output=>gc_mpa_scen-adjustment.
        ASSIGN  is_mpa_data-mass_adjustment_data TO <lt_asset_data>.
      WHEN if_mpa_output=>gc_mpa_scen-retirement.
        ASSIGN  is_mpa_data-mass_retirement_data TO <lt_asset_data>.
    ENDCASE.

    lo_table_descr ?=  cl_abap_typedescr=>describe_by_data( p_data =  <lt_asset_data> ).
    lo_line = lo_table_descr->get_table_line_type( ).

    IF lo_line->kind = cl_abap_typedescr=>kind_struct.
      lo_struct ?= lo_line.
      "get table name
      lv_line_type = lo_line->get_relative_name( ).

      "get field names
      LOOP AT lo_struct->components ASSIGNING FIELD-SYMBOL(<lv_comp>).
        APPEND <lv_comp>-name TO lt_fieldname.
        IF <lv_comp>-type_kind = 'D'.
          APPEND VALUE #( sign = 'I'
                          option = 'EQ'
                          low = <lv_comp>-name ) TO lt_range_date_fieldname.
        ENDIF.
      ENDLOOP.
    ENDIF.

    LOOP AT <lt_asset_data> ASSIGNING FIELD-SYMBOL(<ls_asset_data>).

      LOOP AT lt_fieldname ASSIGNING FIELD-SYMBOL(<ls_fieldname>).

        ASSIGN COMPONENT <ls_fieldname> OF STRUCTURE <ls_asset_data> TO FIELD-SYMBOL(<ls_asset_field>).

        IF sy-subrc = 0.
          IF <ls_fieldname> IN lt_range_date_fieldname AND lt_range_date_fieldname IS NOT INITIAL.

            DATA(lv_date) = |{ <ls_asset_field>(4) }-{ <ls_asset_field>+4(2) }-{ <ls_asset_field>+6(2) }|.
            APPEND lv_date TO lt_fields.
          ELSE.
            APPEND <ls_asset_field> TO lt_fields.
          ENDIF.
        ELSE.
          APPEND space TO lt_fields.
        ENDIF.

      ENDLOOP.

      CONCATENATE LINES OF lt_fields INTO lv_line_s SEPARATED BY iv_delimiter.

      APPEND lv_line_s TO lt_line_table.
      CLEAR : lt_fields,lv_line_s.
    ENDLOOP.

    CONCATENATE LINES OF lt_line_table INTO  ev_data_line SEPARATED BY cl_abap_char_utilities=>cr_lf.

  ENDMETHOD.


  METHOD generate_csv_download_file.

    DATA: lv_header_str   TYPE string,
          lv_template_str TYPE string,
          ls_stream       TYPE /iwbep/if_mgw_core_srv_runtime=>ty_s_media_resource,
          lv_content      TYPE xstring,
          lv_header       TYPE string,
          lt_comment      TYPE string_table,
          lv_remark_1     TYPE string,
          lv_remark_2     TYPE string.

    " Get Commented text
    get_comment_text( IMPORTING et_comment = lt_comment ).

    lv_remark_1 = |"{ cl_fac_xlsx_parse_utils=>gc_comment_symbol } { lt_comment[ 1 ] }"| .
    lv_remark_2 = |"{ cl_fac_xlsx_parse_utils=>gc_comment_symbol } { lt_comment[ 2 ] }"|.

    concatenate_field_line( EXPORTING it_fields_mapping = it_field_mapping
                                      iv_delimiter      = iv_delimiter
                            IMPORTING ev_line           = lv_header_str ).

    concatenate_data_line( EXPORTING is_mpa_data =  is_data
                                     iv_mpa_Type = iv_mpa_type
                                     iv_delimiter = iv_delimiter
                           IMPORTING ev_data_line = DATA(lv_data_str) ).

    CONCATENATE gv_title
                lv_remark_1
                lv_remark_2
                lv_header_str
                lv_data_str
           INTO lv_template_str SEPARATED BY cl_abap_char_utilities=>cr_lf.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text     = lv_template_str
        mimetype = 'CSV'
      IMPORTING
        buffer   = lv_content.

    " Set file content
    ls_stream-value     = lv_content.
    ls_stream-mime_type = gc_mime_type-app_csv.

    copy_data_to_ref( EXPORTING is_data = ls_stream
                      CHANGING cr_data = er_stream ).

    " Set file name that get from request URL
    READ TABLE mt_param_tab INTO DATA(ls_filename) WITH KEY name = gc_param_name-filename.
    IF sy-subrc IS INITIAL.
      ev_filename = ls_filename-value.
    ELSE.
      ev_filename = gc_file_name-csv_file       ##NO_TEXT.
    ENDIF.

  ENDMETHOD.


  METHOD zcreate_info_for_new_sheet.

*   Fill the Sheet Info
    rs_sheet_info-name = iv_sheet_name.
    TRY.
        rs_sheet_info-rid  = io_workbookpart->get_id_for_part( io_worksheet_part ).
      CATCH cx_openxml_not_found.
*       This should never happen as we just created the sheet
        ASSERT 1 = 2.
    ENDTRY.

*   Generate a new sheet ID
*    ADD 1 TO mv_last_sheet_id.
    rs_sheet_info-sheet_id = 2.

*   Add the info to the table
*    INSERT rs_sheet_info INTO TABLE ms_workbook_data-sheet_ids_htab.

  ENDMETHOD.


  METHOD zupdate_wb_xml_after_add_sheet.

    DATA lv_workbook_xml  TYPE xstring.
    lv_workbook_xml =  io_workbookpart->get_data( ).

    CALL TRANSFORMATION xl_mpa_insert_sheet
             PARAMETERS active_sheet = 2
                        sheet_name   = is_sheet_info-name
                        sheet_id     = is_sheet_info-sheet_id
                        sheet_rid    = is_sheet_info-rid
             SOURCE XML lv_workbook_xml
             RESULT XML lv_workbook_xml.

  ENDMETHOD.


  METHOD get_asset_data.

    DATA lo_xlsx_parse_util TYPE REF TO zcl_xlsx_parse_util.
    lo_xlsx_parse_util = CAST #( zcl_xlsx_parse_util=>get_instance( ) ).

    lo_xlsx_parse_util->parse_xlsx( EXPORTING  ix_file     = ix_file
                                    IMPORTING ev_mpa_type = DATA(lv_mpa_file_type)
                                              et_asset    = r_result
                                              et_message  = DATA(et_message) ).

  ENDMETHOD.

ENDCLASS.













