*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section

    TYPES: BEGIN OF gty_s_ooxml_worksheet,
             name     TYPE string,
             id       TYPE string,
             location TYPE string,
           END OF gty_s_ooxml_worksheet ,
           gty_t_ooxml_worksheets TYPE STANDARD TABLE OF gty_s_ooxml_worksheet.

    TYPES: BEGIN OF gty_s_numfmtid,
             id         TYPE i,
             formatcode TYPE string,
           END OF gty_s_numfmtid ,
           gty_t_numfmtids TYPE STANDARD TABLE OF gty_s_numfmtid.

    TYPES: gty_t_dd04l TYPE STANDARD TABLE OF dd04l.

    INTERFACE lif_mpa_xlsx_parse_util.

      "! Convert the excel data from XSTRING format to tabular format
      METHODS transform_xstring_2_tab
        IMPORTING
          !ix_file              TYPE xstring
          !iv_download_file_ind TYPE abap_bool OPTIONAL
        EXPORTING
          !et_table             TYPE mpa_t_index_value_pair
          !et_dd04l             TYPE gty_t_dd04l
        CHANGING
          !ct_fieldinfo         TYPE dd_x031l_table
        RAISING
          cx_mpa_exception_handler
          cx_openxml_not_found
          cx_openxml_format .

      "! Get the attribute from excel node
      METHODS get_attr_from_node
        IMPORTING
          !iv_name        TYPE string
          !io_node        TYPE REF TO if_ixml_node
        RETURNING
          VALUE(rv_value) TYPE string .

    ENDINTERFACE.

    INTERFACE lif_function_module.

      TYPES lty_tt_x255 TYPE STANDARD TABLE OF x255.

      "! Call function module date_check_plausibility
      METHODS date_check_plausibility
        IMPORTING  date TYPE syst_datum
        EXCEPTIONS plausibility_check_failed.

      METHODS scms_xstring_to_binary
        IMPORTING buffer                 TYPE xstring
                  VALUE(append_to_table) TYPE c DEFAULT space
        EXPORTING VALUE(output_length)   TYPE i
        CHANGING  binary_tab             TYPE lty_tt_x255.

      METHODS scms_binary_to_string
        IMPORTING  VALUE(input_length)  TYPE i
                   VALUE(first_line)    TYPE i DEFAULT 0
                   VALUE(last_line)     TYPE i DEFAULT 0
                   VALUE(mimetype)      TYPE c DEFAULT space
                   VALUE(encoding)      TYPE abap_encoding OPTIONAL
        EXPORTING  text_buffer          TYPE string
                   VALUE(output_length) TYPE i
        CHANGING   binary_tab           TYPE lty_tt_x255
        EXCEPTIONS failed.

    ENDINTERFACE.

    CLASS lcl_function_module DEFINITION.

      PUBLIC SECTION.
        INTERFACES lif_function_module.

        CLASS-DATA go_instance TYPE REF TO lif_function_module.
        CLASS-METHODS get_instance RETURNING VALUE(ro_instance)  TYPE REF TO lif_function_module.

    ENDCLASS.

    CLASS lcl_function_module_mock DEFINITION.

      PUBLIC SECTION.
        INTERFACES lif_function_module.

    ENDCLASS.

    CLASS lcl_mpa_xlsx_parse_util DEFINITION.
      PUBLIC SECTION.
        INTERFACES lif_mpa_xlsx_parse_util.

        CONSTANTS gc_transfer_struct_name TYPE string VALUE 'MPA_S_ASSET_TRANSFER' ##NO_TEXT.
        CONSTANTS gc_create_struct_name   TYPE string VALUE 'MPA_S_ASSET_CREATE'   ##NO_TEXT.
        DATA dateformat1904 TYPE abap_bool.

        "! Add the additional format for excel columns data
        METHODS add_additional_format
          CHANGING
            !ct_numfmtids TYPE gty_t_numfmtids .

        "! Format the cell with number format
        METHODS convert_cell_value_by_numfmt
          IMPORTING
            !iv_cell_value            TYPE string
            !iv_number_format         TYPE string
          RETURNING
            VALUE(ev_formatted_value) TYPE string .

        "! Get the date from excel internal numbering
        METHODS convert_ser_val_to_date_time
          IMPORTING
            !iv_serial_value_string TYPE string
          EXPORTING
            !ev_date                TYPE d
            !ev_time                TYPE t .

        "! Convert from long to date format
        METHODS convert_long_to_date
          IMPORTING
            !iv_date_string TYPE string
          RETURNING
            VALUE(rv_date)  TYPE d .

        "! Convert the date from DECIMAL to 'HHMMSS' format
        METHODS convert_dec_time_to_hhmmss
          IMPORTING
            !iv_dec_time_string TYPE string
          RETURNING
            VALUE(rv_time)      TYPE t .
      PRIVATE SECTION.
        METHODS parse_document_to_xml
          IMPORTING
            ix_sheet               TYPE xstring
          RETURNING
            VALUE(ro_xml_document) TYPE REF TO if_ixml_document.

    ENDCLASS.

    CLASS lcl_mpa_xlsx_prs_util_mock DEFINITION.

      PUBLIC SECTION.
        INTERFACES : lif_mpa_xlsx_parse_util.

    ENDCLASS.
