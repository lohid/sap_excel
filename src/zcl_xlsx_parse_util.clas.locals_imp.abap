*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations


    CLASS lcl_mpa_xlsx_parse_util IMPLEMENTATION.

      METHOD lif_mpa_xlsx_parse_util~transform_xstring_2_tab .

        DATA: lo_xlsx_doc        TYPE REF TO cl_xlsx_document,
              lo_workbookpart    TYPE REF TO cl_xlsx_workbookpart,
              lx_workbook        TYPE xstring,
              lt_worksheets      TYPE gty_t_ooxml_worksheets,
              lo_worksheetpart   TYPE REF TO cl_xlsx_sharedstringspart,
              lx_file            TYPE xstring,
              lt_shared_string   TYPE TABLE OF string,
              lx_sheet           TYPE xstring,
              lv_dim             TYPE string,
              lo_worksheet       TYPE REF TO cl_openxml_part,
              lo_ixml            TYPE REF TO if_ixml,
              lo_streamfactory   TYPE REF TO if_ixml_stream_factory,
              lo_parser          TYPE REF TO if_ixml_parser,
              lo_istream         TYPE REF TO if_ixml_istream,
              lo_node            TYPE REF TO if_ixml_node,
              lo_xml_document    TYPE REF TO if_ixml_document,
              lo_row             TYPE REF TO if_ixml_node,
              lo_rows            TYPE REF TO if_ixml_node_collection,
              lo_row_iterator    TYPE REF TO if_ixml_node_iterator,
              lv_row_index       TYPE string,
              lo_row_element     TYPE REF TO if_ixml_element,
              lo_col             TYPE REF TO if_ixml_node,
              lo_cols            TYPE REF TO if_ixml_node_collection,
              lo_col_iterator    TYPE REF TO if_ixml_node_iterator,
              lv_col_index       TYPE string,
              lo_cell_element    TYPE REF TO if_ixml_element,
              lo_col_element     TYPE REF TO if_ixml_element,
              lv_cell_value      TYPE string,
              lv_cell_index      TYPE string,
              lo_xlsx_stylesheet TYPE REF TO cl_xlsx_stylespart,
              lx_styles          TYPE xstring,
              lt_cellxfs         TYPE STANDARD TABLE OF string,
              lv_style_id        TYPE string,
              lt_numfmtids       TYPE gty_t_numfmtids,
              ls_numfmtids       TYPE gty_s_numfmtid,
              ls_df1904          TYPE string,
              lt_fieldinfo       TYPE dd_x031l_table,
              lr_type_desc       TYPE REF TO cl_abap_typedescr,
              ls_line            TYPE mpa_s_index_value_pair,
              ls_cell            TYPE mpa_s_index_value_pair.

        FIELD-SYMBOLS: <fs_celldata> TYPE any,
                       <ft_linedata> TYPE mpa_t_index_value_pair.

        IF ix_file IS INITIAL.
          RAISE EXCEPTION TYPE cx_openxml_format
            EXPORTING
              textid = cx_openxml_format=>cx_openxml_empty.
        ENDIF.

        "===== get work sheet =====
        lo_xlsx_doc     = cl_xlsx_document=>load_document( ix_file ).
        lo_workbookpart = lo_xlsx_doc->get_workbookpart( ).
        lx_workbook     = lo_workbookpart->get_data( ).

        CALL TRANSFORMATION xl_mpa_get_worksheets SOURCE XML lx_workbook  RESULT worksheets = lt_worksheets.

        IF lt_worksheets IS INITIAL.
          RAISE EXCEPTION TYPE cx_openxml_format
            EXPORTING
              textid = cx_openxml_format=>cx_openxml_empty.
        ENDIF.

        "===== get work sheet data =====
        lo_worksheetpart = lo_workbookpart->get_sharedstringspart( ).
        lx_file          = lo_worksheetpart->get_data( ).

        CALL TRANSFORMATION xl_mpa_get_shared_strings SOURCE XML lx_file RESULT shared_strings = lt_shared_string.

        "===== get style and format =====
        lo_xlsx_stylesheet = lo_workbookpart->get_stylespart( ).
        lx_styles          = lo_xlsx_stylesheet->get_data( ).

        CALL TRANSFORMATION xl_mpa_get_cellxfs   SOURCE XML lx_styles RESULT numfmids = lt_cellxfs.
        CALL TRANSFORMATION xl_mpa_get_numfmtids SOURCE XML lx_styles RESULT numfmts  = lt_numfmtids.

        add_additional_format( CHANGING ct_numfmtids = lt_numfmtids ).

        CALL TRANSFORMATION xl_mpa_get_date_format SOURCE XML lx_workbook RESULT dateformat_1904 = ls_df1904.
        IF ls_df1904 = '1'.
          dateformat1904 = abap_true.
        ELSE.
          dateformat1904 = abap_false.
        ENDIF.

        "===== get shared strings =====
        lo_worksheet = lo_workbookpart->get_part_by_id( lt_worksheets[ 1 ]-location ).
        lx_sheet = lo_worksheet->get_data( ).
        CALL TRANSFORMATION xl_mpa_get_sheet_dimension SOURCE XML lx_sheet RESULT dimension = lv_dim.

        "===== get field information =====
        lr_type_desc = cl_abap_structdescr=>describe_by_name( gc_transfer_struct_name ).
        lt_fieldinfo = lr_type_desc->get_ddic_object( ).
        APPEND LINES OF lt_fieldinfo TO ct_fieldinfo.

        CLEAR lt_fieldinfo.
        lr_type_desc = cl_abap_structdescr=>describe_by_name( gc_create_struct_name ).
        lt_fieldinfo = lr_type_desc->get_ddic_object( ).
        APPEND LINES OF lt_fieldinfo TO ct_fieldinfo.

        DATA lt_data_element_all TYPE TABLE OF rollname.
        LOOP AT ct_fieldinfo ASSIGNING FIELD-SYMBOL(<fs_fieldinfo>).
          APPEND <fs_fieldinfo>-rollname TO lt_data_element_all.
        ENDLOOP.

        SELECT rollname, domname, leng, decimals, outputlen FROM dd04l INTO CORRESPONDING FIELDS OF TABLE @et_dd04l
          FOR ALL ENTRIES IN @lt_data_element_all WHERE as4local = 'A' AND rollname = @lt_data_element_all-table_line.

        lo_xml_document = parse_document_to_xml( lx_sheet ).

        "========== Process Begin ==========
        "===== get rows from xml =====
        lo_rows         = lo_xml_document->get_elements_by_tag_name_ns( name = 'row'  uri = if_xl_types=>ns_ooxml_ssheet_main ).  "get all data rows
        lo_row_iterator = lo_rows->create_iterator( ).
        lo_row          = lo_row_iterator->get_next( ).


        "===== loop rows and process =====
        WHILE lo_row IS NOT INITIAL.
          " 1. get line number
          lv_row_index = lif_mpa_xlsx_parse_util~get_attr_from_node( iv_name  = 'r' io_node  = lo_row ).  " get row index

          " 2. create line data.
          ls_line-index = lv_row_index.
          CREATE DATA ls_line-value TYPE mpa_t_index_value_pair.
          ASSIGN ls_line-value->* TO <ft_linedata>.

          " 3. loop cols from rol =====
          lo_row_element ?= lo_row->query_interface( ixml_iid_element ).
          lo_cols         = lo_row_element->get_elements_by_tag_name_ns( name = 'c' uri = if_xl_types=>ns_ooxml_ssheet_main ).  " get all columns
          IF lo_cols IS INITIAL.
            "step to next col
            lo_col = lo_col_iterator->get_next( ).
            CONTINUE.
          ENDIF.
          lo_col_iterator = lo_cols->create_iterator( ).
          lo_col          = lo_col_iterator->get_next( ).

          "===== loop cols =====
          WHILE lo_col IS NOT INITIAL.
            lo_col_element ?= lo_col->query_interface( ixml_iid_element ).
            lo_cell_element = lo_col_element->find_from_name_ns( name = 'v'  uri = if_xl_types=>ns_ooxml_ssheet_main ). "get cell from colums
            IF lo_cell_element IS INITIAL.
              "step to next col
              lo_col = lo_col_iterator->get_next( ).
              CONTINUE.
            ENDIF.

            " 1. get cell value
            lv_cell_value = lo_cell_element->get_value( ).
            IF lif_mpa_xlsx_parse_util~get_attr_from_node( iv_name = 't' io_node = lo_col ) = 's'. " get cell value's data type, 's' means string
              READ TABLE lt_shared_string INDEX ( lv_cell_value + 1 ) INTO lv_cell_value.
            ENDIF.
"lohid_new
*            IF iv_download_file_ind = abap_false.
*              "date conversion should not be done while downloading the file using Download Asset File option !
*              " transform the date format
*              READ TABLE lt_cellxfs   INTO lv_style_id  INDEX ( lif_mpa_xlsx_parse_util~get_attr_from_node( iv_name = 's' io_node = lo_col ) + 1 ).
*              READ TABLE lt_numfmtids INTO ls_numfmtids WITH KEY id = lv_style_id.
*              lv_cell_value = convert_cell_value_by_numfmt( iv_cell_value = lv_cell_value iv_number_format = ls_numfmtids-formatcode ).
*              REPLACE ALL OCCURRENCES OF REGEX '^\s*|\s*$' IN lv_cell_value WITH ''.
*            ENDIF.

            IF lv_cell_value IS INITIAL.
              "step to next col
              lo_col = lo_col_iterator->get_next( ).
              CONTINUE.
            ENDIF.

            " 2. get cell index and col index
            lv_cell_index   = lif_mpa_xlsx_parse_util~get_attr_from_node( iv_name  = 'r' io_node  = lo_col ).
            "find col index
            DATA(lv_length) = strlen( lv_cell_index ) - strlen( lv_row_index ).
            lv_col_index    = lv_cell_index(lv_length).

            " 3. set cell index and data
            ls_cell-index = lv_col_index.
            CREATE DATA ls_cell-value TYPE string.
            ASSIGN ls_cell-value->* TO <fs_celldata>.
            <fs_celldata> = lv_cell_value.

            " 4. add cell into line
            INSERT ls_cell INTO TABLE <ft_linedata>.

            " 5. step to next col
            lo_col = lo_col_iterator->get_next( ).

            CLEAR lv_cell_value. "lohid
          ENDWHILE.
          "===== loop cols end =====

          " 4. add line to table
          IF <ft_linedata> IS NOT INITIAL.
            INSERT ls_line INTO TABLE et_table.
          ENDIF.

          " 5. step to next row
          lo_row = lo_row_iterator->get_next( ).

        ENDWHILE.


      ENDMETHOD.

      METHOD parse_document_to_xml.

        DATA lo_ixml TYPE REF TO if_ixml.
        DATA lo_streamfactory TYPE REF TO if_ixml_stream_factory.
        DATA lo_parser TYPE REF TO if_ixml_parser.
        DATA lo_istream TYPE REF TO if_ixml_istream.

        "===== parse document to xml =====
        lo_ixml          = cl_ixml=>create( ).
        lo_streamfactory = lo_ixml->create_stream_factory( ).
        lo_istream       = lo_streamfactory->create_istream_xstring( ix_sheet ).
        ro_xml_document  = lo_ixml->create_document( ).
        lo_parser        = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                   istream        = lo_istream
                                                   document       = ro_xml_document ).
        lo_parser->parse( ).

      ENDMETHOD.



      METHOD lif_mpa_xlsx_parse_util~get_attr_from_node.

        DATA: lo_attrib_map TYPE REF TO if_ixml_named_node_map,
              lo_attribute  TYPE REF TO if_ixml_attribute,
              lo_attr_node  TYPE REF TO if_ixml_node.

        TYPE-POOLS ixml.
        CHECK io_node IS NOT INITIAL.

        lo_attrib_map = io_node->get_attributes( ).
        CHECK lo_attrib_map IS NOT INITIAL.

        lo_attr_node = lo_attrib_map->get_named_item( iv_name ).
        CHECK lo_attr_node IS NOT INITIAL.

*lo_attrib_map->set_named_item_ns( node =  ).

        lo_attribute ?= lo_attr_node->query_interface( ixml_iid_attribute ).
        rv_value = lo_attribute->get_value( ).

      ENDMETHOD.

      METHOD add_additional_format.
        DATA ls_format TYPE gty_s_numfmtid.
        ls_format-id = 0.
        ls_format-formatcode = `General`.                   "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 1.
        ls_format-formatcode = `0`.                         "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 2.
        ls_format-formatcode = `0.00`.                      "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 3.
        ls_format-formatcode = `#,##0`.                     "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 4.
        ls_format-formatcode = `#,##0.00`.                  "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 9.
        ls_format-formatcode = `0%`.                        "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 10.
        ls_format-formatcode = `0.00%`.                     "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 11.
        ls_format-formatcode = `0.00E+00`.                  "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 12.
        ls_format-formatcode = `# ?/?`.                     "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 13.
        ls_format-formatcode = `# ??/??`.                   "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 14.
        ls_format-formatcode = `mm-dd-yy`.                  "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 15.
        ls_format-formatcode = `d-mmm-yy`.                  "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 16.
        ls_format-formatcode = `d-mmm`.                     "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 17.
        ls_format-formatcode = `mmm-yy`.                    "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 18.
        ls_format-formatcode = `h:mm AM/PM`.                "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 19.
        ls_format-formatcode = `h:mm:ss AM/PM`.             "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 20.
        ls_format-formatcode = `h:mm`.                      "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 21.
        ls_format-formatcode = `h:mm:ss`.                   "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 22.
        ls_format-formatcode = `m/d/yy h:mm`.               "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 37.
        ls_format-formatcode = `#,##0;(#,##0)`.             "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 38.
        ls_format-formatcode = `#,##0;[Red](#,##0)`.        "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 39.
        ls_format-formatcode = `#,##0.00;(#,##0.00)`.       "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 40.
        ls_format-formatcode = `#,##0.00;[Red](#,##0.00)`.  "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 45.
        ls_format-formatcode = `mm:ss`.                     "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 46.
        ls_format-formatcode = `[h]:mm:ss`.                 "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 47.
        ls_format-formatcode = `mmss.0`.                    "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 48.
        ls_format-formatcode = `##0.0E+0`.                  "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
        ls_format-id = 49.
        ls_format-formatcode = `@`.                         "#EC NOTEXT
        APPEND ls_format TO ct_numfmtids.
      ENDMETHOD.

      METHOD convert_cell_value_by_numfmt.

        DATA: lv_clean_format TYPE string,
              lv_date         TYPE d.

        ev_formatted_value = ''.
        lv_clean_format    = iv_number_format.

        REPLACE REGEX '"[^"]*"' IN lv_clean_format WITH ''.
        REPLACE REGEX '\[Red\]' IN lv_clean_format WITH ''.

        IF cl_abap_matcher=>matches( pattern = '.*(y+|m+|d+|h+|s+).*' text = lv_clean_format ) = abap_true.

          convert_ser_val_to_date_time(
            EXPORTING
              iv_serial_value_string = iv_cell_value
            IMPORTING
              ev_date                = lv_date ).

          ev_formatted_value = lv_date.
        ELSE.
          ev_formatted_value = iv_cell_value.
        ENDIF.

        IF ev_formatted_value = ''.
          ev_formatted_value = iv_cell_value.
        ENDIF.

      ENDMETHOD.

      METHOD convert_ser_val_to_date_time.

        DATA: lv_date_str  TYPE string,
              lv_time_str  TYPE string,
              lv_date_time TYPE decfloat34.

        TRY.
            lv_date_time = iv_serial_value_string.
          CATCH cx_root.
            EXIT.
        ENDTRY.

        lv_date_str = floor( lv_date_time ).
        IF lv_date_str NE '0'.
          ev_date = convert_long_to_date( lv_date_str ).
        ENDIF.

        lv_time_str = frac( lv_date_time ).
        IF lv_time_str NE '0'.
          ev_time = convert_dec_time_to_hhmmss( lv_time_str ).
        ENDIF.

      ENDMETHOD.


      METHOD convert_long_to_date.

        DATA lv_num_days TYPE i.

        lv_num_days = floor( iv_date_string ).

        IF dateformat1904 = abap_false.
          rv_date = '18991231'.
          IF iv_date_string > 59.
            rv_date = rv_date + lv_num_days - 1.
          ELSE.
            rv_date = rv_date + lv_num_days.
          ENDIF.
        ELSE.
          rv_date = '19040101'.
          rv_date = rv_date + lv_num_days.
        ENDIF.

      ENDMETHOD.


      METHOD convert_dec_time_to_hhmmss.

        DATA: lv_dec_time   TYPE decfloat16,
              lv_hour       TYPE i,
              lv_hour_str   TYPE string,
              lv_minute     TYPE i,
              lv_minute_str TYPE string,
              lv_second     TYPE decfloat16.

        TRY.
            lv_dec_time = iv_dec_time_string.
          CATCH cx_root.
            rv_time = iv_dec_time_string.
            EXIT.
        ENDTRY.

        lv_dec_time = frac( lv_dec_time ).
        lv_dec_time = round( val = lv_dec_time dec = 15 ).
        lv_dec_time = lv_dec_time * 24.
        lv_hour     = floor( lv_dec_time ).
        lv_dec_time = ( lv_dec_time - lv_hour ) * 60.
        lv_minute   = floor( lv_dec_time ).
        lv_second   = round( val = ( ( lv_dec_time - lv_minute ) * 60 ) dec = 3 ).

        IF lv_second >= 60.
          lv_second = 0.
          lv_minute = lv_minute + 1.
        ENDIF.

        IF lv_hour < 10.
          lv_hour_str = '0' && lv_hour.
        ELSE.
          lv_hour_str = lv_hour.
        ENDIF.
        cl_abap_string_utilities=>del_trailing_blanks( CHANGING str = lv_hour_str ).

        IF lv_minute < 10.
          lv_minute_str = '0' && lv_minute.
        ELSE.
          lv_minute_str = lv_minute.
        ENDIF.
        cl_abap_string_utilities=>del_trailing_blanks( CHANGING str = lv_minute_str ).

        rv_time = lv_hour_str && lv_minute_str &&  lv_second .

      ENDMETHOD.

    ENDCLASS.


    CLASS lcl_mpa_xlsx_prs_util_mock IMPLEMENTATION.

      METHOD lif_mpa_xlsx_parse_util~transform_xstring_2_tab .

      ENDMETHOD.

      METHOD lif_mpa_xlsx_parse_util~get_attr_from_node.

      ENDMETHOD.

    ENDCLASS.

    CLASS lcl_function_module IMPLEMENTATION.

      METHOD lif_function_module~date_check_plausibility.

        CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
          EXPORTING
            date                      = date
          EXCEPTIONS
            plausibility_check_failed = 1
            OTHERS                    = 2.

        IF sy-subrc <> 0.
          RAISE    plausibility_check_failed.
        ENDIF.

      ENDMETHOD.

      METHOD get_instance.

        go_instance = COND #( WHEN go_instance IS BOUND
                              THEN go_instance
                              ELSE NEW lcl_function_module( ) ).

        ro_instance = go_instance.

      ENDMETHOD.

      METHOD lif_function_module~scms_xstring_to_binary.

        CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
          EXPORTING
            buffer        = buffer
          IMPORTING
            output_length = output_length
          TABLES
            binary_tab    = binary_tab.

      ENDMETHOD.

      METHOD lif_function_module~scms_binary_to_string.

        CALL FUNCTION 'SCMS_BINARY_TO_STRING'
          EXPORTING
            input_length = input_length
            mimetype     = mimetype
          IMPORTING
            text_buffer  = text_buffer
          TABLES
            binary_tab   = binary_tab
          EXCEPTIONS
            failed       = 1
            OTHERS       = 2.

        IF sy-subrc <> 0.
          RAISE  failed.
        ENDIF.

      ENDMETHOD.

    ENDCLASS.

    CLASS lcl_function_module_mock IMPLEMENTATION.

      METHOD lif_function_module~date_check_plausibility.
        "do nothing
      ENDMETHOD.

      METHOD lif_function_module~scms_binary_to_string.

        text_buffer = |Asset Mass Transfer;;;;;;;;;;;;;;;;;;;;;;;;;;;;\r\n| &&
                      |//Andern Sie die Vorlage nicht. Fugen Sie stattdessen gemass Szenario Daten im entsprechenden Feld hinzu.;;;;;;;;;;;;;;;;;;;;;;;;;;;;\r\n| &&
                      |//Mit einem Asterisk  markierte Felder sind Pflichtfelder. Nach dem Fullen der Vorlage laden Sie diese zur Weiterverarbeitung hoch.;;;;;;;;;;;;;;;;;;;;;;;;;;;;\r\n| &&
                      |SLNO;BLART;BLDAT;BUDAT;BZDAT;SGTXT;MONAT;WWERT;BUKRS;ANLN1;ANLN2;ACC_PRINCIPLE;AFABER;PBUKRS;PANL1;PANL2;ANLKL;KOSTL;TEXT;TRAVA;ANBTR;WAERS;MENGE;MEINS;PROZS;XANEU;RECID;XBLNR;DZUONR\r\n| &&
                      |*Row;DocType;*DocDate;*PostDate;*AssetDate;ItText);FisPer;*Transdate;*CompCode;*AssetNum;Subnum;AP;DepArea;PComp;MainPAsset;PartnerSubnumber;AssetClass;CC;Desc;Trans var;Amt;CurKey;Quan;Unit;Perc;*Ind;RecInd;RDocNumber;Assnum\r\n|
                      && |4;;10/25/2020;10/25/2020;10/25/2020;;;10/25/2020;JVU1;10000000166;0;;;;10000000166;1;;;;4;2;EUR;;;;X;;;\r\n| &&
                      |5;;06/15/2020;06/15/2020;06/15/2020;;;06/15/2020;JVU1;10000000073;0;;;;10000000074;0;;;;4;10;EUR;10;kg;;X;;;\r\n|.

      ENDMETHOD.

      METHOD lif_function_module~scms_xstring_to_binary.
        "do nothing
      ENDMETHOD.

    ENDCLASS.
