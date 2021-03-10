*&---------------------------------------------------------------------*
*& Report zexport_xlsx
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zexport_xlsx.

DATA: lv_file              TYPE xstring,
      lo_doc               TYPE REF TO zif_mpa_xlsx_doc,
      lt_sheet_info        TYPE zcl_mpa_xlsx=>gty_th_sheet_info,
      lo_sheet             TYPE REF TO zif_mpa_xlsx_sheet,
      lo_sheet_2             TYPE REF TO zif_mpa_xlsx_sheet,
      lo_sheet_3             TYPE REF TO zif_mpa_xlsx_sheet,
      lt_file              TYPE cpt_x255,
      lv_filename          TYPE localfile,
      lv_bytes_transferred TYPE i.

PARAMETERS p_fname TYPE localfile OBLIGATORY.

DATA(lo_xlsx) = zcl_mpa_xlsx=>get_instance( ).
lo_doc = lo_xlsx->create_doc( ).



lt_sheet_info = lo_doc->get_sheets( ).
lo_sheet = lo_doc->get_sheet_by_id( lt_sheet_info[ 1 ]-sheet_id ).
lo_sheet->change_sheet_name( iv_new_name = 'sheet_1' ). "not working lohid
lo_sheet->set_cell_content( iv_row = 1 iv_column = 1 iv_value = 'TagID' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 2 iv_value = 'AmountDate' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 3 iv_value = 'AmountTime' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 4 iv_value = 'AmountValue' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 5 iv_value = 'AmountUnit' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 6 iv_value = 'FaultInd' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 7 iv_value = 'CalibrationInd' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 8 iv_value = 'OutOfPrecisenessOperator' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 9 iv_value = 'NotAvailableInd' ).
lo_sheet->set_cell_content( iv_row = 1 iv_column = 10 iv_value = 'Remark' ).

lo_doc->add_new_sheet( iv_sheet_name = 'sheet_2'  ).
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

lo_doc->add_new_sheet( iv_sheet_name = 'sheet_3'  ).
*CATCH cx_openxml_format.
*CATCH cx_openxml_not_allowed.
*CATCH cx_dynamic_check.
lo_sheet_3 = lo_doc->get_sheet_by_id( 3 ).

lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 1 iv_value = 'TagID_3' ).
*lo_sheet_3->
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 2 iv_value = 'AmountDate_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 3 iv_value = 'AmountTime_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 4 iv_value = 'AmountValue_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 5 iv_value = 'AmountUnit_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 6 iv_value = 'FaultInd_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 7 iv_value = 'CalibrationInd_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 8 iv_value = 'OutOfPrecisenessOperator_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 9 iv_value = 'NotAvailableInd_3' ).
lo_sheet_3->set_cell_content( iv_row = 3 iv_column = 10 iv_value = 'Remark_3' ).

*lo_sheet_3->

lv_file = lo_doc->save( ).

CONCATENATE p_fname '.xlsx' INTO lv_filename.
TRY.


    cl_scp_change_db=>xstr_to_xtab( EXPORTING im_xstring = lv_file
                                    IMPORTING ex_xtab    = lt_file ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize              = xstrlen( lv_file )
        filename                  = |{ lv_filename }|
        filetype                  = 'BIN'
        confirm_overwrite         = abap_true
      IMPORTING
        filelength                = lv_bytes_transferred
      CHANGING
        data_tab                  = lt_file
      EXCEPTIONS
        file_write_error          = 1
        no_batch                  = 2
        gui_refuse_filetransfer   = 3
        invalid_type              = 4
        no_authority              = 5
        unknown_error             = 6
        header_not_allowed        = 7
        separator_not_allowed     = 8
        filesize_not_allowed      = 9
        header_too_long           = 10
        dp_error_create           = 11
        dp_error_send             = 12
        dp_error_write            = 13
        unknown_dp_error          = 14
        access_denied             = 15
        dp_out_of_memory          = 16
        disk_full                 = 17
        dp_timeout                = 18
        file_not_found            = 19
        dataprovider_exception    = 20
        control_flush_error       = 21
        not_supported_by_gui      = 22
        error_no_gui              = 23
        OTHERS                    = 24
    ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ELSE.
      MESSAGE s001(00) WITH lv_bytes_transferred ' bytes transferred'.
    ENDIF.

  CATCH cx_ehfnd_exp_export_err INTO DATA(lx_exception).
    WRITE: / lx_exception->mo_log_string.
ENDTRY.
