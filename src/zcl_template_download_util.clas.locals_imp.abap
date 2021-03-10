*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations


    CLASS lcl_function_module IMPLEMENTATION.

      METHOD lif_function_module~number_get_next.

        CALL FUNCTION 'NUMBER_GET_NEXT'
          EXPORTING
            nr_range_nr             = nr_range_nr
            object                  = object
          IMPORTING
            number                  = number
            returncode              = returncode
          EXCEPTIONS
            interval_not_found      = 1
            number_range_not_intern = 2
            object_not_found        = 3
            quantity_is_0           = 4
            quantity_is_not_1       = 5
            interval_overflow       = 6
            buffer_overflow         = 7
            OTHERS                  = 8.

        CASE sy-subrc.
          WHEN 1.
            RAISE interval_not_found.
          WHEN 2.
            RAISE number_range_not_intern.
          WHEN 3.
            RAISE object_not_found.
          WHEN 4.
            RAISE quantity_is_0.
          WHEN 5.
            RAISE quantity_is_not_1.
          WHEN 6.
            RAISE interval_overflow.
          WHEN 7.
            RAISE buffer_overflow.
        ENDCASE.

      ENDMETHOD.

      METHOD lif_function_module~scms_string_to_xstring.

        CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
          EXPORTING
            text     = text
            mimetype = mimetype
          IMPORTING
            buffer   = buffer.

      ENDMETHOD.

    ENDCLASS.


    CLASS ltd_function_module IMPLEMENTATION.

      METHOD lif_function_module~number_get_next.
        number = '1234567890'.
        returncode = abap_true.
      ENDMETHOD.

      METHOD lif_function_module~scms_string_to_xstring.
        "do nothing
      ENDMETHOD.

    ENDCLASS.

    CLASS lcl_mpa_template_download_util IMPLEMENTATION.

      METHOD lif_mpa_template_download_util~set_parameter .

        DATA: lo_request         TYPE REF TO /iwbep/cl_mgw_request,
              lo_request_context TYPE /iwbep/if_mgw_core_srv_runtime=>ty_s_mgw_request_context,
              lt_filter_select   TYPE /iwbep/t_mgw_select_option,
              lt_select_option   TYPE /iwbep/t_cod_select_options,
              lt_param_tab       TYPE /iwbep/t_mgw_name_value_pair,
              ls_param_tab       TYPE /iwbep/s_mgw_name_value_pair.

        CLEAR ct_param_tab.

** Up-cast for getting filter values
        lo_request         ?= io_tech_request_context.
        lo_request_context =  lo_request->get_request_details( ).
        lt_filter_select   =  lo_request_context-filter_select_options.

*    lo_filter                = io_tech_request_context->get_filter( ).
*    lt_filter_select_options = lo_filter->get_filter_select_options( ).
*    lv_filter_str            = lo_filter->get_filter_string( ).

        LOOP AT lt_filter_select INTO DATA(ls_filter_select).
          lt_select_option  = ls_filter_select-select_options.
          ls_param_tab-name = ls_filter_select-property.

          LOOP AT lt_select_option INTO DATA(ls_select_option).
            ls_param_tab-value = ls_select_option-low.
            APPEND ls_param_tab TO lt_param_tab.
          ENDLOOP.
        ENDLOOP.

** Append to the param tab
        APPEND LINES OF it_key_tab   TO ct_param_tab.
        APPEND LINES OF lt_param_tab TO ct_param_tab.

      ENDMETHOD.

    ENDCLASS.


    CLASS lcl_mpa_tmplt_dl_util_mock IMPLEMENTATION.

      METHOD lif_mpa_template_download_util~set_parameter .
        APPEND LINES OF it_key_tab   TO ct_param_tab.
      ENDMETHOD.

    ENDCLASS.
