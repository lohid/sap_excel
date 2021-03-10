*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section

  types:
    begin of gty_s_field_name_mapping,
      position   type i,
      name       type string,
      label      type string,
      f_label    type string,
      mandatory  type abap_bool,
      stru_name  type char30,
      stru_type  type char30,
      length     type n length 4,
      data_type  type c length 4,
      edm_type   type char30,
      label_type type char30,
      prefix     type string,
      values     type string_table,
    end of gty_s_field_name_mapping .
  types:
    gty_t_field_name_mappings type table of gty_s_field_name_mapping .

  interface lif_mpa_template_download_util.

    "! Set all the request parameter from UI to the global parameter table
    methods set_parameter
      importing
        !io_tech_request_context type ref to /iwbep/if_mgw_req_entity
        !it_key_tab              type /iwbep/t_mgw_name_value_pair
      changing
        ct_param_tab             type /iwbep/t_mgw_name_value_pair.

  endinterface.

  class lcl_mpa_template_download_util definition.
    public section.
      interfaces lif_mpa_template_download_util.

  endclass.

  class lcl_mpa_tmplt_dl_util_mock definition.

    public section.
      interfaces : lif_mpa_template_download_util.

  endclass.


    INTERFACE lif_function_module.
      METHODS:
        scms_string_to_xstring
          IMPORTING text     TYPE string
                    mimetype TYPE c DEFAULT space
          EXPORTING buffer   TYPE xstring,
        number_get_next
          IMPORTING  nr_range_nr TYPE nrnr
                     object      TYPE nrobj
          EXPORTING  number      TYPE any
                     returncode  TYPE nrreturn
          EXCEPTIONS interval_not_found
                     number_range_not_intern
                     object_not_found
                     quantity_is_0
                     quantity_is_not_1
                     interval_overflow
                     buffer_overflow.
    ENDINTERFACE.


    CLASS lcl_function_module DEFINITION.
      PUBLIC SECTION.
        INTERFACES lif_function_module.
    ENDCLASS.


    CLASS ltd_function_module DEFINITION.
      PUBLIC SECTION.
        INTERFACES lif_function_module.
    ENDCLASS.
