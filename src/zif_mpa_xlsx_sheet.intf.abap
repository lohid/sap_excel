interface ZIF_MPA_XLSX_SHEET
  public .
    TYPES: celltype        TYPE c length 1.
    TYPES: cell_value_type TYPE string.
    CONSTANTS:
      BEGIN OF gc_celltype,
        charlike    TYPE celltype VALUE cl_abap_typedescr=>typekind_clike,
        date        TYPE celltype VALUE cl_abap_typedescr=>typekind_date,
        time        TYPE celltype VALUE cl_abap_typedescr=>typekind_time,
        timestamp   TYPE celltype VALUE 'U',
        integer     TYPE celltype VALUE cl_abap_typedescr=>typekind_int,
        float       TYPE celltype VALUE cl_abap_typedescr=>typekind_decfloat,
        numericchar TYPE celltype VALUE cl_abap_typedescr=>typekind_num,
        other       TYPE celltype VALUE abap_undefined,
      END OF gc_celltype.
    CONSTANTS:
      BEGIN OF gc_cell_value_type,
          boolean       TYPE cell_value_type VALUE 'b',
          number        TYPE cell_value_type VALUE 'n',
          error         TYPE cell_value_type VALUE 'e',
          shared_string TYPE cell_value_type VALUE 's',
          string        TYPE cell_value_type VALUE 'str',
          inline_string TYPE cell_value_type VALUE 'inlineStr',
          date          TYPE cell_value_type VALUE 'd',
          undefined     TYPE cell_value_type VALUE '',
      END OF gc_cell_value_type.




    METHODS: set_cell_content IMPORTING iv_row          TYPE i
                                        iv_column       TYPE i
                                        iv_value        TYPE any
                                        iv_input_type   TYPE celltype DEFAULT gc_celltype-charlike
                                        iv_force_string TYPE abap_bool DEFAULT abap_false.
    METHODS: has_cell_content IMPORTING iv_row    TYPE i
                                        iv_column TYPE i
                              RETURNING VALUE(rv_has_content) TYPE abap_bool.
    METHODS: get_cell_content IMPORTING iv_row    TYPE i
                                        iv_column TYPE i
                              RETURNING VALUE(rv_content) TYPE string.
    METHODS: get_last_row_number RETURNING VALUE(rv_row) TYPE i.
    METHODS: get_last_column_number_in_row IMPORTING iv_row TYPE i
                                           RETURNING VALUE(rv_column) TYPE i.
    METHODS: change_sheet_name IMPORTING iv_new_name TYPE clike.


endinterface.
