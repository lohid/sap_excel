*"* use this source file for your ABAP unit test classes

CLASS ltc_test_helper DEFINITION FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
    METHODS: test_get_cell_column FOR TESTING.
    METHODS: test_get_cell_position FOR TESTING.


ENDCLASS.

CLASS ltc_test_helper IMPLEMENTATION.

  METHOD test_get_cell_column.
    DATA: lv_column TYPE i.

    lv_column = zcl_xlsx_helper=>get_cell_column( iv_cell_position = 'A1' ).
    cl_abap_unit_assert=>assert_equals( msg = 'A1 should return 1' exp = 1 act = lv_column ).

    lv_column = zcl_xlsx_helper=>get_cell_column( iv_cell_position = 'K14' ).
    cl_abap_unit_assert=>assert_equals( msg = 'K14 should return 11' exp = 11 act = lv_column ).

    lv_column = zcl_xlsx_helper=>get_cell_column( iv_cell_position = 'AA345' ).
    cl_abap_unit_assert=>assert_equals( msg = 'AA345 should return 27' exp = 27 act = lv_column ).

    lv_column = zcl_xlsx_helper=>get_cell_column( iv_cell_position = 'BDA85' ).
    cl_abap_unit_assert=>assert_equals( msg = 'BDA85 should return 1457' exp = 1457 act = lv_column ).

    lv_column = zcl_xlsx_helper=>get_cell_column( iv_cell_position = 'a1' ).
    cl_abap_unit_assert=>assert_equals( msg = 'A1 should return 1' exp = 1 act = lv_column ).
  ENDMETHOD.

  METHOD test_get_cell_position.

    DATA: lv_position TYPE string.

    lv_position = zcl_xlsx_helper=>get_cell_position( iv_column = 1    iv_row = 1   ).
    cl_abap_unit_assert=>assert_equals( msg = '1/1 should be A1' exp = 'A1' act = lv_position ).

    lv_position = zcl_xlsx_helper=>get_cell_position( iv_column = 11    iv_row = 14   ).
    cl_abap_unit_assert=>assert_equals( msg = '11/14 should be K14' exp = 'K14' act = lv_position ).

    lv_position = zcl_xlsx_helper=>get_cell_position( iv_column = 27    iv_row = 345   ).
    cl_abap_unit_assert=>assert_equals( msg = '27/345 should be AA345' exp = 'AA345' act = lv_position ).

    lv_position = zcl_xlsx_helper=>get_cell_position( iv_column = 1457    iv_row = 85   ).
    cl_abap_unit_assert=>assert_equals( msg = '1457/85 should be BDA85' exp = 'BDA85' act = lv_position ).

  ENDMETHOD.

ENDCLASS.
