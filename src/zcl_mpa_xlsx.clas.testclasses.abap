*"* use this source file for your ABAP unit test classes
CLASS ltc_test DEFINITION FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.
  PUBLIC SECTION.
  METHODS: test_singleton FOR TESTING.

ENDCLASS.

CLASS ltc_test IMPLEMENTATION.

  METHOD test_singleton.

    DATA: lo_xlsx_1 TYPE REF TO zcl_mpa_xlsx.
    DATA: lo_xlsx_2 TYPE REF TO zcl_mpa_xlsx.

    lo_xlsx_1 = zcl_mpa_xlsx=>get_instance( ).
    lo_xlsx_2 = zcl_mpa_xlsx=>get_instance( ).

    cl_abap_unit_assert=>assert_equals( msg = 'Singleton should always return the same instance' exp = lo_xlsx_1 act = lo_xlsx_2 ).

  ENDMETHOD.

ENDCLASS.
