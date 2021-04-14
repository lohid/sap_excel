CLASS zcl_xlsx_helper DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    CLASS-METHODS convert_date_to_xlsx_date
      IMPORTING
        !iv_date       TYPE d
      RETURNING
        VALUE(rv_date) TYPE string .
    CLASS-METHODS get_cell_position
      IMPORTING
        !iv_row            TYPE i
        !iv_column         TYPE i
      RETURNING
        VALUE(rv_position) TYPE string .
    CLASS-METHODS convert_time_to_xlsx_time
      IMPORTING
        !iv_time       TYPE t
      RETURNING
        VALUE(rv_time) TYPE string .
    CLASS-METHODS convert_timestamp_to_xlsx
      IMPORTING
        !iv_timestamp       TYPE p
      RETURNING
        VALUE(rv_timestamp) TYPE string .
    CLASS-METHODS get_cell_column
      IMPORTING
        !iv_cell_position TYPE string
      RETURNING
        VALUE(rv_column)  TYPE i .
  PROTECTED SECTION.
  PRIVATE SECTION.

    CLASS-DATA mt_column_pos_cache TYPE TABLE OF string .
ENDCLASS.



CLASS zcl_xlsx_helper IMPLEMENTATION.


  METHOD convert_date_to_xlsx_date.

**********************************************************************
* define data
**********************************************************************

    DATA: lv_date(20)               TYPE c.
    DATA: lv_i_date                 TYPE i.
    DATA: lv_i_19000100             TYPE i.
    CONSTANTS: lc_hlpdate           TYPE d VALUE '19000101'.

**********************************************************************
* functional body
**********************************************************************

    IF iv_date IS NOT INITIAL AND iv_date >= lc_hlpdate.

*   store the integer representation of the date in local variable
      lv_i_date    = lc_hlpdate.
*   XLSX calculation requires two days less
      lv_i_19000100 = lv_i_date - 2.
*   calculate the delta
      lv_i_date = iv_date - lv_i_19000100.
*   convert integer back to date (string)
      MOVE lv_i_date TO lv_date.
      CONDENSE lv_date.
*   put the converted date into the cell value
      rv_date = lv_date.
    ELSE.
      CLEAR: rv_date.
    ENDIF.

  ENDMETHOD.


  METHOD convert_timestamp_to_xlsx.

**********************************************************************
* define data
**********************************************************************

    DATA: lv_timestamp(30)            TYPE c.
    DATA: lv_date                     TYPE d.
    DATA: lv_time                     TYPE t.
    DATA: lv_datestr                  TYPE string.
    DATA: lv_timestr                  TYPE string.
    DATA: lv_empty_timezone           TYPE tznzone.

**********************************************************************
* functional body
**********************************************************************

* get the date and time out of the timestamp
* an empty timezone means UTC
    CONVERT TIME STAMP iv_timestamp TIME ZONE lv_empty_timezone INTO DATE lv_date TIME lv_time.
* convert the date
    lv_datestr = convert_date_to_xlsx_date( iv_date = lv_date ).
* convert the time
    lv_timestr = convert_time_to_xlsx_time( iv_time = lv_time ).
* shorten the time
    lv_timestr = substring( val = lv_timestr off = 1 ).
* build the timestamp
    CONCATENATE lv_datestr lv_timestr INTO lv_timestamp.
* return the timestamp
    rv_timestamp = lv_timestamp.

  ENDMETHOD.


  METHOD convert_time_to_xlsx_time.

**********************************************************************
* define data
**********************************************************************
    DATA: lv_p_day                 TYPE p DECIMALS 14.
    DATA: lv_i_hours               TYPE i.
    DATA: lv_i_minutes             TYPE i.
    DATA: lv_i_seconds             TYPE i.

**********************************************************************
* functional body
**********************************************************************

    IF iv_time IS NOT INITIAL.
*   calculate the time for xlsx
      lv_i_hours   = iv_time(2).
      lv_i_minutes = iv_time+2(2).
      lv_i_seconds = iv_time+4(2).

*   calculate the percentage of the time in the day
      lv_p_day = lv_i_hours / 24
               + lv_i_minutes / ( 24 * 60 )
               + lv_i_seconds / ( 24 * 3600 ).
*   put the percentage to the returning value
      rv_time = lv_p_day.
      CONDENSE rv_time.

    ELSE.
*    CLEAR: lv_p_day.
      rv_time = '0'.
    ENDIF.

  ENDMETHOD.


  METHOD get_cell_column.
**********************************************************************
* define data
**********************************************************************

    DATA: lv_str_length   TYPE i.
    DATA: lv_current_pos  TYPE i.
    DATA: lv_current_char TYPE c.
    CONSTANTS: lc_column  TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.

**********************************************************************
* functional body
**********************************************************************

*   Start with the first column
    rv_column = 0.

    lv_str_length = strlen( iv_cell_position ).
    WHILE lv_current_pos < lv_str_length.
      lv_current_char = iv_cell_position+lv_current_pos(1).
      TRANSLATE lv_current_char TO UPPER CASE.
      IF lc_column CA lv_current_char.
*       If the current character is one of A...Z
*       We first multiply the column by 26 for all the columns before
        rv_column = rv_column * 26.
*       then we add the offset of the Character to the column
        rv_column = rv_column + sy-fdpos + 1.
*       Then evaluate the next character
        ADD 1 TO lv_current_pos.
      ELSE.
*       The character is not in A...Z - so we can directly return.
        RETURN.
      ENDIF.

    ENDWHILE.




  ENDMETHOD.


  METHOD get_cell_position.

**********************************************************************
* define data
**********************************************************************

    DATA: lv_column           TYPE i.
    DATA: lv_tmp_column       TYPE i.
    DATA: lv_row              TYPE string.
* Begin Correction 13.12.2012 1800433 *********************************
    DATA: lv_columns_to_add   TYPE i.
    DATA: lv_next_column      TYPE i.
    DATA: lv_position         TYPE string.
* End Correction 13.12.2012 1800433 ***********************************
    CONSTANTS: lc_column      TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.

**********************************************************************
* functional body
**********************************************************************

* this method returns the calculated position of a cell (e.g. BA21)

* check import parameters
    ASSERT iv_row > 0.
    ASSERT iv_column > 0.

    CLEAR: rv_position.

* Begin Correction 13.12.2012 1800433 *********************************
* Check if the requested position is already in the cache
    lv_columns_to_add = iv_column - lines( mt_column_pos_cache ).
    IF ( lv_columns_to_add > 0 ).

*   The position is not in the cache so we have to calculate the next
*   position(s) and extend the cache. We have to calculate
*   as many positions so that the requested one will be part of the
*   cache afterwards. Thus the cache always contains a contingous
*   sequence of column positions even if the positions are not requested
*   in sequential order.

*   Determine the next missing column
      lv_next_column = lines( mt_column_pos_cache ) + 1.

      DO lv_columns_to_add TIMES.

        CLEAR lv_position.
        lv_column = lv_next_column.

*     we build the position from left to right
        WHILE lv_column > 0.
*       calculate the new position
          lv_column = lv_column - 1.
          lv_tmp_column = lv_column MOD 26.

*       check the tmp column and map it to characters:
*         0 -> A
*         1 -> B
*         2 -> C
*         3 -> D
*         ...

          CONCATENATE lc_column+lv_tmp_column(1) lv_position INTO lv_position.

          lv_column = lv_column DIV 26.
        ENDWHILE.

*     Add the new position to the cache
        APPEND lv_position TO mt_column_pos_cache.

        ADD 1 TO lv_next_column.
      ENDDO.
    ENDIF.

* Read the requested position from the cache
    READ TABLE mt_column_pos_cache INTO rv_position INDEX iv_column.
* We should always find it in the cache since the cache is extended when necessary
    ASSERT sy-subrc = 0.
* End Correction 13.12.2012 1800433 *********************************

* finally put the row number to the position string
    lv_row = iv_row.
    CONCATENATE rv_position lv_row INTO rv_position.
    CONDENSE rv_position.

  ENDMETHOD.
ENDCLASS.
