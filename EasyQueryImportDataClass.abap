
"-----------------------------------------------------------------------
"-
"- ABAP class to import data from an Easy Query
"-
"-----------------------------------------------------------------------

"! <p>ABAP Class to get data from an Easy Query</p>
CLASS z_cl_eq_data_import DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES: BEGIN OF str_selopt,
             param_name TYPE rs38l_par_,
             sign       TYPE char1,
             option     TYPE char2,
             low        TYPE char24,
             high       TYPE char24,
           END OF str_selopt.

    TYPES tab_selopt TYPE STANDARD TABLE OF str_selopt.
    TYPES tab_funcmod_params TYPE STANDARD TABLE OF rs38l_par_.

    METHODS get_eq_data
      IMPORTING
        VALUE(iv_query)        TYPE rszcompid
        VALUE(it_selopt)       TYPE tab_selopt
      EXPORTING
        et_grid_data           TYPE REF TO data
        VALUE(et_message_log)  TYPE bapirettab
        VALUE(et_column_descr) TYPE rseq_t_column_description .

    METHODS get_eq_grid_data_type
      IMPORTING
        VALUE(iv_query)     TYPE rszcompid
      RETURNING
        VALUE(ev_data_type) TYPE string .

  PROTECTED SECTION.

  PRIVATE SECTION.

    METHODS get_funcmod_to_eq
      IMPORTING
        VALUE(iv_query)    TYPE rszcompid
      RETURNING
        VALUE(ev_funcname) TYPE funcname .

    METHODS get_funcmod_imprt_params_to_eq
      IMPORTING
        VALUE(iv_query)                TYPE rszcompid
      EXPORTING
        VALUE(et_importing_param_name) TYPE tab_funcmod_params.

ENDCLASS.



CLASS z_cl_eq_data_import IMPLEMENTATION.


  METHOD get_eq_data.
************************************************************************
* Gets data from an Easy Query
************************************************************************

    FIELD-SYMBOLS <eq_grid_data> TYPE STANDARD TABLE.
    FIELD-SYMBOLS <func_import_paramname> TYPE rseq_s_select_option.

    DATA:
      lr_selopt                TYPE REF TO data,
      lv_funcname              TYPE funcname,
      lv_datatype              TYPE string,
      lr_eq_grid_data          TYPE REF TO data,
      lt_col_descr             TYPE STANDARD TABLE OF rseq_s_column_description,
      lt_row_descr             TYPE STANDARD TABLE OF rseq_s_row_description,
      lv_bapi_ret              TYPE bapiret2,
      lt_bapi_ret              TYPE STANDARD TABLE OF bapiret2,
      lt_ptab                  TYPE abap_func_parmbind_tab,
      ls_ptab                  TYPE abap_func_parmbind,
      lt_exctab                TYPE abap_func_excpbind_tab,
      lt_func_import_paramname TYPE STANDARD TABLE OF rs38l_par_,
      lv_func_import_paramname TYPE rs38l_par_,
      lv_selopt                TYPE str_selopt,
      lr_cx_root               TYPE REF TO cx_root.

    lv_datatype = me->get_eq_grid_data_type( iv_query ).
    IF lv_datatype IS INITIAL.
      lv_bapi_ret-type = 'E'.
      lv_bapi_ret-message = `Es konnte kein Datentyp zum Query` &&
        iv_query && ` ermittelt werden.`.
      APPEND lv_bapi_ret TO et_message_log.
      RETURN.
    ENDIF.

    lv_funcname = me->get_funcmod_to_eq( iv_query ).

    CREATE DATA lr_eq_grid_data TYPE TABLE OF (lv_datatype).
    ASSIGN lr_eq_grid_data->* TO <eq_grid_data>.
    CHECK <eq_grid_data> IS ASSIGNED.

    me->get_funcmod_imprt_params_to_eq( EXPORTING iv_query = iv_query
      IMPORTING et_importing_param_name = lt_func_import_paramname ).
    CHECK lt_func_import_paramname IS NOT INITIAL.
    LOOP AT lt_func_import_paramname INTO lv_func_import_paramname.
      LOOP AT it_selopt INTO lv_selopt.
        IF lv_func_import_paramname CS lv_selopt-param_name.
          ls_ptab-name = lv_func_import_paramname.
          ls_ptab-kind = abap_func_exporting.
          CREATE DATA lr_selopt TYPE rseq_s_select_option.
          ASSIGN lr_selopt->* TO <func_import_paramname>.
          <func_import_paramname>-sign = lv_selopt-sign.
          <func_import_paramname>-option = lv_selopt-option.
          <func_import_paramname>-low = lv_selopt-low.
          <func_import_paramname>-high = lv_selopt-high.
          GET REFERENCE OF <func_import_paramname> INTO ls_ptab-value.
          INSERT ls_ptab INTO TABLE lt_ptab.
          UNASSIGN <func_import_paramname>.
        ENDIF.
      ENDLOOP.
    ENDLOOP.

    ls_ptab-name = 'E_T_GRID_DATA'.
    ls_ptab-kind = abap_func_tables.
    GET REFERENCE OF <eq_grid_data> INTO ls_ptab-value.
    INSERT ls_ptab INTO TABLE lt_ptab.

    ls_ptab-name = 'E_T_COLUMN_DESCRIPTION'.
    ls_ptab-kind = abap_func_tables.
    GET REFERENCE OF lt_col_descr INTO ls_ptab-value.
    INSERT ls_ptab INTO TABLE lt_ptab.

    ls_ptab-name = 'E_T_ROW_DESCRIPTION'.
    ls_ptab-kind = abap_func_tables.
    GET REFERENCE OF lt_row_descr INTO ls_ptab-value.
    INSERT ls_ptab INTO TABLE lt_ptab.

    ls_ptab-name = 'E_T_MESSAGE_LOG'.
    ls_ptab-kind = abap_func_tables.
    GET REFERENCE OF lt_bapi_ret INTO ls_ptab-value.
    INSERT ls_ptab INTO TABLE lt_ptab.

    TRY.
        CALL FUNCTION lv_funcname
          PARAMETER-TABLE lt_ptab
          EXCEPTION-TABLE lt_exctab.
        IF sy-subrc = 0.
          et_grid_data = lr_eq_grid_data.
          et_message_log = lt_bapi_ret.
          et_column_descr = lt_col_descr.
        ELSE.
          lv_bapi_ret-type = 'E'.
          lv_bapi_ret-message = `Ein Fehler ist aufgetreten in ` &&
            lv_funcname && ` im ` && sy-repid.
          APPEND lv_bapi_ret TO et_message_log.
        ENDIF.
      CATCH cx_root INTO lr_cx_root.
        lv_bapi_ret-type = 'E'.
        lv_bapi_ret-message = lr_cx_root->get_text( ).
        APPEND lv_bapi_ret TO et_message_log.
    ENDTRY.

  ENDMETHOD.


  METHOD get_eq_grid_data_type.
************************************************************************
* Gets the datatype from E_T_GRID_DATA from the function module
* of an Easy Query
************************************************************************

    DATA lv_funcname TYPE eu_lname.
    DATA lv_funcintf TYPE rsfbintfv.
    DATA lv_r3state TYPE r3state.
    DATA lt_funcpara TYPE STANDARD TABLE OF rsfbpara.
    DATA lv_funcpara TYPE rsfbpara.
    DATA lv_typename TYPE string.
    DATA lv_pos TYPE i.
    DATA table_descr TYPE REF TO cl_abap_tabledescr.
    DATA struct_descr TYPE REF TO cl_abap_structdescr.

    lv_funcname = me->get_funcmod_to_eq( iv_query ).
    CHECK lv_funcname IS NOT INITIAL.

    cl_fb_function_utility=>meth_get_interface(
      EXPORTING
        im_name             = lv_funcname
      IMPORTING
        ex_interface        = lv_funcintf
        ex_readed_state     = lv_r3state
      EXCEPTIONS
        error_occured       = 1
        object_not_existing = 2
        OTHERS              = 3
    ).
    CHECK sy-subrc = 0.

    lt_funcpara = lv_funcintf-tables.
    LOOP AT lt_funcpara INTO lv_funcpara.
      IF lv_funcpara-parameter = 'E_T_GRID_DATA'.
        table_descr ?= cl_abap_typedescr=>describe_by_name( lv_funcpara-structure ).
        struct_descr ?= table_descr->get_table_line_type( ).
        lv_typename = struct_descr->absolute_name.
        lv_pos = strlen( lv_typename ) - 6.
        MOVE lv_typename+6(lv_pos) TO ev_data_type.
        EXIT.
      ENDIF.
    ENDLOOP.


  ENDMETHOD.


  METHOD get_funcmod_to_eq.
************************************************************************
* Gets the name of the function module from an Easy Query
************************************************************************

    CALL FUNCTION 'RSEQ_GET_EQRFC_NAME'
      EXPORTING
        i_query   = iv_query
      IMPORTING
        e_rfcname = ev_funcname.

  ENDMETHOD.


  METHOD get_funcmod_imprt_params_to_eq.
************************************************************************
* Gets the names of the import variables of the function module
* from an Easy Query
************************************************************************

    DATA lv_funcname TYPE eu_lname.
    DATA lv_funcintf TYPE rsfbintfv.
    DATA lv_r3state TYPE r3state.
    DATA lt_funcpara TYPE STANDARD TABLE OF rsfbpara.
    DATA lv_funcpara TYPE rsfbpara.

    lv_funcname = me->get_funcmod_to_eq( iv_query ).
    CHECK lv_funcname IS NOT INITIAL.

    cl_fb_function_utility=>meth_get_interface(
      EXPORTING
        im_name             = lv_funcname
      IMPORTING
        ex_interface        = lv_funcintf
        ex_readed_state     = lv_r3state
      EXCEPTIONS
        error_occured       = 1
        object_not_existing = 2
        OTHERS              = 3
    ).

    CHECK sy-subrc = 0.

    lt_funcpara = lv_funcintf-import.
    LOOP AT lt_funcpara INTO lv_funcpara.
      APPEND lv_funcpara-parameter TO et_importing_param_name.
    ENDLOOP.

  ENDMETHOD.


ENDCLASS.


*-An exmaple how to use it----------------------------------------------
*  FIELD-SYMBOLS <eq_grid_data> TYPE STANDARD TABLE.
*  DATA lo_eq_imp TYPE REF TO z_cl_eq_data_import.
*  DATA lt_selopt TYPE STANDARD TABLE OF str_selopt.
*  DATA lr_eq_grid_data TYPE REF TO data.
*  DATA lt_bapiret TYPE bapirettab.
*  DATA lt_headline TYPE STANDARD TABLE OF rseq_s_column_description.

*  CREATE OBJECT lo_eq_imp.
*  TRY.
*      lo_eq_imp->get_eq_data(
*        EXPORTING
*          iv_query = 'NameOfEasyQuery'
*          it_selopt = lt_selopt
*        IMPORTING
*          et_grid_data   = lr_eq_grid_data
*          et_message_log = lt_bapiret
*          et_column_descr = lt_headline
*      ).
*    CATCH cx_root.
*      "Your error routine here
*  ENDTRY.
*  ASSIGN lr_eq_grid_data->* TO <eq_grid_data>.