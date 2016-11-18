
"-Begin-----------------------------------------------------------------
"-
"- Function module to export any table or view, include CDS, as Excel
"- table on the frontend server via ABAP2XLSX. The IV_NAME parameter
"- expects the table or view name. The IV_WHERE parameter expects for a
"- table the SQL where clause and the IV_CDS_PARAMS the necessary
"- parameter(s) of the CDS view.
"-
"- Hint: bind_table throws an exception if the table contains a hex
"-       string field.
"-
"-----------------------------------------------------------------------
FUNCTION zExportAsExcel
  IMPORTING
    VALUE(IV_NAME) TYPE STRING
    VALUE(IV_WHERE) TYPE STRING OPTIONAL
    VALUE(IV_CDS_PARAMS) TYPE STRING OPTIONAL.

"-Variables-------------------------------------------------------------
  DATA lo_excel TYPE REF TO ZCL_EXCEL.
  DATA lo_worksheet TYPE REF TO ZCL_EXCEL_WORKSHEET.
  DATA lo_excelwriter TYPE REF TO ZIF_EXCEL_WRITER.
  DATA lv_xlsxdatastream TYPE xstring.
  DATA lt_rawdata TYPE solix_tab.
  DATA lv_bytecount TYPE i.
  DATA lv_name TYPE string.

"-Main------------------------------------------------------------------
  DATA dref TYPE REF TO data.
  CREATE DATA dref TYPE TABLE OF (iv_name).
  FIELD-SYMBOLS: <wa> TYPE ANY TABLE.
  ASSIGN dref->* TO <wa>.

  "-Get data from table, view or CDS view-------------------------------
  IF iv_cds_params <> ''.
    lv_name = iv_name && `( ` && iv_cds_params && ` )`.
    SELECT * FROM (lv_name) INTO TABLE @<wa>.
  ELSE.
    IF iv_where <> ''.
      SELECT * FROM (iv_name) INTO TABLE @<wa> WHERE (iv_where).
    ELSE.
      SELECT * FROM (iv_name) INTO TABLE @<wa>.
    ENDIF.
  ENDIF.

  "-Convert data to Excel-----------------------------------------------
  TRY.
    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->bind_table( ip_table = <wa> ).
    CREATE OBJECT lo_excelwriter TYPE ZCL_EXCEL_WRITER_2007.
    "CREATE OBJECT lo_excelwriter TYPE ZCL_EXCEL_WRITER_HUGE_FILE.
    lv_xlsxdatastream = lo_excelwriter->write_file( lo_excel ).
  CATCH cx_root.
    "Your error routine here
  ENDTRY.

  "-Convert xstring to binary-------------------------------------------
  lt_rawdata = cl_bcs_convert=>xstring_to_solix(
    iv_xstring  = lv_xlsxdatastream ).
  lv_bytecount = xstrlen( lv_xlsxdatastream ).

  "-Replace all slashes-------------------------------------------------
  lv_Name = iv_name.
  Replace All Occurrences Of '/' In lv_name With '_'.

  "-Save file on frontend server----------------------------------------
  cl_gui_frontend_services=>gui_download(
    EXPORTING
      bin_filesize = lv_bytecount
      filename     = 'C:\Users\Public\Documents\' && lv_name && '.xlsx'
      filetype     = 'BIN'
    CHANGING
      data_tab     = lt_rawdata
    EXCEPTIONS
      OTHERS = 1 ).
  IF sy-subrc <> 0.
    "Your error routine here
  ENDIF.

ENDFUNCTION.
"-End-------------------------------------------------------------------
