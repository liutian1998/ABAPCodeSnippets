
"-----------------------------------------------------------------------
"-
"- abap2xlsx formatter class to set special formats for a range of
"- fields in an Excel table.
"-
"-----------------------------------------------------------------------

CLASS zcl_xlsx_format DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS constructor
      IMPORTING
        !i_rcl_excel TYPE REF TO zcl_excel
        VALUE(i_active_worksheet) TYPE zexcel_active_worksheet OPTIONAL .

    METHODS set_border_outline_range
      IMPORTING
        VALUE(i_col_start)    TYPE string DEFAULT 'A'
        VALUE(i_row_start)    TYPE zexcel_cell_row DEFAULT 1
        VALUE(i_col_end)      TYPE string OPTIONAL
        VALUE(i_row_end)      TYPE zexcel_cell_row OPTIONAL
        VALUE(i_border_style) TYPE zexcel_border DEFAULT zcl_excel_style_border=>c_border_thin
        VALUE(i_border_color) TYPE zexcel_s_style_color-rgb DEFAULT zcl_excel_style_color=>c_black .

    METHODS set_border_range
      IMPORTING
        VALUE(i_col_start)    TYPE string DEFAULT 'A'
        VALUE(i_row_start)    TYPE zexcel_cell_row DEFAULT 1
        VALUE(i_col_end)      TYPE string OPTIONAL
        VALUE(i_row_end)      TYPE zexcel_cell_row OPTIONAL
        VALUE(i_border_style) TYPE zexcel_border DEFAULT zcl_excel_style_border=>c_border_thin
        VALUE(i_border_color) TYPE zexcel_s_style_color-rgb DEFAULT zcl_excel_style_color=>c_black .

    METHODS set_bgcolor_range
      IMPORTING
        VALUE(i_col_start) TYPE string DEFAULT 'A'
        VALUE(i_row_start) TYPE zexcel_cell_row DEFAULT 1
        VALUE(i_col_end)   TYPE string OPTIONAL
        VALUE(i_row_end)   TYPE zexcel_cell_row OPTIONAL
        VALUE(i_bg_color)  TYPE zexcel_s_style_color-rgb DEFAULT zcl_excel_style_color=>c_white .

    METHODS set_fontcolor_range
      IMPORTING
        VALUE(i_col_start)  TYPE string DEFAULT 'A'
        VALUE(i_row_start)  TYPE zexcel_cell_row DEFAULT 1
        VALUE(i_col_end)    TYPE string OPTIONAL
        VALUE(i_row_end)    TYPE zexcel_cell_row OPTIONAL
        VALUE(i_text_color) TYPE zexcel_s_style_color-rgb DEFAULT zcl_excel_style_color=>c_black .

    METHODS set_fontsize_range
      IMPORTING
        VALUE(i_col_start) TYPE string DEFAULT 'A'
        VALUE(i_row_start) TYPE zexcel_cell_row DEFAULT 1
        VALUE(i_col_end)   TYPE string OPTIONAL
        VALUE(i_row_end)   TYPE zexcel_cell_row OPTIONAL
        VALUE(i_text_size) TYPE zexcel_style_font_size DEFAULT 11 .

    METHODS set_fontstyle_bold_range
      IMPORTING
        VALUE(i_col_start) TYPE string DEFAULT 'A'
        VALUE(i_row_start) TYPE zexcel_cell_row DEFAULT 1
        VALUE(i_col_end)   TYPE string OPTIONAL
        VALUE(i_row_end)   TYPE zexcel_cell_row OPTIONAL .

  PROTECTED SECTION.

    DATA m_rcl_worksheet TYPE REF TO zcl_excel_worksheet .

  PRIVATE SECTION.

ENDCLASS.



CLASS zcl_xlsx_format IMPLEMENTATION.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->CONSTRUCTOR
* +--------------------------------------------------------------------+
* | [--->] I_RCL_EXCEL          TYPE REF TO ZCL_EXCEL
* | [--->] I_ACTIVE_WORKSHEET   TYPE ZEXCEL_ACTIVE_WORKSHEET(optional)
* +---------------------------------------------------------</SIGNATURE>
  METHOD constructor.

    IF i_rcl_excel IS NOT INITIAL AND i_active_worksheet IS INITIAL.
      m_rcl_worksheet = i_rcl_excel->get_active_worksheet( ).
    ELSEIF i_rcl_excel IS NOT INITIAL AND i_active_worksheet IS NOT INITIAL.
      TRY.
          i_rcl_excel->set_active_sheet_index( i_active_worksheet = i_active_worksheet ).
        CATCH zcx_excel.
          "Your error routine here
      ENDTRY.
      m_rcl_worksheet = i_rcl_excel->get_active_worksheet( ).
    ELSE.
      "Your error routine here
    ENDIF.

  ENDMETHOD.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->SET_BGCOLOR_RANGE
* +--------------------------------------------------------------------+
* | [--->] I_COL_START          TYPE        STRING (default ='A')
* | [--->] I_ROW_START          TYPE        zexcel_cell_row (default =1)
* | [--->] I_COL_END            TYPE        STRING(optional)
* | [--->] I_ROW_END            TYPE        zexcel_cell_row(optional)
* | [--->] I_BG_COLOR           TYPE        zexcel_s_style_color-RGB
* |                                         (default =zcl_excel_style_color=>C_WHITE)
* +---------------------------------------------------------</SIGNATURE>
  METHOD set_bgcolor_range.

    "-Variables---------------------------------------------------------
    DATA lv_col_start TYPE i.
    DATA lv_row_start TYPE zexcel_cell_row.
    DATA lv_col_end TYPE i.
    DATA lv_row_end TYPE zexcel_cell_row.
    DATA lv_row TYPE i.
    DATA lv_col TYPE i.
    DATA lv_col_alpha TYPE string.
    DATA lv_fill TYPE zexcel_s_cstyle_fill.

    "-Main--------------------------------------------------------------
    TRY.

        lv_col_start =
          zcl_excel_common=>convert_column2int( i_col_start ).

        lv_row_start = i_row_start.

        IF i_col_end IS INITIAL.
          lv_col_end = m_rcl_worksheet->get_highest_column( ).
        ELSE.
          lv_col_end =
            zcl_excel_common=>convert_column2int( i_col_end ).
        ENDIF.

        IF i_row_end IS INITIAL.
          lv_row_end = m_rcl_worksheet->get_highest_row( ).
        ELSE.
          lv_row_end = i_row_end.
        ENDIF.

        lv_fill-filltype = zcl_excel_style_fill=>c_fill_solid.
        lv_fill-fgcolor-rgb = i_bg_color.

        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          lv_col = lv_col_start.
          WHILE lv_col <= lv_col_end.
            lv_col_alpha =
              /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col ).
            m_rcl_worksheet->change_cell_style(
              EXPORTING
                ip_column = lv_col_alpha
                ip_row = lv_row
                ip_fill = lv_fill ).
            lv_col = lv_col + 1.
          ENDWHILE.
          lv_row = lv_row + 1.
        ENDWHILE.

      CATCH cx_root.
        "Your error routine here

    ENDTRY.

  ENDMETHOD.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->SET_BORDER_OUTLINE_RANGE
* +--------------------------------------------------------------------+
* | [--->] I_COL_START          TYPE        STRING (default ='A')
* | [--->] I_ROW_START          TYPE        zexcel_cell_row (default =1)
* | [--->] I_COL_END            TYPE        STRING(optional)
* | [--->] I_ROW_END            TYPE        zexcel_cell_row(optional)
* | [--->] I_BORDER_STYLE       TYPE        zexcel_border
* |                                         (default =zcl_excel_style_border=>C_BORDER_THIN)
* | [--->] I_BORDER_COLOR       TYPE        zexcel_s_style_color-RGB
* |                                         (default =zcl_excel_style_color=>C_BLACK)
* +---------------------------------------------------------</SIGNATURE>
  METHOD set_border_outline_range.

    "-Variables---------------------------------------------------------
    DATA lv_col_start TYPE i.
    DATA lv_row_start TYPE zexcel_cell_row.
    DATA lv_col_end TYPE i.
    DATA lv_row_end TYPE zexcel_cell_row.
    DATA lv_row TYPE i.
    DATA lv_col TYPE i.

    "-Main--------------------------------------------------------------
    TRY.

        lv_col_start =
          zcl_excel_common=>convert_column2int( i_col_start ).

        lv_row_start = i_row_start.

        IF i_col_end IS INITIAL.
          lv_col_end = m_rcl_worksheet->get_highest_column( ).
        ELSE.
          lv_col_end =
            zcl_excel_common=>convert_column2int( i_col_end ).
        ENDIF.

        IF i_row_end IS INITIAL.
          lv_row_end = m_rcl_worksheet->get_highest_row( ).
        ELSE.
          lv_row_end = i_row_end.
        ENDIF.

        "-Left----------------------------------------------------------
        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          m_rcl_worksheet->change_cell_style(
            EXPORTING
              ip_column = /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col_start )
              ip_row = lv_row
              ip_borders_left_style = i_border_style
              ip_borders_left_color_rgb = i_border_color ).
          lv_row = lv_row + 1.
        ENDWHILE.

        "-Right---------------------------------------------------------
        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          m_rcl_worksheet->change_cell_style(
            EXPORTING
              ip_column = /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col_end )
              ip_row = lv_row
              ip_borders_right_style = i_border_style
              ip_borders_right_color_rgb = i_border_color ).
          lv_row = lv_row + 1.
        ENDWHILE.

        "-Top-----------------------------------------------------------
        lv_col = lv_col_start.
        WHILE lv_col <= lv_col_end.
          m_rcl_worksheet->change_cell_style(
            EXPORTING
              ip_column = lv_col
              ip_row = lv_row_start
              ip_borders_top_style = i_border_style
              ip_borders_top_color_rgb = i_border_color ).
          lv_col = lv_col + 1.
        ENDWHILE.

        "-Bottom--------------------------------------------------------
        lv_col = lv_col_start.
        WHILE lv_col <= lv_col_end.
          m_rcl_worksheet->change_cell_style(
            EXPORTING
              ip_column = lv_col
              ip_row = lv_row_end
              ip_borders_down_style = i_border_style
              ip_borders_down_color_rgb = i_border_color ).
          lv_col = lv_col + 1.
        ENDWHILE.

      CATCH cx_root.
        "Your error routine here

    ENDTRY.

  ENDMETHOD.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->SET_BORDER_RANGE
* +--------------------------------------------------------------------+
* | [--->] I_COL_START          TYPE        STRING (default ='A')
* | [--->] I_ROW_START          TYPE        zexcel_cell_row (default =1)
* | [--->] I_COL_END            TYPE        STRING(optional)
* | [--->] I_ROW_END            TYPE        zexcel_cell_row(optional)
* | [--->] I_BORDER_STYLE       TYPE        zexcel_border
* |                                         (default =zcl_excel_style_border=>C_BORDER_THIN)
* | [--->] I_BORDER_COLOR       TYPE        zexcel_s_style_color-RGB
* |                                         (default =zcl_excel_style_color=>C_BLACK)
* +---------------------------------------------------------</SIGNATURE>
  METHOD set_border_range.

    "-Variables---------------------------------------------------------
    DATA lv_col_start TYPE i.
    DATA lv_row_start TYPE zexcel_cell_row.
    DATA lv_col_end TYPE i.
    DATA lv_row_end TYPE zexcel_cell_row.
    DATA lv_row TYPE i.
    DATA lv_col TYPE i.
    DATA lv_col_alpha TYPE string.
    DATA lv_borders TYPE zexcel_s_cstyle_borders.

    "-Main--------------------------------------------------------------
    TRY.

        lv_col_start =
          zcl_excel_common=>convert_column2int( i_col_start ).

        lv_row_start = i_row_start.

        IF i_col_end IS INITIAL.
          lv_col_end = m_rcl_worksheet->get_highest_column( ).
        ELSE.
          lv_col_end =
            zcl_excel_common=>convert_column2int( i_col_end ).
        ENDIF.

        IF i_row_end IS INITIAL.
          lv_row_end = m_rcl_worksheet->get_highest_row( ).
        ELSE.
          lv_row_end = i_row_end.
        ENDIF.

        lv_borders-left-border_style = i_border_style.
        lv_borders-left-border_color-rgb = i_border_color.
        lv_borders-right-border_style = i_border_style.
        lv_borders-right-border_color-rgb = i_border_color.
        lv_borders-top-border_style = i_border_style.
        lv_borders-top-border_color-rgb = i_border_color.
        lv_borders-down-border_style = i_border_style.
        lv_borders-down-border_color-rgb = i_border_color.

        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          lv_col = lv_col_start.
          WHILE lv_col <= lv_col_end.
            lv_col_alpha =
              /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col ).
            m_rcl_worksheet->change_cell_style(
              EXPORTING
                ip_column = lv_col_alpha
                ip_row = lv_row
                ip_borders = lv_borders ).
            lv_col = lv_col + 1.
          ENDWHILE.
          lv_row = lv_row + 1.
        ENDWHILE.

      CATCH cx_root.
        "Your error routine here

    ENDTRY.

  ENDMETHOD.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->SET_FONTCOLOR_RANGE
* +--------------------------------------------------------------------+
* | [--->] I_COL_START          TYPE        STRING (default ='A')
* | [--->] I_ROW_START          TYPE        zexcel_cell_row (default =1)
* | [--->] I_COL_END            TYPE        STRING(optional)
* | [--->] I_ROW_END            TYPE        zexcel_cell_row(optional)
* | [--->] I_TEXT_COLOR         TYPE        zexcel_s_style_color-RGB
* |                                         (default =zcl_excel_style_color=>C_BLACK)
* +---------------------------------------------------------</SIGNATURE>
  METHOD set_fontcolor_range.

    "-Variables---------------------------------------------------------
    DATA lv_col_start TYPE i.
    DATA lv_row_start TYPE zexcel_cell_row.
    DATA lv_col_end TYPE i.
    DATA lv_row_end TYPE zexcel_cell_row.
    DATA lv_row TYPE i.
    DATA lv_col TYPE i.
    DATA lv_col_alpha TYPE string.
    DATA lv_fonts TYPE zexcel_s_cstyle_font.

    "-Main--------------------------------------------------------------
    TRY.

        lv_col_start =
          zcl_excel_common=>convert_column2int( i_col_start ).

        lv_row_start = i_row_start.

        IF i_col_end IS INITIAL.
          lv_col_end = m_rcl_worksheet->get_highest_column( ).
        ELSE.
          lv_col_end =
            zcl_excel_common=>convert_column2int( i_col_end ).
        ENDIF.

        IF i_row_end IS INITIAL.
          lv_row_end = m_rcl_worksheet->get_highest_row( ).
        ELSE.
          lv_row_end = i_row_end.
        ENDIF.

        lv_fonts-color-rgb = i_text_color.

        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          lv_col = lv_col_start.
          WHILE lv_col <= lv_col_end.
            lv_col_alpha =
              /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col ).
            m_rcl_worksheet->change_cell_style(
              EXPORTING
                ip_column = lv_col_alpha
                ip_row = lv_row
                ip_font = lv_fonts ).
            lv_col = lv_col + 1.
          ENDWHILE.
          lv_row = lv_row + 1.
        ENDWHILE.

      CATCH cx_root.
        "Your error routine here

    ENDTRY.

  ENDMETHOD.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->SET_FONTSIZE_RANGE
* +--------------------------------------------------------------------+
* | [--->] I_COL_START          TYPE        STRING (default ='A')
* | [--->] I_ROW_START          TYPE        zexcel_cell_row (default =1)
* | [--->] I_COL_END            TYPE        STRING(optional)
* | [--->] I_ROW_END            TYPE        zexcel_cell_row(optional)
* | [--->] I_TEXT_SIZE          TYPE        zexcel_style_font_size
* |                                         (default =11)
* +---------------------------------------------------------</SIGNATURE>
  METHOD set_fontsize_range.

    "-Variables---------------------------------------------------------
    DATA lv_col_start TYPE i.
    DATA lv_row_start TYPE zexcel_cell_row.
    DATA lv_col_end TYPE i.
    DATA lv_row_end TYPE zexcel_cell_row.
    DATA lv_row TYPE i.
    DATA lv_col TYPE i.
    DATA lv_col_alpha TYPE string.
    DATA lv_fonts TYPE zexcel_s_cstyle_font.

    "-Main--------------------------------------------------------------
    TRY.

        lv_col_start =
          zcl_excel_common=>convert_column2int( i_col_start ).

        lv_row_start = i_row_start.

        IF i_col_end IS INITIAL.
          lv_col_end = m_rcl_worksheet->get_highest_column( ).
        ELSE.
          lv_col_end =
            zcl_excel_common=>convert_column2int( i_col_end ).
        ENDIF.

        IF i_row_end IS INITIAL.
          lv_row_end = m_rcl_worksheet->get_highest_row( ).
        ELSE.
          lv_row_end = i_row_end.
        ENDIF.

        lv_fonts-size = i_text_size.

        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          lv_col = lv_col_start.
          WHILE lv_col <= lv_col_end.
            lv_col_alpha =
              /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col ).
            m_rcl_worksheet->change_cell_style(
              EXPORTING
                ip_column = lv_col_alpha
                ip_row = lv_row
                ip_font = lv_fonts ).
            lv_col = lv_col + 1.
          ENDWHILE.
          lv_row = lv_row + 1.
        ENDWHILE.

      CATCH cx_root.
        "Your error routine here

    ENDTRY.

  ENDMETHOD.


* <SIGNATURE>----------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_FORMAT->SET_FONTSTYLE_BOLD_RANGE
* +--------------------------------------------------------------------+
* | [--->] I_COL_START          TYPE        STRING (default ='A')
* | [--->] I_ROW_START          TYPE        zexcel_cell_row (default =1)
* | [--->] I_COL_END            TYPE        STRING(optional)
* | [--->] I_ROW_END            TYPE        zexcel_cell_row(optional)
* +---------------------------------------------------------</SIGNATURE>
  METHOD set_fontstyle_bold_range.

    "-Variables---------------------------------------------------------
    DATA lv_col_start TYPE i.
    DATA lv_row_start TYPE zexcel_cell_row.
    DATA lv_col_end TYPE i.
    DATA lv_row_end TYPE zexcel_cell_row.
    DATA lv_row TYPE i.
    DATA lv_col TYPE i.
    DATA lv_col_alpha TYPE string.
    DATA lv_fonts TYPE zexcel_s_cstyle_font.

    "-Main--------------------------------------------------------------
    TRY.

        lv_col_start =
          zcl_excel_common=>convert_column2int( i_col_start ).

        lv_row_start = i_row_start.

        IF i_col_end IS INITIAL.
          lv_col_end = m_rcl_worksheet->get_highest_column( ).
        ELSE.
          lv_col_end =
            zcl_excel_common=>convert_column2int( i_col_end ).
        ENDIF.

        IF i_row_end IS INITIAL.
          lv_row_end = m_rcl_worksheet->get_highest_row( ).
        ELSE.
          lv_row_end = i_row_end.
        ENDIF.

        lv_fonts-bold = abap_true.

        lv_row = lv_row_start.
        WHILE lv_row <= lv_row_end.
          lv_col = lv_col_start.
          WHILE lv_col <= lv_col_end.
            lv_col_alpha =
              /gkv/ca03_cl_common=>convert_column2alpha( ip_column = lv_col ).
            m_rcl_worksheet->change_cell_style(
              EXPORTING
                ip_column = lv_col_alpha
                ip_row = lv_row
                ip_font = lv_fonts ).
            lv_col = lv_col + 1.
          ENDWHILE.
          lv_row = lv_row + 1.
        ENDWHILE.

      CATCH cx_root.
        "Your error routine here

    ENDTRY.

  ENDMETHOD.

ENDCLASS.
