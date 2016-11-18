
"-Begin-----------------------------------------------------------------
  Function ZREADVBAINCLASSTRING.
*"--------------------------------------------------------------------
*"  Local Interface:
*"  IMPORTING
*"     VALUE(I_INCLNAME) TYPE  SOBJ_NAME
*"  EXPORTING
*"     VALUE(E_STRINCL) TYPE  STRING
*"--------------------------------------------------------------------

    "-Variables---------------------------------------------------------
      Data resTADIR Type TADIR.
      Data tabIncl Type Table Of String.
      Data lineIncl Type String Value ''.
      Data strIncl Type String Value ''.

    "-Main--------------------------------------------------------------
      Select Single * From TADIR Into resTADIR
        Where OBJ_NAME = I_InclName.
      If sy-subrc = 0.

        Read Report I_InclName Into tabIncl.
        If sy-subrc = 0.
          Loop At tabIncl Into lineIncl.
            Condense lineIncl.
            Concatenate strIncl '.' lineIncl
              cl_abap_char_utilities=>cr_lf Into strIncl.
            lineIncl = ''.
          EndLoop.
        EndIf.

      EndIf.
      E_strIncl = strIncl.

  EndFunction.

"-End-------------------------------------------------------------------
