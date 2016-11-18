
"-Begin-----------------------------------------------------------------

Function zExcelViaVBScript.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     VALUE(I_INCLNAME) TYPE  SOBJ_NAME
*"----------------------------------------------------------------------

  "-Type pools----------------------------------------------------------
    Type-Pools:
      OLE2.

  "-Constants-----------------------------------------------------------
    Constants:
      CrLf(2) Type c Value cl_abap_char_utilities=>cr_lf.

  "-Variables-----------------------------------------------------------
    Data:
      oScript Type OLE2_OBJECT,
      VBCode Type String,
      VBACode Type String.

  "-Main----------------------------------------------------------------
    Create Object oScript 'MSScriptControl.ScriptControl'.
    Check sy-subrc = 0 And oScript-Handle > 0 And oScript-Type = 'OLE2'.

    "-Allow to display UI elements--------------------------------------
      Set Property Of oScript 'AllowUI' = 1.

    "-Intialize the VBScript language-----------------------------------
      Set Property Of oScript 'Language' = 'VBScript'.

    "-Code preparation for Excel VBA------------------------------------
      VBCode = 'Set oExcel = CreateObject("Excel.Application")'.
      VBCode = VBCode && CrLf.
      VBCode = VBCode && 'oExcel.Visible = True'.
      VBCode = VBCode && CrLf.
      VBCode = VBCode && 'Set oWorkbook = oExcel.Workbooks.Add()'.
      VBCode = VBCode && CrLf.
      VBCode = VBCode && 'Set oSheet = oWorkbook.ActiveSheet'.
      VBCode = VBCode && CrLf.
      VBCode = VBCode && 'With oExcel'.
      VBCode = VBCode && CrLf.

      "-Add VBA code----------------------------------------------------
        Call Function 'ZREADVBAINCLASSTRING'
          Exporting
            I_INCLNAME = I_INCLNAME
          Importing
            E_STRINCL  = VBACode.
        VBCode = VBCode && VBACode.

      VBCode = VBCode && 'End With'.
      VBCode = VBCode && CrLf.

    "-Execute VBScript code---------------------------------------------
      Call Method Of oScript 'ExecuteStatement' Exporting #1 = VBCode.

    "-Free the object---------------------------------------------------
      Free Object oScript.

EndFunction.

"-End-------------------------------------------------------------------
