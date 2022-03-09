PROGRAM zconfig_management_tool.

DATA: gv_file_path  TYPE string.
DATA: go_excel     TYPE ole2_object.

TYPES : BEGIN OF ty_ddtext,
          tabname   TYPE tabname,
          fieldname	TYPE fieldname,
          ddtext    TYPE as4text,
        END OF ty_ddtext.

TYPES : BEGIN OF ty_out_txt,
          line(5000) TYPE c,
        END OF ty_out_txt.

TYPES : BEGIN OF ty_excel,
          line(262142) TYPE c,
        END OF ty_excel.

TYPES : BEGIN OF ty_header,
          tablename TYPE scpr_tabl,
          fieldname	TYPE scpr_fld,
          flag      TYPE scpr_flag,
          langu     TYPE spras,
        END OF ty_header.

TYPES: tt_ddtext  TYPE STANDARD TABLE OF ty_ddtext,
       tt_out_txt TYPE STANDARD TABLE OF ty_out_txt,
       tt_excel   TYPE STANDARD TABLE OF ty_excel,
       tt_header  TYPE STANDARD TABLE OF ty_header.

CONSTANTS: gc_tab TYPE c VALUE cl_abap_char_utilities=>horizontal_tab.

FORM set_file_path USING pv_filename TYPE string.
  CLEAR gv_file_path.
  gv_file_path = pv_filename.
ENDFORM.

FORM download_file TABLES pt_values     STRUCTURE scprvals
                          pt_valuesl    STRUCTURE scprvall
                          pt_varfldtxts STRUCTURE scprfldv
                          pt_recattr    STRUCTURE scprreca
                   USING  pv_profid     TYPE scpr_id
                          pv_proftext   TYPE scpr_text
                          pv_category   TYPE scpr_ctgry.
* Local data declaration
  DATA: lt_ddtext   TYPE tt_ddtext,
        lt_key_flds TYPE ujctrl_t_integer.

  DATA: lt_data       TYPE tt_out_txt.
  DATA: lv_header1    TYPE string,
        lv_header2    TYPE string,
        lv_line       TYPE string,
        lv_cell       TYPE syst_index,
        lv_sheet_num  TYPE syst_index,
        lv_tot_sheets TYPE text3.

* Text Table for IMG Activity
  SELECT activity,
         text
    INTO TABLE @DATA(lt_img_text)
    FROM cus_imgact
    FOR ALL ENTRIES IN @pt_recattr
    WHERE activity = @pt_recattr-activity+0(20) AND
          spras    = @sy-langu.
  IF sy-subrc <> 0.
    MESSAGE i208(00) WITH 'IMG Activity Text not found'.
    RETURN.
  ELSE.
    SORT lt_img_text BY activity.
    DESCRIBE TABLE lt_img_text LINES lv_tot_sheets.
    lv_tot_sheets = lv_tot_sheets + 1.
    CONDENSE lv_tot_sheets NO-GAPS.
  ENDIF.

* Sort tables
  SORT pt_values  BY recnumber tablename flag fieldname.
  SORT pt_valuesl BY recnumber tablename langu flag fieldname.
  SORT pt_recattr BY activity recnumber tablename.

  LOOP AT pt_values ASSIGNING FIELD-SYMBOL(<lfs_val>).
    CONDENSE <lfs_val>-recnumber NO-GAPS.
  ENDLOOP.
  LOOP AT pt_valuesl ASSIGNING FIELD-SYMBOL(<lfs_val1>).
    CONDENSE <lfs_val1>-recnumber NO-GAPS.
  ENDLOOP.
  LOOP AT pt_recattr ASSIGNING FIELD-SYMBOL(<lfs_attr>).
    CONDENSE <lfs_attr>-recnumber NO-GAPS.
  ENDLOOP.

  PERFORM get_ddtext TABLES pt_values pt_valuesl  CHANGING lt_ddtext.
  PERFORM append_line USING 'SHEETCOUNT' lv_tot_sheets '' ''
                      CHANGING lt_data.
  PERFORM add_header USING pv_profid pv_category pv_proftext
                  CHANGING lt_data.

  PERFORM add_sheet USING 'Header' '4' lt_key_flds lv_tot_sheets
                 CHANGING lt_data lv_sheet_num.

  LOOP AT pt_recattr ASSIGNING FIELD-SYMBOL(<lfs_reca>)
                     GROUP BY ( activity = <lfs_reca>-activity
                                recnumber = <lfs_reca>-recnumber ).
    lv_header1 = <lfs_reca>-activity.
    lv_header2 = 'Record Number'.
    lv_line    = <lfs_reca>-recnumber.
    CONDENSE lv_line.
    IF <lfs_reca>-recnumber = 1.
      lv_cell = 1.
    ENDIF.
    LOOP AT GROUP <lfs_reca> ASSIGNING FIELD-SYMBOL(<lfs_recax>).      " Loop all tables of a Config Node
      DATA(lv_index) = sy-tabix + 1.
      PERFORM add_values  TABLES   pt_values
                          USING    <lfs_recax>
                                   lt_ddtext
                          CHANGING lt_key_flds lv_header1 lv_header2
                                   lv_line lv_cell.

      PERFORM add_values1 TABLES   pt_valuesl
                          USING    <lfs_recax>
                                   lt_ddtext
                          CHANGING lt_key_flds lv_header1 lv_header2
                                   lv_line lv_cell.
    ENDLOOP.

    IF <lfs_reca>-recnumber = 1.
      APPEND lv_header1 TO lt_data.
      APPEND lv_header2 TO lt_data.
    ENDIF.

    APPEND lv_line TO lt_data.
    CLEAR: lv_header1, lv_header2, lv_line.

    READ TABLE pt_recattr INDEX lv_index INTO DATA(ls_temp)
                                               TRANSPORTING activity.
    IF sy-subrc <> 0 OR ls_temp-activity <> <lfs_reca>-activity.
      READ TABLE lt_img_text ASSIGNING FIELD-SYMBOL(<lfs_img_text>)
                             WITH KEY activity = <lfs_reca>-activity
                             BINARY SEARCH.
      IF sy-subrc = 0.
        PERFORM add_sheet USING <lfs_img_text>-text lv_cell lt_key_flds
                                lv_tot_sheets
                          CHANGING lt_data lv_sheet_num.
      ENDIF.
      CLEAR lt_key_flds.
    ENDIF.
    CLEAR ls_temp.
  ENDLOOP.
  PERFORM file_transfer.

  CLEAR gv_file_path.
ENDFORM.

FORM add_sheet  USING    pv_text        TYPE hier_text
                         pv_cell        TYPE syst_index
                         pt_key_flds    TYPE ujctrl_t_integer
                         pv_tot_sheets  TYPE text3 "syst_index
                CHANGING ct_data        TYPE tt_out_txt
                         cv_sheetno     TYPE syst_index.

  DATA: lo_workbooks TYPE ole2_object, " list of workbooks
        lo_workbook  TYPE ole2_object, " workbook
        lo_worksheet TYPE ole2_object,
        lo_columns   TYPE ole2_object,
        lo_cells     TYPE ole2_object,
        lo_range     TYPE ole2_object,
        lo_cell_frm  TYPE ole2_object,
        lo_cell_to   TYPE ole2_object.
  DATA: lo_interior TYPE ole2_object.
  DATA lv_rc        TYPE i.

  cv_sheetno = cv_sheetno + 1.                                         " Add new sheet
  IF cv_sheetno = 1.
    CREATE OBJECT go_excel 'EXCEL.APPLICATION'.
    SET PROPERTY OF go_excel 'Visible'  = 0.
    CALL METHOD OF go_excel 'Workbooks' = lo_workbooks.
    SET PROPERTY OF go_excel 'SheetsInNewWorkbook' = pv_tot_sheets.       " No of sheets
    CALL METHOD OF lo_workbooks 'Add' = lo_workbook.
  ENDIF.

  CALL METHOD OF go_excel 'WORKSHEETS' = lo_worksheet
    EXPORTING
     #1 = cv_sheetno.
  CALL METHOD OF lo_worksheet 'ACTIVATE'.
  SET PROPERTY OF lo_worksheet 'Name'  = pv_text+0(30).
  IF sy-subrc <> 0.
    SET PROPERTY OF lo_worksheet 'Name'  = 'Error in name' .
  ENDIF.

  CALL METHOD cl_gui_frontend_services=>clipboard_export
    IMPORTING
      data                 = ct_data[]
    CHANGING
      rc                   = lv_rc
    EXCEPTIONS
      cntl_error           = 1
      error_no_gui         = 2
      not_supported_by_gui = 3
      OTHERS               = 4.
  IF sy-subrc = 0.
*  Get the number of rows in each sheet
    DESCRIBE TABLE ct_data LINES DATA(lv_lines).

    CALL METHOD OF go_excel 'Cells' = lo_cell_frm
      EXPORTING
        #1 = 1
        #2 = 1.

    CALL METHOD OF go_excel 'Cells' = lo_cell_to
      EXPORTING
        #1 = lv_lines
        #2 = pv_cell.

    CALL METHOD OF go_excel 'Range' = lo_range
      EXPORTING
        #1 = lo_cell_frm
        #2 = lo_cell_to.

    SET PROPERTY OF lo_range 'NumberFormat' = '@' . "To disply zeros

    CALL METHOD OF lo_range 'Select'.
    CALL METHOD OF lo_worksheet 'Paste'.

    CALL METHOD OF go_excel 'Columns' = lo_columns.
    CALL METHOD OF lo_columns 'Autofit'.

    LOOP AT pt_key_flds ASSIGNING FIELD-SYMBOL(<lfs_key>).

      CALL METHOD OF go_excel 'Cells' = lo_cells
        EXPORTING
          #1 = 2
          #2 = <lfs_key>.

      GET PROPERTY OF lo_cells 'Interior' = lo_interior .
      SET PROPERTY OF lo_interior 'Color' = 255000255. "200000200         "colour code for yellow
    ENDLOOP.
    FREE OBJECT lo_columns.
  ENDIF.
  CLEAR ct_data.
ENDFORM.

FORM get_ddtext TABLES   pt_values  STRUCTURE scprvals
                         pt_values1 STRUCTURE scprvall
                CHANGING pc_ddtext TYPE tt_ddtext.

  IF pt_values IS NOT INITIAL.
* Get R/3 DD: Data element texts
    SELECT tabname
           fieldname
           ddtext
           FROM dd03m
           INTO TABLE pc_ddtext
           FOR ALL ENTRIES IN pt_values
           WHERE tabname    = pt_values-tablename AND
                 fieldname  = pt_values-fieldname AND
                 ddlanguage = sy-langu.
    IF sy-subrc = 0.
    ENDIF.
  ENDIF.
  IF pt_values1 IS NOT INITIAL.
    SELECT tabname
           fieldname
           ddtext
           FROM dd03m
           APPENDING TABLE pc_ddtext
           FOR ALL ENTRIES IN pt_values1
           WHERE tabname    = pt_values1-tablename AND
                 fieldname  = pt_values1-fieldname AND
                 ddlanguage = sy-langu.
    IF sy-subrc = 0.
    ENDIF.
  ENDIF.
  SORT pc_ddtext BY tabname fieldname.
  DELETE ADJACENT DUPLICATES FROM pc_ddtext
                             COMPARING tabname fieldname.
ENDFORM.

FORM add_values TABLES   pt_values  STRUCTURE scprvals
                USING    ps_recax   TYPE scprreca
                         pt_ddtext  TYPE tt_ddtext
                CHANGING ct_key_flds TYPE ujctrl_t_integer
                         pv_header1 TYPE string
                         pv_header2 TYPE string
                         pv_line    TYPE string
                         cv_cell    TYPE syst_index.

  READ TABLE pt_values TRANSPORTING NO FIELDS WITH KEY
                                recnumber = ps_recax-recnumber
                                tablename = ps_recax-tablename
                                BINARY SEARCH.
  IF sy-subrc = 0.
    LOOP AT pt_values ASSIGNING FIELD-SYMBOL(<lfs_values>)
                                       FROM sy-tabix.
      IF <lfs_values>-recnumber <> ps_recax-recnumber OR
         <lfs_values>-tablename <> ps_recax-tablename.
        EXIT.
      ELSEIF <lfs_values>-fieldname = 'MANDT'.
        CONTINUE.
      ENDIF.

      CONCATENATE pv_line <lfs_values>-value INTO pv_line
              SEPARATED BY gc_tab.

      IF ps_recax-recnumber = 1.
        cv_cell = cv_cell + 1.
        IF <lfs_values>-flag = 'KEY' OR <lfs_values>-flag = 'FKY'.
          APPEND cv_cell TO ct_key_flds.
        ENDIF.
        DATA(lv_temp) = <lfs_values>-tablename && '-' &&
                    <lfs_values>-fieldname && '-' && <lfs_values>-flag.

        CONCATENATE pv_header1 lv_temp INTO pv_header1
          SEPARATED BY gc_tab.

        READ TABLE pt_ddtext INTO DATA(wa_ddtext)
                             WITH KEY tabname   = <lfs_values>-tablename
                                      fieldname = <lfs_values>-fieldname
                             BINARY SEARCH.

        CONCATENATE pv_header2 wa_ddtext-ddtext INTO pv_header2
          SEPARATED BY gc_tab.

      ENDIF.
      CLEAR wa_ddtext.
    ENDLOOP.
  ENDIF.
ENDFORM.

FORM add_values1 TABLES  pt_values1 STRUCTURE scprvall
                USING    ps_recax   TYPE scprreca
                         pt_ddtext  TYPE tt_ddtext
                CHANGING ct_key_flds TYPE ujctrl_t_integer
                         pv_header1 TYPE string
                         pv_header2 TYPE string
                         pv_line    TYPE string
                         cv_cell    TYPE syst_index.
  DATA: lv_langu TYPE char02.
  READ TABLE pt_values1 TRANSPORTING NO FIELDS WITH KEY
                                recnumber = ps_recax-recnumber
                                tablename = ps_recax-tablename
                                BINARY SEARCH.
  IF sy-subrc = 0.
    LOOP AT pt_values1 ASSIGNING FIELD-SYMBOL(<lfs_values1>)
                                       FROM sy-tabix.
      IF <lfs_values1>-recnumber <> ps_recax-recnumber OR
         <lfs_values1>-tablename <> ps_recax-tablename.
        EXIT.
      ELSEIF <lfs_values1>-fieldname = 'MANDT'.
        CONTINUE.
      ENDIF.

      CONCATENATE pv_line <lfs_values1>-value INTO pv_line
        SEPARATED BY gc_tab.

      IF ps_recax-recnumber = 1.
        cv_cell = cv_cell + 1.
        APPEND cv_cell TO ct_key_flds.

        CALL FUNCTION 'CONVERSION_EXIT_ISOLA_OUTPUT'
          EXPORTING
            input  = <lfs_values1>-langu
          IMPORTING
            output = lv_langu.

        DATA(lv_temp) = <lfs_values1>-tablename && '-' &&
                        <lfs_values1>-fieldname && '-' &&
                        <lfs_values1>-flag      && '-' && lv_langu.

        CONCATENATE pv_header1 lv_temp INTO pv_header1
          SEPARATED BY gc_tab.

        READ TABLE pt_ddtext INTO DATA(wa_ddtext)
                            WITH KEY tabname   = <lfs_values1>-tablename
                                     fieldname = <lfs_values1>-fieldname
                            BINARY SEARCH.

        CONCATENATE pv_header2 wa_ddtext-ddtext INTO pv_header2
          SEPARATED BY gc_tab.
      ENDIF.
      CLEAR: lv_temp, lv_langu, wa_ddtext.
    ENDLOOP.
  ENDIF.
ENDFORM.

FORM file_transfer .
  DATA: lo_workbooks TYPE ole2_object, " list of workbooks
        lo_workbook  TYPE ole2_object. " workbook

  REPLACE '.bcs' IN gv_file_path WITH ' '.

  GET PROPERTY OF go_excel 'ActiveSheet' = lo_workbook.
  GET PROPERTY OF go_excel 'ActiveWorkbook' = lo_workbooks.

  CALL FUNCTION 'FLUSH'
    EXCEPTIONS
      cntl_system_error = 1
      cntl_error        = 2
      OTHERS            = 3.
  IF sy-subrc = 0.
    CALL METHOD OF lo_workbook 'SAVEAS'
      EXPORTING
        #1 = gv_file_path.
  ENDIF.

  CALL METHOD OF lo_workbooks 'CLOSE'.
  CALL  METHOD OF lo_workbooks 'QUIT'.

  CALL METHOD OF go_excel 'QUIT'.

  FREE OBJECT lo_workbooks.
  FREE OBJECT lo_workbook.
  FREE OBJECT go_excel.
  go_excel-handle = -1.
ENDFORM.

FORM upload_file CHANGING pc_transfer  TYPE scpr_transfertab
                          pc_filename  TYPE localfile
                          pc_file_path TYPE string
                          pv_subrc     TYPE syst_subrc.
* Local Data
  DATA: lo_excel  TYPE REF TO cl_fdt_xl_spreadsheet.
  FIELD-SYMBOLS : <lt_data> TYPE STANDARD TABLE.

  CLEAR pc_transfer.
  PERFORM get_file_path CHANGING pc_filename pc_file_path pv_subrc.

  IF NOT pc_file_path CS 'xls' OR pv_subrc IS NOT INITIAL.
    RETURN.
  ENDIF.

  PERFORM read_excel   USING pc_file_path CHANGING lo_excel pv_subrc.
  lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
               IMPORTING worksheet_names = DATA(lt_worksheets) ).      " Get List of Worksheets

  LOOP AT lt_worksheets ASSIGNING FIELD-SYMBOL(<lfs_sheet>).
    DATA(lo_data) =
     lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                                                          <lfs_sheet> ).
    ASSIGN lo_data->* TO <lt_data>.
    IF sy-subrc <> 0.
      CONTINUE.
    ELSEIF <lfs_sheet> = 'Header'.
      PERFORM: append_init_values USING <lt_data> CHANGING pc_transfer.
    ELSE.
      PERFORM build_transfer_tab USING <lt_data> CHANGING pc_transfer.
    ENDIF.
  ENDLOOP.

ENDFORM.

FORM insert_line USING ps_header    TYPE ty_header
                       pv_recnumber TYPE any
                       pv_value     TYPE any
                 CHANGING pt_transfer TYPE scpr_transfertab.

  APPEND INITIAL LINE TO pt_transfer
                      ASSIGNING FIELD-SYMBOL(<lfs_transfer>).
  IF ps_header-langu IS INITIAL.
    <lfs_transfer>-line     = 'SCPRVALS'.
  ELSE.
    <lfs_transfer>-line     = 'SCPRVALL'.
  ENDIF.
  <lfs_transfer>-line+50  = ps_header-tablename.
  <lfs_transfer>-line+100 = pv_recnumber.
  <lfs_transfer>-line+110 = ps_header-fieldname.
  <lfs_transfer>-line+160 = ps_header-langu.
  <lfs_transfer>-line+170 = ps_header-flag.
  <lfs_transfer>-line+180 = pv_value.

ENDFORM.

FORM insert_mandt USING   pv_tablename TYPE any
                          pv_recnumber TYPE any
                          pv_langu     TYPE any
                 CHANGING pt_transfer TYPE scpr_transfertab.

  APPEND INITIAL LINE TO pt_transfer
                      ASSIGNING FIELD-SYMBOL(<lfs_transfer>).
  IF pv_langu  IS INITIAL.
    <lfs_transfer>-line   = 'SCPRVALS'.
  ELSE.
    <lfs_transfer>-line   = 'SCPRVALL'.
  ENDIF.
  <lfs_transfer>-line+50  = pv_tablename.
  <lfs_transfer>-line+100 = pv_recnumber.
  <lfs_transfer>-line+110 = 'MANDT'.
  <lfs_transfer>-line+170 = 'KEY'.
  <lfs_transfer>-line+180 = sy-mandt.

ENDFORM.

FORM get_file_path CHANGING pc_filename  TYPE localfile
                            pc_file_path TYPE string
                            pv_subrc     TYPE syst_subrc.
* Open File
  CALL FUNCTION 'SCPR_IF_DOWNLOAD_FILENAME_GET'
    EXPORTING
      up_or_down        = 'UP'
    CHANGING
      path_and_filename = pc_file_path
      filename_only     = pc_filename
    EXCEPTIONS
      user_abort        = 1.
  IF sy-subrc <> 0.
    pv_subrc = 6.
  ENDIF.
ENDFORM.

FORM append_init_values  USING    pt_data     TYPE STANDARD TABLE
                         CHANGING pc_transfer TYPE scpr_transfertab.

  LOOP AT pt_data ASSIGNING FIELD-SYMBOL(<lfs_data>).
    ASSIGN COMPONENT 1 OF STRUCTURE <lfs_data>
                                     TO FIELD-SYMBOL(<lv_field1>).
    IF sy-subrc = 0 AND <lv_field1> <> 'SHEETCOUNT' AND
       <lv_field1> <> 'SYSID' AND <lv_field1> <> 'CLIENT'.

      APPEND INITIAL LINE TO pc_transfer ASSIGNING
                                         FIELD-SYMBOL(<lfs_transfer>).
      ASSIGN COMPONENT 2 OF STRUCTURE <lfs_data>
                             TO FIELD-SYMBOL(<lv_fieldx>).
      IF sy-subrc = 0.
        <lfs_transfer>-line    = <lv_field1>.
        <lfs_transfer>-line+50 = <lv_fieldx>.
      ENDIF.

      CASE <lv_field1>.
        WHEN 'DATE'.
          ASSIGN COMPONENT 3 OF STRUCTURE <lfs_data> TO <lv_fieldx>.
          IF sy-subrc = 0.
            <lfs_transfer>-line+70 = <lv_fieldx>.
          ENDIF.
        WHEN 'BCSET'.
          ASSIGN COMPONENT 3 OF STRUCTURE <lfs_data> TO <lv_fieldx>.
          IF sy-subrc = 0.
            <lfs_transfer>-line+85 = <lv_fieldx>.
          ENDIF.
          ASSIGN COMPONENT 4 OF STRUCTURE <lfs_data> TO <lv_fieldx>.
          IF sy-subrc = 0.
            <lfs_transfer>-line+90 = <lv_fieldx>.
          ENDIF.
        WHEN OTHERS.
      ENDCASE.
    ENDIF.
  ENDLOOP.
ENDFORM.

FORM insert_bc_attributes  USING    ps_actobj    TYPE cus_actobj
                                    pt_header    TYPE ty_header
                                    pv_recno     TYPE any
                           CHANGING pc_transfer  TYPE scpr_transfertab.

  APPEND INITIAL LINE TO pc_transfer ASSIGNING
                               FIELD-SYMBOL(<lfs_transfer>).
  <lfs_transfer>+0(50)   = 'SCPRRECA'.
  <lfs_transfer>+50(50)  = pt_header-tablename.
  <lfs_transfer>+100(10) = pv_recno.
  <lfs_transfer>+110(50) = ps_actobj-objectname.
  <lfs_transfer>+160(10) = ps_actobj-objecttype.
  <lfs_transfer>+170(50) = ps_actobj-act_id.

ENDFORM.

FORM read_excel  USING    pc_file_path TYPE string
                 CHANGING po_excel     TYPE REF TO cl_fdt_xl_spreadsheet
                          pv_subrc     TYPE syst_subrc.

  DATA: lt_records       TYPE solix_tab,
        lv_headerxstring TYPE xstring,
        lv_filelength    TYPE i.

  CALL FUNCTION 'GUI_UPLOAD'
    EXPORTING
      filename                = pc_file_path
      filetype                = 'BIN'
    IMPORTING
      filelength              = lv_filelength
      header                  = lv_headerxstring
    TABLES
      data_tab                = lt_records
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      header_too_long         = 11
      unknown_dp_error        = 12
      access_denied           = 13
      dp_out_of_memory        = 14
      disk_full               = 15
      dp_timeout              = 16
      OTHERS                  = 17.
  IF sy-subrc <> 0.
    pv_subrc = sy-subrc.
    RETURN.
  ENDIF.
  "convert binary data to xstring
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_filelength
    IMPORTING
      buffer       = lv_headerxstring
    TABLES
      binary_tab   = lt_records
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.

  IF sy-subrc <> 0.
    pv_subrc = sy-subrc.
    RETURN.
  ENDIF.

  TRY.
      po_excel = NEW cl_fdt_xl_spreadsheet(
                              document_name = pc_file_path
                              xdocument     = lv_headerxstring ) .
    CATCH cx_fdt_excel_core.
  ENDTRY.

ENDFORM.

FORM get_header  USING    pt_data TYPE STANDARD TABLE
                 CHANGING pc_header TYPE tt_header
                          pc_actobj TYPE cus_actobj.
  CLEAR: pc_header.

  READ TABLE pt_data INDEX 1 ASSIGNING FIELD-SYMBOL(<lfs_data>).
  IF sy-subrc = 0.
    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE <lfs_data>
                                           TO FIELD-SYMBOL(<lv_field>).
      IF sy-subrc = 0.
        IF sy-index = 1.
          SELECT *
            FROM cus_actobj
            INTO @pc_actobj
           UP TO 1 ROWS
           WHERE act_id = @<lv_field>.
          ENDSELECT.
        ENDIF.
        APPEND INITIAL LINE TO pc_header ASSIGNING
                                         FIELD-SYMBOL(<lfs_header>).
        SPLIT <lv_field> AT '-' INTO <lfs_header>-tablename
              <lfs_header>-fieldname <lfs_header>-flag
              <lfs_header>-langu.
      ELSE.
        EXIT.
      ENDIF.
    ENDDO.
  ENDIF.
ENDFORM.

FORM add_header  USING    pv_profid    TYPE scpr_id
                          pv_category  TYPE scpr_ctgry
                          pv_proftext  TYPE scpr_text
                 CHANGING pt_data      TYPE tt_out_txt.
  SELECT SINGLE * FROM scprattr
    INTO @DATA(ls_attr)
   WHERE id = @pv_profid AND version = 'N' AND category = @pv_category.

  PERFORM: append_line USING 'VERSION' ' 1' '' '' CHANGING pt_data,
           append_line USING 'DATE' ls_attr-moddate ls_attr-modtime ''
                       CHANGING pt_data,
           append_line USING  'BCSET' ls_attr-id ls_attr-type
                              ls_attr-category CHANGING pt_data,
           append_line USING 'ORGID' ls_attr-id ls_attr-orgid ''
                       CHANGING pt_data,
           append_line USING 'COMPONENT' ls_attr-component '' ''
                       CHANGING pt_data,
           append_line USING 'MINRELEASE' ls_attr-minrelease '' ''
                       CHANGING pt_data,
           append_line USING 'MAXRELEASE' ls_attr-maxrelease '' ''
                       CHANGING pt_data,
           append_line USING 'BCSET_TEXT' pv_proftext '' ''
                       CHANGING pt_data,
           append_line USING 'SYSID'  sy-sysid '' '' CHANGING pt_data,
           append_line USING 'CLIENT' sy-mandt '' '' CHANGING pt_data.
ENDFORM.

FORM append_line USING    pv_name   TYPE any
                          pv_value1 TYPE any
                          pv_value2 TYPE any
                          pv_value3 TYPE any
                 CHANGING pt_data   TYPE tt_out_txt.

  APPEND INITIAL LINE TO pt_data ASSIGNING FIELD-SYMBOL(<lfs_data>).
  CONCATENATE pv_name pv_value1 pv_value2 pv_value3 INTO
                                   <lfs_data>-line SEPARATED BY gc_tab.
ENDFORM.

FORM build_transfer_tab   USING   pt_data     TYPE STANDARD TABLE
                         CHANGING pc_transfer TYPE scpr_transfertab.
  DATA: lt_header TYPE tt_header,
        lt_temp   TYPE scpr_transfertab.
  DATA: ls_actobj TYPE cus_actobj.

  PERFORM get_header USING pt_data CHANGING lt_header ls_actobj.

  LOOP AT pt_data ASSIGNING FIELD-SYMBOL(<lfs_data>) FROM 3.

    ASSIGN COMPONENT 1 OF STRUCTURE <lfs_data>
                                      TO FIELD-SYMBOL(<lfs_recno>).    " Get Record number
    IF sy-subrc <> 0 OR <lfs_recno> IS INITIAL. CONTINUE. ENDIF.

    LOOP AT lt_header ASSIGNING FIELD-SYMBOL(<lfs_headerx>)
                    GROUP BY ( tablename = <lfs_headerx>-tablename ).
      IF <lfs_headerx>-fieldname IS INITIAL. CONTINUE. ENDIF.
      PERFORM insert_mandt USING <lfs_headerx>-tablename <lfs_recno>
                                 <lfs_headerx>-langu
                          CHANGING pc_transfer.
      LOOP AT GROUP <lfs_headerx> ASSIGNING FIELD-SYMBOL(<lfs_header>).
        IF <lfs_header>-fieldname IS INITIAL.
          CONTINUE.
        ELSE.
          DATA(lv_flag) = abap_true.
        ENDIF.
        ASSIGN COMPONENT sy-tabix OF STRUCTURE <lfs_data>
                                      TO FIELD-SYMBOL(<lfs_value>).
        IF sy-subrc = 0.
          PERFORM insert_line USING <lfs_header> <lfs_recno> <lfs_value>
                           CHANGING pc_transfer.
        ENDIF.
      ENDLOOP.
      IF lv_flag = abap_true.
        PERFORM insert_bc_attributes USING    ls_actobj <lfs_headerx>
                       <lfs_recno>   CHANGING lt_temp.
      ENDIF.
      CLEAR lv_flag.
    ENDLOOP.
  ENDLOOP.
  APPEND LINES OF lt_temp TO pc_transfer.
ENDFORM.
