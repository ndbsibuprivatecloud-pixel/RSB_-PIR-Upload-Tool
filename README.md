# RSB_-PIR-Upload-Tool
PIR Upload Tool

Classes 

class ZMASSPIRUPLOAD definition
  public
  final
  create public .

public section.

  types:
*    TYPES:
*      " Type Declaration for Planned Independent Requirements
*      BEGIN OF ty_pir,
*        sno              TYPE int4,
*        material         TYPE bapisitemr-material,
*        plant            TYPE bapisitemr-plant,
*        requirementstype TYPE bapisitemr-requ_type,
**        Version          TYPE bapisitemr-version,
**        reqmtsplannumber TYPE bapisitemr-req_number,
**        VersionActive    TYPE bapisitemr-vers_activ,
**        deleteold        TYPE bapisparam-delete_old,
**        Datetype         TYPE bapisshdin-date_type,
*        reqdate          TYPE bapisshdin-req_date,
*        reqqty           TYPE bapisshdin-req_qty,
*      END OF ty_pir .
      " Type Declaration for Planned Independent Requirements
    BEGIN OF ty_pir,
        sno              TYPE int4,
        material         TYPE bapisitemr-material,
        plant            TYPE bapisitemr-plant,
        requirementstype TYPE bapisitemr-requ_type,
        reqqty           TYPE bapisshdin-req_qty,
        reqdate          TYPE bapisshdin-req_date,

        req01            TYPE bapisshdin-req_qty,
        req02            TYPE bapisshdin-req_qty,
        req03            TYPE bapisshdin-req_qty,
        req04            TYPE bapisshdin-req_qty,
        req05            TYPE bapisshdin-req_qty,
        req06            TYPE bapisshdin-req_qty,
        req07            TYPE bapisshdin-req_qty,
        req08            TYPE bapisshdin-req_qty,
        req09            TYPE bapisshdin-req_qty,
        req10            TYPE bapisshdin-req_qty,
        req11            TYPE bapisshdin-req_qty,
        req12            TYPE bapisshdin-req_qty,
        req13            TYPE bapisshdin-req_qty,
        req14            TYPE bapisshdin-req_qty,
        req15            TYPE bapisshdin-req_qty,
        req16            TYPE bapisshdin-req_qty,
        req17            TYPE bapisshdin-req_qty,
        req18            TYPE bapisshdin-req_qty,
        req19            TYPE bapisshdin-req_qty,
        req20            TYPE bapisshdin-req_qty,
        req21            TYPE bapisshdin-req_qty,
        req22            TYPE bapisshdin-req_qty,
        req23            TYPE bapisshdin-req_qty,
        req24            TYPE bapisshdin-req_qty,
        req25            TYPE bapisshdin-req_qty,
        req26            TYPE bapisshdin-req_qty,
        req27            TYPE bapisshdin-req_qty,
        req28            TYPE bapisshdin-req_qty,
        req29            TYPE bapisshdin-req_qty,
        req30            TYPE bapisshdin-req_qty,
        req31            TYPE bapisshdin-req_qty,

      END OF ty_pir .
  types:
    BEGIN OF ty_pirdetails,
        sno              TYPE int4,
        material         TYPE bapisitemr-material,
        plant            TYPE bapisitemr-plant,
        requirementstype TYPE bapisitemr-requ_type,
*        deleteold        TYPE bapisparam-delete_old,
        reqdate          TYPE bapisshdin-req_date,
        reqqty           TYPE bapisshdin-req_qty,
        recordno         TYPE int4,
      END OF ty_pirdetails .

  constants XLSX type STRING value 'XLSX' ##NO_TEXT.
  constants FILENAMEDEF type STRING value 'PIR Upload' ##NO_TEXT.
  constants BINTYPE type CHAR10 value 'BIN' ##NO_TEXT.

  class-methods OPENFILE
    returning
      value(FILE) type RLGRAP-FILENAME .
  class-methods DOWNLOADTEMPLATE
    importing
      !SY_UCOMM type SYST_UCOMM optional .
  class-methods LOADPIR
    importing
      !FILE type RLGRAP-FILENAME optional
      !PIR_DATA type ref to DATA optional
      !FLAG type CHAR1
      !FLAG_ZPIR type CHAR1 .
protected section.
private section.
ENDCLASS.



CLASS ZMASSPIRUPLOAD IMPLEMENTATION.


METHOD downloadtemplate.

  "--------------------------------------------------------
  " Field Symbols & Data Declarations
  "--------------------------------------------------------
  FIELD-SYMBOLS <gt_final> TYPE ANY TABLE.

  DATA: columns    TYPE if_fdt_doc_spreadsheet=>t_column,
        lt_comp    TYPE if_fdt_doc_spreadsheet=>t_column,
        lt_comp_rtts TYPE cl_abap_structdescr=>component_table,
        path       TYPE string,
        fullpath   TYPE string,
        action     TYPE i,
        filename   TYPE string,
        lr_struct  TYPE REF TO cl_abap_structdescr,
        lr_table   TYPE REF TO cl_abap_tabledescr,
        gr_data    TYPE REF TO data.

  DATA: act_dated TYPE TABLE OF rke_dat,
        lv_first  TYPE d,
        lv_last   TYPE d,
        lv_date   TYPE d.

  "--------------------------------------------------------
  " Determine selected period (current month / next month)
  "--------------------------------------------------------
  IF sy_ucomm = 'FC01'.

    CONCATENATE sy-datum(4) sy-datum+4(2) '01' INTO lv_first.

    CALL FUNCTION 'RP_LAST_DAY_OF_MONTHS'
      EXPORTING day_in = lv_first
      IMPORTING last_day_of_month = lv_last.

  ELSEIF sy_ucomm = 'FC02'.

    DATA lv_newdate TYPE d.

      CALL FUNCTION 'RP_CALC_DATE_IN_INTERVAL'
        EXPORTING
          date      = sy-datum
          months    = 1
          days      = 0
          years     = 0
        IMPORTING
          calc_date = lv_newdate.

    CONCATENATE lv_newdate(4) lv_newdate+4(2) '01' INTO lv_first.

    CALL FUNCTION 'RP_LAST_DAY_OF_MONTHS'
      EXPORTING day_in = lv_first
      IMPORTING last_day_of_month = lv_last.

  ENDIF.


  "--------------------------------------------------------
  " Read Fact Days
  "--------------------------------------------------------
  CALL FUNCTION 'RKE_SELECT_FACTDAYS_FOR_PERIOD'
    EXPORTING
      i_datab = lv_first
      i_datbi = lv_last
      i_factid = 'ZO'
    TABLES
      eth_dats = act_dated.


  "--------------------------------------------------------
  " STATIC Columns (Technical Name must be valid ABAP field!)
  "--------------------------------------------------------
  APPEND VALUE #(
      name         = 'SNO'
      display_name = 'SNO'
      is_result    = abap_true
      type         = cl_abap_elemdescr=>get_int8( )
  ) TO lt_comp.

  APPEND VALUE #(
      name         = 'MATERIAL'
      display_name = 'MATERIAL'
      is_result    = abap_true
      type         = cl_abap_elemdescr=>get_c( 18 )
  ) TO lt_comp.

  APPEND VALUE #(
      name         = 'PLANT'
      display_name = 'PLANT'
      is_result    = abap_true
      type         = cl_abap_elemdescr=>get_c( 4 )
  ) TO lt_comp.

  APPEND VALUE #(
      name         = 'REQUIREMENTSTYPE'
      display_name = 'REQUIREMENT TYPE'
      is_result    = abap_true
      type         = cl_abap_elemdescr=>get_c( 3 )
  ) TO lt_comp.


  "--------------------------------------------------------
  " DYNAMIC DATE Columns
  "--------------------------------------------------------
  lv_date = lv_first.

  WHILE lv_date <= lv_last.

    IF line_exists( act_dated[ periodat = lv_date ] ).

      APPEND VALUE #(
          name         = |D{ lv_date }|     " technical field name
          display_name = |{ lv_date+6(2) }/{ lv_date+4(2) }/{ lv_date+0(4) }|  " visible Excel header
*          display_name = |{ lv_date DATE = USER }|  " visible Excel header
          is_result    = abap_true
          type         = cl_abap_elemdescr=>get_c( 10 )
      ) TO lt_comp.

    ENDIF.

    lv_date = lv_date + 1.

  ENDWHILE.


  "--------------------------------------------------------
  " Convert Excel Column Definition → RTTS ABAP Structure Components
  "--------------------------------------------------------
  LOOP AT lt_comp ASSIGNING FIELD-SYMBOL(<col>).

    APPEND VALUE #(
       name = <col>-name
       type = <col>-type
    ) TO lt_comp_rtts.

  ENDLOOP.


  "--------------------------------------------------------
  " Create Dynamic Structure & Internal Table
  "--------------------------------------------------------
  lr_struct = cl_abap_structdescr=>create( lt_comp_rtts ).
  lr_table  = cl_abap_tabledescr=>create( lr_struct ).

  CREATE DATA gr_data TYPE HANDLE lr_table.
  ASSIGN gr_data->* TO <gt_final>.


  "--------------------------------------------------------
  " Create Excel Document
  "--------------------------------------------------------
  DATA(bin) = cl_fdt_xl_spreadsheet=>if_fdt_doc_spreadsheet~create_document(
                columns      = lt_comp
                itab         = REF #( <gt_final> )
                iv_call_type = '' ).


  "--------------------------------------------------------
  " File Save Dialog
  "--------------------------------------------------------
  IF xstrlen( bin ) > 0.

    filename = filenamedef.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING
        default_file_name = filename
        default_extension = 'XLSX'
        file_filter       = cl_gui_frontend_services=>filetype_all
      CHANGING
        filename = filename
        path     = path
        fullpath = fullpath
        user_action = action ).

  ENDIF.


  "--------------------------------------------------------
  " Download Excel File
  "--------------------------------------------------------
  IF action = cl_gui_frontend_services=>action_ok.

    DATA(raws) =
      cl_bcs_convert=>xstring_to_solix( iv_xstring = bin ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        filename     = fullpath
        filetype     = 'BIN'
        bin_filesize = xstrlen( bin )
      CHANGING
        data_tab     = raws ).

  ENDIF.

ENDMETHOD.


  METHOD loadpir.

    TYPES:tt_bom_create TYPE TABLE FOR CREATE zr_pp__bom_table,
          tt_bom_update TYPE TABLE FOR UPDATE  zr_pp__bom_table.

    DATA: bomcreate    TYPE tt_bom_create,
          updatebom    TYPE tt_bom_update,
          bomcreatetbl TYPE TABLE OF tt_bom_create.
*      bom_data     TYPE TABLE OF ty_bom.


    TYPES:BEGIN OF pirstatus,
            sno     TYPE int4,
            message TYPE char100,
            type    TYPE char5,
          END OF pirstatus.

    TYPES :BEGIN OF coloumnsname,
             column_name TYPE lvc_fname,
             longtext    TYPE scrtext_l,
           END OF coloumnsname.
    TYPES:
      " Type Declaration for Planned Independent Requirements
      BEGIN OF ty_mpr,
        sno              TYPE int4,
        material         TYPE bapisitemr-material,
        plant            TYPE bapisitemr-plant,
        requirementstype TYPE bapisitemr-requ_type,
        bomdate          TYPE  zpp__bom_table-bom_date,
      END OF ty_mpr.

    DATA: mpr TYPE ty_mpr.

    DATA:row                      TYPE REF TO data,
         pirdetail                TYPE ty_pir,
         pirdetails               TYPE TABLE OF ty_pir,
         pir_details              TYPE TABLE OF ty_pirdetails,
         date                     TYPE string,
         requirements_schedule_in	TYPE TABLE OF	bapisshdin,
         returns                  TYPE TABLE OF bapireturn1,
         pirstatusdetails         TYPE TABLE OF pirstatus,
         xtab                     TYPE cpt_x255,
         filedetail               TYPE string,
         clounmsames              TYPE TABLE OF coloumnsname,
         go_columns               TYPE REF TO cl_salv_columns_table,
         go_column                TYPE REF TO cl_salv_column_table,
         delete                   TYPE bapisparam-delete_old,
         reqmtsplannumber         TYPE bapisitemr-req_number.

**************

    filedetail = file.
    " Open File exlorer
    TRY.
        cl_gui_frontend_services=>gui_upload(
          EXPORTING
            filename   = filedetail
            filetype   = bintype
          IMPORTING
            filelength = DATA(filesize)
          CHANGING
            data_tab   = xtab ).
      CATCH cx_root INTO DATA(lo_msg).
        MESSAGE lo_msg TYPE /isdfps/cl_const_abc_123=>gc_e.
    ENDTRY.

    IF xtab IS NOT INITIAL.
      " Converting bin to string
      TRY.
          cl_scp_change_db=>xtab_to_xstr(
            EXPORTING
              im_xtab    = xtab
              im_size    = filesize
            IMPORTING
              ex_xstring = DATA(xstring) ).

          " Convert Excel to String
          DATA(excel) = NEW cl_fdt_xl_spreadsheet( document_name = filedetail      "data declaration to excel
                                                   xdocument     = xstring ).

          excel->if_fdt_doc_spreadsheet~get_worksheet_names( IMPORTING worksheet_names = DATA(worksheets) ).

        CATCH cx_fdt_excel_core INTO DATA(formaterror).
          IF formaterror IS NOT INITIAL.
            MESSAGE 'Invalid document format' TYPE /isdfps/cl_const_abc_123=>gc_e.
          ENDIF.
        CATCH cx_root INTO DATA(error).

          IF error IS NOT INITIAL.
            MESSAGE error TYPE  /isdfps/cl_const_abc_123=>gc_e. EXIT.
          ENDIF.
      ENDTRY.
    ENDIF.
    CHECK worksheets IS NOT INITIAL.

    " Get Table Fields from the System and Map them accordingly
    DATA(table) = excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( worksheets[ 1 ] ).
    FIELD-SYMBOLS: <worksheet> TYPE ANY TABLE.
    ASSIGN table->* TO <worksheet>.                                          "Assign worksheet to row
    CREATE DATA row LIKE LINE OF <worksheet>.
    ASSIGN row->* TO FIELD-SYMBOL(<row>).                                    "Assign references table to row

    IF <row> IS ASSIGNED.

      """" Date from Excel
      LOOP AT <worksheet> ASSIGNING FIELD-SYMBOL(<ls_sheet>).

        DATA: daten TYPE char10.
        ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_e OF STRUCTURE <ls_sheet> TO FIELD-SYMBOL(<celln>).
        IF <celln> IS ASSIGNED.
          daten = <celln>.
          UNASSIGN <celln>.
        ENDIF.
        DATA: lv_first TYPE d,
              lv_last  TYPE d,
              lv_date  TYPE d.

        CONCATENATE daten+6(4) daten+3(2) '01' INTO lv_first. "First day of current month
        CALL FUNCTION 'RP_LAST_DAY_OF_MONTHS'
          EXPORTING
            day_in            = lv_first
          IMPORTING
            last_day_of_month = lv_last.

        DATA(day_one) = lv_first.

        DATA: act_dated TYPE TABLE OF rke_dat.
        CALL FUNCTION 'RKE_SELECT_FACTDAYS_FOR_PERIOD'
          EXPORTING
            i_datab               = lv_first
            i_datbi               = lv_last
            i_factid              = 'ZO'
          TABLES
            eth_dats              = act_dated
          EXCEPTIONS
            date_conversion_error = 1
            OTHERS                = 2.
        IF sy-subrc <> 0.
* Implement suitable error handling here
        ENDIF.

        DATA(no_days) = lines( act_dated ).
        EXIT.
      ENDLOOP.


      FIELD-SYMBOLS: <lt_ws>  TYPE ANY TABLE,
                     <ls_row> TYPE any.
*                     <f_mat>   TYPE any,
*                     <f_plant> TYPE any.

      " Your dynamic internal table reference
      ASSIGN pir_data->* TO <lt_ws>.

      " Key structure for HASHED TABLE
      TYPES: BEGIN OF ty_key,
               material TYPE string,
               plant    TYPE string,
             END OF ty_key.

      DATA: ls_key  TYPE ty_key,
            lt_hash TYPE HASHED TABLE OF ty_key
                     WITH UNIQUE KEY material plant.

      """ Duplicate
      LOOP AT <worksheet> ASSIGNING <ls_row>.
        ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_b OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<f_mat>).

        ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_c OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<f_plant>).

        IF <f_mat> IS INITIAL OR <f_plant> IS INITIAL.
          CONTINUE. "skip blank rows
        ENDIF.

        ls_key-material = <f_mat>.
        ls_key-plant    = <f_plant>.

        " Try insert → detects duplicates
        INSERT ls_key INTO TABLE lt_hash.
        IF sy-subrc <> 0.
          MESSAGE e000(zmsg) WITH
            |Duplicate found for Material { <f_mat> } Plant { <f_plant> }|.
        ENDIF.

      ENDLOOP.


      DATA(count) = -1.
      DATA: alf               TYPE char2,
            requirements_item TYPE bapisitemr.


      SELECT FROM mara
        FIELDS matnr
        WHERE mtart = 'ZFGM'  OR mtart  = 'ZSFG'
        INTO TABLE @DATA(lt_mara).
      IF sy-subrc = 0.
      ELSE.
      ENDIF.


      SELECT id,
             material,
             plant,
            bom_date,
            totalqty
      FROM zpp__bom_table
          INTO TABLE @DATA(lt_existing).
      IF sy-subrc = 0.
      ELSE.
      ENDIF.

      " Converting reference table to internal table
      LOOP AT <worksheet> ASSIGNING <row> .
        count = count + 1.
        IF sy-tabix LE 1.
          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_a OF STRUCTURE <row> TO FIELD-SYMBOL(<cell>).
          IF <cell> IS ASSIGNED.
            IF  <cell> <> 'SNO'.
              MESSAGE 'Please Upload the Proper file' TYPE 'E'.
            ENDIF.
            UNASSIGN <cell>.
          ENDIF.
          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_b OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            IF <cell> <> 'MATERIAL'.
              MESSAGE 'Please Upload the Proper file' TYPE 'E'.
            ENDIF.
            UNASSIGN <cell>.
          ENDIF.
          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_c OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            IF  <cell> <> 'PLANT'.
              MESSAGE 'Please Upload the Proper file' TYPE 'E'.
            ENDIF.
            UNASSIGN <cell>.
          ENDIF.
          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_d OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            IF  <cell> <> 'REQUIREMENT TYPE'.
              MESSAGE 'Please Upload the Proper file' TYPE 'E'.
            ENDIF.
            UNASSIGN <cell>.
          ENDIF.

        ENDIF.

        IF sy-tabix LE 1.                                                        "Not considering first row in excel
          CONTINUE.
        ENDIF.

**        "" Validating material based on Stategy grp field in MARC table
**        ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_b OF STRUCTURE <row> TO FIELD-SYMBOL(<cell2>).
**         ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_c OF STRUCTURE <row> TO FIELD-SYMBOL(<cell3>).
**          IF <cell2> IS ASSIGNED AND <cell3> IS ASSIGNED.
**
**          SELECT FROM zi_marc_cds_view AS marc
**            FIELDS matnr,
**                   werks,
**                   strgr
**            WHERE matnr = @<cell2>
**            into TABLE @DATA(validate_material).
**          IF validate_material IS INITIAL.
**          ELSE.
**          ENDIF.
**
**          UNASSIGN: <cell2>,<cell3>.
**        ENDIF.
***

        DATA lv_day TYPE n LENGTH 2.
        lv_first = day_one.
        alf = 'E'.
        DO no_days TIMES.
          DATA req_qty TYPE string.
          ASSIGN COMPONENT alf OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            req_qty = <cell>.
            UNASSIGN <cell>.
          ENDIF.
          DATA(off) = sy-index + 4.

          IF off < 26 . "AND off <> 5.
            alf = sy-abcde+off(1).
          ELSEIF off = 26.
            alf = /isdfps/cl_const_abc_123=>gc_a && /isdfps/cl_const_abc_123=>gc_a.
          ELSEIF off = 27.
            alf = /isdfps/cl_const_abc_123=>gc_a && /isdfps/cl_const_abc_123=>gc_b.
          ELSEIF off = 28.
            alf = /isdfps/cl_const_abc_123=>gc_a && /isdfps/cl_const_abc_123=>gc_c.
          ELSEIF off = 29.
            alf = /isdfps/cl_const_abc_123=>gc_a && /isdfps/cl_const_abc_123=>gc_d.
          ENDIF.

          lv_day = sy-index.
          lv_day = |{ lv_day ALPHA = IN }|.

          DATA(lv_field) = |REQ{ lv_day }|.
          DATA(req_date) = VALUE #( act_dated[ lv_day ]-periodat OPTIONAL ).
          requirements_schedule_in = VALUE #( ( date_type = '1' "<pirdetails>-datetype
                                                req_date  = req_date
                                                req_qty   = req_qty ) ).
          IF req_qty = ''.
            CONTINUE.
          ENDIF.

          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_b OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            mpr-material = <cell>.
            mpr-material = |{ mpr-material ALPHA = IN }|.
            UNASSIGN <cell>.
          ENDIF.
          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_c OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            mpr-plant = <cell>.
            UNASSIGN <cell>.
          ENDIF.
          ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_d OF STRUCTURE <row> TO <cell>.
          IF <cell> IS ASSIGNED.
            mpr-requirementstype = <cell>.
            UNASSIGN <cell>.
          ENDIF.
******************************************************************************************************************************
          " Start of Below Logic is for Upload ZPP__BOM_TABLE

          IF flag IS NOT INITIAL.
            READ TABLE lt_mara INTO DATA(ls_mara) WITH KEY matnr  = mpr-material.
            IF sy-subrc  = 0.
              ASSIGN lt_existing[ material =   mpr-material "<fs_row>-material
                                  plant = mpr-plant
                                  bom_date  =  req_date ]
                     TO FIELD-SYMBOL(<fs_exist>).
              IF sy-subrc = 0.

                INSERT VALUE #(

                    id                        =  <fs_exist>-id
*                material                  = |{ <fs_row>-material ALPHA = IN }|
                    material                  = mpr-material
                    plant                     =  mpr-plant
                    bomdate                   = req_date
                    totalqty                  = req_qty
*                %control-material         = if_abap_behv=>mk-on  "mk-on
*                %control-plant            = if_abap_behv=>mk-on             "mk-on
*                %control-bomdate         = if_abap_behv=>mk-on            " mk-on
                    %control-totalqty         = if_abap_behv=>mk-on             "mk-on
                ) INTO TABLE updatebom.

                IF updatebom IS NOT INITIAL.

                  MODIFY ENTITIES OF zr_pp__bom_table
                         ENTITY zr_pp__bom_table
                           UPDATE
                             FROM updatebom
                       FAILED DATA(updatebom_failed)
                       REPORTED DATA(updatebom_reported)
                       MAPPED DATA(updatebom_mapped).
                  COMMIT ENTITIES .

                  CLEAR updatebom.
                ENDIF.
              ELSE.

                INSERT VALUE #(

                  %cid                    = sy-tabix
*                id                        =  <fs_exist>-id
*                material                  = |{ <fs_row>-material ALPHA = IN }|
                    material                  =  mpr-material "<fs_row>-material
                    plant                     =  mpr-plant
                    bomdate                  = req_date
                    totalqty                 = req_qty
                    %control-material         = if_abap_behv=>mk-on  "mk-on
                    %control-plant            = if_abap_behv=>mk-on             "mk-on
                    %control-bomdate         = if_abap_behv=>mk-on            " mk-on
                    %control-totalqty        = if_abap_behv=>mk-on             "mk-on

              ) INTO TABLE bomcreate.

                IF bomcreate IS NOT INITIAL.
                  MODIFY ENTITIES OF zr_pp__bom_table
                         ENTITY zr_pp__bom_table
                           CREATE
                             FROM bomcreate
                       FAILED DATA(bomcreate_failed)
                       REPORTED DATA(bomcreate_reported)
                       MAPPED DATA(bomcreate_mapped).
                  COMMIT ENTITIES .

                  CLEAR bomcreate.
                ENDIF.


              ENDIF.

            ENDIF.

          ENDIF.

          " End of Logic is for Upload ZPP__BOM_TABLE

*******************************************************************************************************************************

          IF flag_zpir IS NOT INITIAL.


            ""new changes 18.12.25
            "" Validating material based on Stategy grp field in MARC table
            ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_b OF STRUCTURE <row> TO FIELD-SYMBOL(<cell2>).
            ASSIGN COMPONENT /isdfps/cl_const_abc_123=>gc_c OF STRUCTURE <row> TO FIELD-SYMBOL(<cell3>).
            IF <cell2> IS ASSIGNED AND <cell3> IS ASSIGNED.
              DATA: valdmtnr TYPE bapisitemr-material.
              valdmtnr = |{ <cell2> ALPHA = IN }|.
              SELECT SINGLE FROM zi_marc_cds_view AS marc
                FIELDS matnr,
                       werks,
                       strgr
*                WHERE werks = @<cell3>
                WHERE matnr = @valdmtnr AND werks = @<cell3>
                INTO @DATA(validate_material).
              IF validate_material-strgr IS NOT INITIAL.
**              ELSE.
**              ENDIF.
**
**              UNASSIGN: <cell2>,<cell3>.
**            ENDIF.
                ""end changes

                CALL FUNCTION 'BAPI_REQUIREMENTS_CHANGE'
                  EXPORTING
                    material                 = mpr-material
                    plant                    = mpr-plant
                    requirementstype         = mpr-requirementstype
                    version                  = '00' "<pirdetails>-version
                    reqmtsplannumber         = reqmtsplannumber "<pirdetails>-reqmtsplannumber
                    vers_activ               = abap_true "<pirdetails>-versionactive
                    do_commit                = abap_true
                    update_mode              = abap_true
                    delete_old               = delete
                  TABLES
                    requirements_schedule_in = requirements_schedule_in
                    return                   = returns.

                WAIT UP TO 1 / 12 SECONDS.
              IF VALUE #( returns[ type = /isdfps/cl_const_abc_123=>gc_e ]-type OPTIONAL ) IS NOT INITIAL AND VALUE #( returns[ type = /isdfps/cl_const_abc_123=>gc_e ]-message OPTIONAL ) CS 'No new requirements can be created during a change transaction'.

                  requirements_item = VALUE #( material = mpr-material plant = mpr-plant
                                               requ_type = mpr-requirementstype version = '00' vers_activ = 'X' ).
                  CALL FUNCTION 'BAPI_REQUIREMENTS_CREATE'
                    EXPORTING
                      requirements_item        = requirements_item
*                     REQUIREMENT_PARAM        =
                      do_commit                = 'X'
                      update_mode              = 'X'
*                     REFER_TYPE               = ' '
*                     PROFILID                 = ' '
                    TABLES
                      requirements_schedule_in = requirements_schedule_in
                      return                   = returns.

                  WAIT UP TO 1 / 12 SECONDS.
                  IF VALUE #( returns[ type = /isdfps/cl_const_abc_123=>gc_e ]-type OPTIONAL ) IS NOT INITIAL.
                    pirstatusdetails = VALUE #( BASE pirstatusdetails FOR <pirstatusdetails> IN returns WHERE ( type    = /isdfps/cl_const_abc_123=>gc_e )
                                                                                                              ( sno     = count"<pirdetails>-sno
                                                                                                                message = <pirstatusdetails>-message
                                                                                                                type    = <pirstatusdetails>-type ) ).

                    EXIT.
                  ENDIF.
                ELSEIF VALUE #( returns[ type = /isdfps/cl_const_abc_123=>gc_e ]-type OPTIONAL ) IS NOT INITIAL.
                  pirstatusdetails = VALUE #( BASE pirstatusdetails FOR <pirstatusdetails> IN returns WHERE ( type    = /isdfps/cl_const_abc_123=>gc_e )
                                                                                                            ( sno     = count"<pirdetails>-sno
                                                                                                              message = <pirstatusdetails>-message
                                                                                                              type    = <pirstatusdetails>-type ) ).

                  EXIT.
                ENDIF.

                ""new changes 18.12.25
              ELSE.
*
*                pirstatusdetails = VALUE #(   ( sno     = count"<pirdetails>-sno
*                                                 message = |Strategy Group not maintained { valdmtnr } (MATNR) in { <cell3> } (WERKS)| "<pirstatusdetails>-message
*                                                 type    = 'E' ) ).

                APPEND VALUE #(
                                sno     = count "<pirdetails>-sno
                                message = |Strategy Group not maintained { valdmtnr } (MATNR) in { <cell3> } (WERKS)|
                                type    = 'E'
                                ) TO pirstatusdetails.
                EXIT.
              ENDIF.

              UNASSIGN: <cell2>,<cell3>.
            ENDIF.
            ""end changes



            lv_first = lv_first + 1.

            CLEAR:returns,requirements_schedule_in,reqmtsplannumber,req_qty,req_date.

          ENDIF.
        ENDDO.
        CLEAR alf.

        """ End of changes

      ENDLOOP.
    ENDIF.

***Display
    IF pirstatusdetails IS INITIAL.
      " populate success message
      MESSAGE s001(zmasspiruploadval).
    ELSE.
      " Display
      TRY.
          cl_salv_table=>factory(                       "#EC CI_NOORDER
            IMPORTING
              r_salv_table = DATA(salvpirstatus)
            CHANGING
              t_table      = pirstatusdetails ).
        CATCH cx_salv_msg.
      ENDTRY.
      IF salvpirstatus IS BOUND.
        DATA(lo_functions) = salvpirstatus->get_functions( ).
        lo_functions->set_all( abap_true ).

        clounmsames = VALUE #( ( column_name = TEXT-001 longtext = TEXT-004 )
                               ( column_name = TEXT-002 longtext = TEXT-005 )
                               ( column_name = TEXT-003 longtext = TEXT-006 ) ).

        LOOP AT clounmsames ASSIGNING FIELD-SYMBOL(<coloumn>).
          TRY.
              go_columns ?= salvpirstatus->get_columns( ).
              go_column ?= go_columns->get_column( <coloumn>-column_name ).
              go_column->set_long_text( value = <coloumn>-longtext ).
              go_column->set_medium_text( value = '' ).
              go_column->set_short_text( value = '' ).
            CATCH cx_salv_not_found.
          ENDTRY.
        ENDLOOP.
        salvpirstatus->display( ).
      ENDIF.
    ENDIF.
*    ENDIF.
  ENDMETHOD.


  METHOD openfile.
    "Declartion For the Open Dialog
    DATA:windowtitle TYPE string,
         filetable   TYPE filetable,
         rc          TYPE i.
    "Call the File Open Dialog to select a file
    cl_gui_frontend_services=>file_open_dialog(
   EXPORTING
     window_title            = windowtitle
   CHANGING
     file_table              = filetable
     rc                      = rc
   EXCEPTIONS
     file_open_dialog_failed = 1
     cntl_error              = 2
     error_no_gui            = 3
     not_supported_by_gui    = 4
     OTHERS                  = 5 ).
    IF sy-subrc = 0.
      file = VALUE #( filetable[ 1 ]-filename OPTIONAL ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.
***********************************************************************************************************************

Report / program 

* Report ZMASSPIRUPLOAD                                                *
*----------------------------------------------------------------------*
* Title:          PIR Upload                                           *
* RICEF#:         RSB_PP_01                                            *
* Transaction:    ZMASSPIRUPLOAD                                       *
*----------------------------------------------------------------------*
* Copyright:      NDBS, Inc.                                           *
* Client:         RSB                                                  *
*----------------------------------------------------------------------*
* Developer:      Saakshi.D (NTT_ABAP3)                                *
* Creation Date:  12/05/2025                                           *
* Description:    Mass uploading Planned Independent Requirements(PIRs)*
*----------------------------------------------------------------------*
* Modification History                                                 *
*----------------------------------------------------------------------*
* Modified by:    <Developer (full name and user name)>                *
* Date:           <Date>                                               *
* Transport:      <Transport Request #>                                *
* Description:                                                         *
*<Description of the change (or the source for the initial creation if *
* a template or SAP program was used as a starting point>              *
*----------------------------------------------------------------------*
REPORT zmasspirupload NO STANDARD PAGE HEADING
                           MESSAGE-ID  zmasspiruploadval.
" Tables
TABLES:sscrfields.
" Selection-screen
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME.
*  SELECTION-SCREEN COMMENT 01(11) TEXT-001 FOR FIELD file MODIF ID new.

  PARAMETERS: file     TYPE rlgrap-filename,
              p_pirflg AS CHECKBOX,
              p_flag   AS CHECKBOX.

SELECTION-SCREEN END OF BLOCK b1.

**AT SELECTION-SCREEN.



" Initialization

INITIALIZATION.

  sscrfields-functxt_01 = 'Present Month Template' ."TEXT-002."'Download Template'.

  SELECTION-SCREEN FUNCTION KEY 1. "button on the application toolbar
  "Begin of changes
  sscrfields-functxt_02 = 'Next Month Template'.

  SELECTION-SCREEN FUNCTION KEY 2. "button on the application toolbar
  "End of changes

AT SELECTION-SCREEN ON VALUE-REQUEST FOR file.
  file = zmasspirupload=>openfile(  ).

AT SELECTION-SCREEN.
  IF file IS INITIAL AND sy-ucomm = TEXT-040."ONLI
    MESSAGE e002(zmasspiruploadval).
  ENDIF.
  IF p_flag IS INITIAL AND p_pirflg IS INITIAL.
    MESSAGE 'Please select at least one Checkbox  (PP or PIR)' TYPE 'E'.
  ENDIF.

  " Download Template
  CHECK sy-ucomm = TEXT-022 OR sy-ucomm = TEXT-023. "'FC01'.
  TRY.
      zmasspirupload=>downloadtemplate( sy-ucomm ).
    CATCH cx_root.
  ENDTRY.


START-OF-SELECTION.
  " Upload PIR
  TRY.
      zmasspirupload=>loadpir( file = file
                               flag = p_flag
                               flag_zpir = p_pirflg  ).
    CATCH cx_root.
  ENDTRY.

  ***************************************************************************************************************
