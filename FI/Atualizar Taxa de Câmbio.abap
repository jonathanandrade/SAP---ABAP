* ======================================================================
*                         � Abap
* ======================================================================
* Name........: Atualiza��o autom�tica das taxas de c�mbio
* Author......: Jonathan de Andrade
* Date........: 08.06.2016 10:45:22
*-----------------------------------------------------------------------
* Description: Atualiza��o autom�tica das taxas de c�mbio
* - Busca arquivo .csv via http
* - Gera��o de Logs
* - Enviar email
* ======================================================================

REPORT zfir0021.

CONSTANTS: gc_taxas         TYPE tvarvc-name VALUE 'Z_FI_ATUALIZA_OB08',
           gc_cat_taxa      TYPE tvarvc-name VALUE 'Z_FI_CAT_TX_CAMBIO',
           gc_alerta_email  TYPE tvarvc-name VALUE 'Z_FI_ALERTA_OB08',
           gc_log_object    TYPE balobj_d    VALUE 'Z_TAXA_CAMBIO',
           gc_log_subobject TYPE balsubobj   VALUE 'ZTAXA'.

RANGES: r_taxa FOR tcurr-fcurr,
        r_cat_taxa FOR tcurr-kurst.

" Vari�veis Globais
DATA: gv_buffer  TYPE string,
      gv_url1    TYPE char300 VALUE 'http://www4.bcb.gov.br/Download/fechamento/AAAAMMDD.csv'.

DATA: text1 TYPE char50,
      text2 TYPE char50,
      text3 TYPE char50,
      text4 TYPE char50.

" Vari�veis Log
DATA: lw_log TYPE bal_s_log,
      lt_log_handle TYPE bal_t_logh,
      lw_log_handle TYPE balloghndl.

*&--------------------------------------------------------------------*
*& TYPES
*&--------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_tvarvc,
    name TYPE tvarvc-name,
    low  TYPE tvarvc-low,
  END OF ty_tvarvc,

  BEGIN OF ty_cat_taxa,
    kurst TYPE tcurr-kurst,
  END OF ty_cat_taxa,

  BEGIN OF ty_tcurf,
    kurst     TYPE tcurf-kurst,
    fcurr     TYPE tcurf-fcurr,
    tcurr     TYPE tcurf-tcurr,
    ffact     TYPE tcurf-ffact,
    tfact     TYPE tcurf-tfact,
  END OF ty_tcurf,

  BEGIN OF ty_line_tab_cotacao_http,
    col1 TYPE char100,
  END OF ty_line_tab_cotacao_http,

  BEGIN OF ty_line_tab_cotacao,
    camp01 TYPE string, " Data
    camp02 TYPE string, " C�digo
    camp03 TYPE string, " Tipo
    camp04 TYPE string, " Moeda
    camp05 TYPE string, " Compra
    camp06 TYPE string, " Venda
    camp07 TYPE string, " Par/Compra
    camp08 TYPE string, " Par/Venda
  END OF ty_line_tab_cotacao.

*&--------------------------------------------------------------------*
*& TABELAS INTERNAS
*&--------------------------------------------------------------------*
DATA: lt_tvarvc TYPE TABLE OF ty_tvarvc,
      lt_cotacao_http TYPE STANDARD TABLE OF ty_line_tab_cotacao_http,
      lt_cotacao      TYPE STANDARD TABLE OF ty_line_tab_cotacao,
      lt_tcurf TYPE TABLE OF ty_tcurf,
      lt_exch_rate TYPE TABLE OF bapi1093_0,
      lt_cat_taxa TYPE TABLE OF ty_cat_taxa.

*&--------------------------------------------------------------------*
*& ESTRUTURAS
*&--------------------------------------------------------------------*
DATA: lw_tvarvc TYPE ty_tvarvc,
      lw_cotacao_http LIKE LINE OF lt_cotacao_http,
      lw_cotacao LIKE LINE OF lt_cotacao,
      lw_tcurf LIKE LINE OF lt_tcurf,
      lw_exch_rate LIKE LINE OF lt_exch_rate,
      lw_cat_taxa TYPE ty_cat_taxa.

*&--------------------------------------------------------------------*
*& INITIALIZATION
*&--------------------------------------------------------------------*
INITIALIZATION.
  PARAMETERS: p_data TYPE sy-datum.

*&--------------------------------------------------------------------*
*& START-OF-SELECTION
*&--------------------------------------------------------------------*
START-OF-SELECTION.

  lw_log-object    = gc_log_object.
  lw_log-subobject = gc_log_subobject.
  lw_log-alprog    = sy-repid.

  " Inicializar Log
  CALL FUNCTION 'BAL_LOG_CREATE'
    EXPORTING
      i_s_log                 = lw_log
    IMPORTING
      e_log_handle            = lw_log_handle
    EXCEPTIONS
      log_header_inconsistent = 1
      OTHERS                  = 2.
  IF sy-subrc <> 0.
    ##NEEDED.
  ENDIF.

  PERFORM f_get_http_data.
  PERFORM f_split_http_data.
  PERFORM f_fill_data.

  PERFORM f_save_log.

  MESSAGE s046(zfi).

*&---------------------------------------------------------------------*
*&      Form  F_GET_HTTP_DATA
*&---------------------------------------------------------------------*
FORM f_get_http_data.

  DATA : lt_response_headers TYPE TABLE OF text WITH HEADER LINE,
         lv_lcstatcode       TYPE char1,
         lv_entbodylen       TYPE i,
         lv_lcstring         TYPE string,
         lv_stat             TYPE char2,
         lv_data             TYPE char10,
         lv_datum            TYPE sy-datum.

  DATA : lt_soli_tab         TYPE soli_tab,
         lw_soli_tab         TYPE soli,
         lt_solix_tab        TYPE solix_tab,
         lt_objhead          TYPE soli_tab,
         lv_len              TYPE so_obj_len,
         lv_tansfbin         TYPE sx_boolean,
         lv_xbuffer          TYPE xstring.

  lv_data = p_data.

  REPLACE FIRST OCCURRENCE  OF 'AAAAMMDD' IN gv_url1 WITH lv_data.

  CALL FUNCTION 'HTTP_GET'
    EXPORTING
      absolute_uri                = gv_url1
    IMPORTING
      status_text                 = lv_stat
      status_code                 = lv_lcstatcode
      response_entity_body_length = lv_entbodylen
    TABLES
      response_entity_body        = lt_solix_tab[]
      response_headers            = lt_response_headers
    EXCEPTIONS
      connect_failed              = 1
      timeout                     = 2
      internal_error              = 3
      tcpip_error                 = 4
      data_error                  = 5
      system_failure              = 6
      communication_failure       = 7
      OTHERS                      = 8.
  IF sy-subrc <> 0.
    text1 = '-' ##NO_TEXT.
    text2 = 'Falha na comunica��o.' ##NO_TEXT.
    PERFORM f_registrar_log USING 'E' '047' 'ZFI' text1 text2 text3 text4.
    PERFORM f_save_log.
    PERFORM f_enviar_email.
    MESSAGE s047(zfi) DISPLAY LIKE 'E' WITH text1 text2.
    LEAVE LIST-PROCESSING.
  ENDIF.

  CALL FUNCTION 'SCMS_BINARY_TO_STRING'
    EXPORTING
      input_length = lv_entbodylen
    IMPORTING
      text_buffer  = gv_buffer
    TABLES
      binary_tab   = lt_solix_tab[]
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.

  FIND FIRST OCCURRENCE OF 'not found' IN gv_buffer.
  IF sy-subrc = 0.
    text1 = '-' ##NO_TEXT.
    text2 = 'Arquivo n�o encontrado.' ##NO_TEXT.
    PERFORM f_registrar_log USING 'E' '047' 'ZFI' text1 text2 text3 text4.
    PERFORM f_save_log.
    PERFORM f_enviar_email.
    MESSAGE s047(zfi) DISPLAY LIKE 'E' WITH text1 text2.
    LEAVE LIST-PROCESSING.
  ENDIF.

ENDFORM.                    " F_GET_HTTP_DATA

*&---------------------------------------------------------------------*
*&      Form  F_SPLIT_HTTP_DATA
*&---------------------------------------------------------------------*
FORM f_split_http_data.

  DATA: ti_tcun   TYPE STANDARD TABLE OF tcurr,
        st_tcur   TYPE tcurr.

  REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN gv_buffer WITH '|'.

  SPLIT gv_buffer AT '|' INTO TABLE lt_cotacao_http.

  LOOP AT lt_cotacao_http INTO lw_cotacao_http.

    SPLIT lw_cotacao_http AT ';' INTO lw_cotacao-camp01
                                      lw_cotacao-camp02
                                      lw_cotacao-camp03
                                      lw_cotacao-camp04
                                      lw_cotacao-camp05
                                      lw_cotacao-camp06
                                      lw_cotacao-camp07
                                      lw_cotacao-camp08.

    lw_cotacao-camp01 = lw_cotacao-camp01+6(4) && lw_cotacao-camp01+3(2) && lw_cotacao-camp01+0(2).

    REPLACE ALL OCCURRENCES OF ',' : IN lw_cotacao-camp05 WITH '.',
                                     IN lw_cotacao-camp06 WITH '.',
                                     IN lw_cotacao-camp07 WITH '.',
                                     IN lw_cotacao-camp08 WITH '.'.

    APPEND lw_cotacao TO lt_cotacao.
    CLEAR : lw_cotacao.

  ENDLOOP.

ENDFORM.                    " F_SPLIT_HTTP_DATA

*&---------------------------------------------------------------------*
*&      Form  F_FILL_DATA
*&---------------------------------------------------------------------*
FORM f_fill_data.

  SELECT name low
    INTO TABLE lt_tvarvc
    FROM tvarvc
    WHERE name = gc_taxas.
  IF sy-subrc = 0.
    LOOP AT lt_tvarvc INTO lw_tvarvc.
      r_taxa-sign   = 'I'.
      r_taxa-option = 'EQ'.
      r_taxa-low    = lw_tvarvc-low.
      APPEND r_taxa.
    ENDLOOP.
  ENDIF.

  CLEAR: lt_tvarvc, lw_tvarvc.

  SELECT name low
    INTO TABLE lt_tvarvc
    FROM tvarvc
    WHERE name = gc_cat_taxa.
  IF sy-subrc = 0.
    LOOP AT lt_tvarvc INTO lw_tvarvc.
      r_cat_taxa-sign   = 'I'.
      r_cat_taxa-option = 'EQ'.
      r_cat_taxa-low    = lw_tvarvc-low.
      APPEND r_cat_taxa.

      lw_cat_taxa-kurst = lw_tvarvc-low.
      APPEND lw_cat_taxa TO lt_cat_taxa.
    ENDLOOP.
  ENDIF.

  IF r_cat_taxa[] IS INITIAL OR r_taxa[] IS INITIAL.
    MESSAGE 'Vari�veis STVARV n�o cadastradas. Verifique.' TYPE 'S' DISPLAY LIKE 'E' ##NO_TEXT.
    LEAVE LIST-PROCESSING.
  ENDIF.

  SELECT DISTINCT kurst fcurr tcurr ffact tfact
    INTO TABLE lt_tcurf
    FROM tcurf
    WHERE kurst IN r_cat_taxa
      AND fcurr IN r_taxa
      AND tcurr = 'BRL'.
  IF sy-subrc = 0.
    ##NEEDED.
  ENDIF.

  DELETE lt_cotacao WHERE camp04 NOT IN r_taxa.

  LOOP AT lt_cat_taxa INTO lw_cat_taxa.

    CASE lw_cat_taxa-kurst.
      WHEN 'B'.

        LOOP AT lt_cotacao INTO lw_cotacao.

          READ TABLE lt_tcurf INTO lw_tcurf WITH KEY kurst = 'B'
                                                     fcurr = lw_cotacao-camp04.

          lw_exch_rate-rate_type = lw_cat_taxa-kurst.  " Categoria
          lw_exch_rate-from_curr = lw_cotacao-camp04.  " Moeda
          lw_exch_rate-to_currncy = 'BRL'.
          lw_exch_rate-valid_from = lw_cotacao-camp01. " Data
          lw_exch_rate-exch_rate = lw_cotacao-camp05.  " Campo Compra
          lw_exch_rate-from_factor = lw_tcurf-ffact.
          lw_exch_rate-to_factor = lw_tcurf-tfact.
          lw_exch_rate-exch_rate_v = ' '.
          lw_exch_rate-from_factor_v = lw_tcurf-ffact.
          lw_exch_rate-to_factor_v = lw_tcurf-tfact.

          PERFORM f_chamar_bapi_create USING lw_exch_rate.

          CLEAR lw_exch_rate.

        ENDLOOP.

      WHEN 'G'.

        LOOP AT lt_cotacao INTO lw_cotacao.

          READ TABLE lt_tcurf INTO lw_tcurf WITH KEY kurst = 'G'
                                                     fcurr = lw_cotacao-camp04.

          lw_exch_rate-rate_type = lw_cat_taxa-kurst.  " Categoria
          lw_exch_rate-from_curr = lw_cotacao-camp04.  " Moeda
          lw_exch_rate-to_currncy = 'BRL'.
          lw_exch_rate-valid_from = lw_cotacao-camp01. " Data
          lw_exch_rate-exch_rate = lw_cotacao-camp06.  " Campo Venda
          lw_exch_rate-from_factor = lw_tcurf-ffact.
          lw_exch_rate-to_factor = lw_tcurf-tfact.
          lw_exch_rate-exch_rate_v = ' '.
          lw_exch_rate-from_factor_v = lw_tcurf-ffact.
          lw_exch_rate-to_factor_v = lw_tcurf-tfact.

          PERFORM f_chamar_bapi_create USING lw_exch_rate.

          CLEAR lw_exch_rate.

        ENDLOOP.

      WHEN 'M'.

        LOOP AT lt_cotacao INTO lw_cotacao.

          READ TABLE lt_tcurf INTO lw_tcurf WITH KEY kurst = 'M'
                                                     fcurr = lw_cotacao-camp04.

          lw_exch_rate-rate_type = lw_cat_taxa-kurst.  " Categoria
          lw_exch_rate-from_curr = lw_cotacao-camp04.  " Moeda
          lw_exch_rate-to_currncy = 'BRL'.
          lw_exch_rate-valid_from = lw_cotacao-camp01. " Data
          lw_exch_rate-exch_rate = ( lw_cotacao-camp05 + lw_cotacao-camp06 ) / 2. " M�dia Compra/Venda
          lw_exch_rate-from_factor = lw_tcurf-ffact.
          lw_exch_rate-to_factor = lw_tcurf-tfact.
          lw_exch_rate-exch_rate_v = ' '.
          lw_exch_rate-from_factor_v = lw_tcurf-ffact.
          lw_exch_rate-to_factor_v = lw_tcurf-tfact.

          PERFORM f_chamar_bapi_create USING lw_exch_rate.

          CLEAR lw_exch_rate.

        ENDLOOP.

      WHEN OTHERS.
    ENDCASE.

  ENDLOOP.

ENDFORM.                    " F_FILL_DATA

*&---------------------------------------------------------------------*
*&      Form  F_REGISTRAR_LOG
*&---------------------------------------------------------------------*
FORM f_registrar_log USING flv_tipo
                           flv_num
                           flv_classe
                           flv_text1
                           flv_text2
                           flv_text3
                           flv_text4.

  DATA: msg TYPE bal_s_msg.

  msg-msgty = flv_tipo.
  msg-msgid = flv_classe.
  msg-msgno = flv_num.
  msg-msgv1 = flv_text1.
  msg-msgv2 = flv_text2.
  msg-msgv3 = flv_text3.
  msg-msgv4 = flv_text4.

  " Adicionar mensagens no log
  CALL FUNCTION 'BAL_LOG_MSG_ADD'
    EXPORTING
      i_log_handle     = lw_log_handle
      i_s_msg          = msg
    EXCEPTIONS
      log_not_found    = 1
      msg_inconsistent = 2
      log_is_full      = 3
      OTHERS           = 4.
  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.

  INSERT lw_log_handle INTO TABLE lt_log_handle.

*  " Salvar log no banco de dados
*  CALL FUNCTION 'BAL_DB_SAVE'
*    EXPORTING
*      i_t_log_handle   = lt_log_handle
*    EXCEPTIONS
*      log_not_found    = 1
*      save_not_allowed = 2
*      numbering_error  = 3
*      OTHERS           = 4.
*  IF sy-subrc <> 0.
** Implement suitable error handling here
*  ENDIF.

  "CLEAR: lw_log_handle. ", lt_log_handle.

ENDFORM.                    " F_REGISTRAR_LOG

*&---------------------------------------------------------------------*
*&      Form  F_CHAMAR_BAPI_CREATE
*&---------------------------------------------------------------------*
* Realiza a atualiza��o no banco de dados
*----------------------------------------------------------------------*
FORM f_chamar_bapi_create USING p_exch_rate.

  DATA: lw_bapiret2 TYPE bapiret2.

  DATA: flw_exch_rate LIKE LINE OF lt_exch_rate.

  flw_exch_rate = p_exch_rate.

  CALL FUNCTION 'BAPI_EXCHANGERATE_CREATE'
    EXPORTING
      exch_rate = flw_exch_rate
      upd_allow = 'X'
      chg_fixed = ' '
      dev_allow = '000'
    IMPORTING
      return    = lw_bapiret2.
  IF sy-subrc = 0.

    IF lw_bapiret2-type = 'E'.
      text1 = '-'.
      CONCATENATE 'Objeto bloqueado pelo usu�rio' lw_bapiret2-message_v1 INTO text2 SEPARATED BY space.
      PERFORM f_registrar_log USING 'E' '047' 'ZFI' text1 text2 text3 text4.
      PERFORM f_save_log.
      MESSAGE s047(zfi) DISPLAY LIKE 'E' WITH text1 text2.
      LEAVE LIST-PROCESSING.
    ENDIF.

    COMMIT WORK.
    text1 = flw_exch_rate-rate_type.
    text2 = flw_exch_rate-from_curr.
    text3 = flw_exch_rate-to_currncy.
    CONCATENATE flw_exch_rate-valid_from+6(2) flw_exch_rate-valid_from+4(2) flw_exch_rate-valid_from+0(4)
           INTO text4 SEPARATED BY '.'.

    PERFORM f_registrar_log USING 'S' '046' 'ZFI' text1 text2 text3 text4.
    CLEAR: text1, text2, text3, text4.

  ELSE.

    text1 = flw_exch_rate-rate_type.
    text2 = flw_exch_rate-from_curr.
    text3 = flw_exch_rate-to_currncy.
    CONCATENATE flw_exch_rate-valid_from+6(2) flw_exch_rate-valid_from+4(2) flw_exch_rate-valid_from+0(4)
           INTO text4 SEPARATED BY '.'.

    PERFORM f_registrar_log USING 'E' '047' 'ZFI' text1 text2 text3 text4.
    PERFORM f_save_log.
    PERFORM f_enviar_email.
    CLEAR: text1, text2, text3, text4.

  ENDIF.

ENDFORM.                    " F_CHAMAR_BAPI_CREATE

*&---------------------------------------------------------------------*
*&      Form  F_ENVIAR_EMAIL
*&---------------------------------------------------------------------*
FORM f_enviar_email.

  " Preparando email
  CLASS cl_bcs DEFINITION LOAD.
  DATA: lo_send_request TYPE REF TO cl_bcs VALUE IS INITIAL,
        lo_document     TYPE REF TO cl_document_bcs VALUE IS INITIAL,  " Document object
        lo_sender       TYPE REF TO if_sender_bcs VALUE IS INITIAL,    " Remetente
        lo_recipient    TYPE REF TO if_recipient_bcs VALUE IS INITIAL. " Destinat�rio

  DATA: lt_text TYPE bcsy_text,       " Tabela interna para o corpo do email
        lw_text LIKE LINE OF lt_text, " Estrutura para o corpo do email
        lt_destinatario TYPE bcsy_smtpa,
        lw_destinatario TYPE ad_smtpadr.

  " Classes de exce��o
  DATA: lx_document_bcs TYPE REF TO cx_document_bcs,
        lx_address_bcs  TYPE REF TO cx_address_bcs,
        lx_send_req_bcs TYPE REF TO cx_send_req_bcs.

  " Vari�veis locais
  DATA: lv_data(10) TYPE c.

  CONCATENATE sy-datum+6(2) sy-datum+4(2) sy-datum+0(4) INTO lv_data SEPARATED BY '/'.

  SELECT low
    INTO TABLE lt_destinatario
    FROM tvarvc
    WHERE name = gc_alerta_email.
  IF sy-subrc = 0.
    ##NEEDED.
  ENDIF.

  TRY.
      lo_send_request = cl_bcs=>create_persistent( ).

      " Set body
      CONCATENATE 'Houve erro na importa��o da taxa de c�mbio no SAP na data' lv_data '. Favor verificar e corrigir a transa��o OB08.'
                  'Para informa��es mais detalhadas consulte a transa��o SLG1: Objeto: Z_TAXA_CAMBIO, Subobjeto: ZTAXA.'
                   INTO lw_text-line SEPARATED BY space.
      APPEND lw_text TO lt_text.
      CLEAR: lw_text, lv_data.

      lo_document = cl_document_bcs=>create_document(                            " Cria��o do documento
                                     i_type = 'TXT'                              " Tipo do documento... XLS, TXT, HTML
                                     i_text =  lt_text                           " Tabela interna para o corpo do email
                                     i_subject = 'Erro na importa��o autom�tica das taxas de c�mbio' ). " Assunto do email

      lo_send_request->set_document( lo_document ). " Pass the document to send request
    CATCH cx_send_req_bcs INTO lx_send_req_bcs.
    CATCH cx_document_bcs INTO lx_document_bcs.
  ENDTRY.

  TRY.
      lo_sender = cl_sapuser_bcs=>create( sy-uname ). " Remetente � o usu�rio logado
      lo_send_request->set_sender(
        EXPORTING
          i_sender = lo_sender ).
    CATCH cx_address_bcs INTO lx_address_bcs.
    CATCH cx_send_req_bcs INTO lx_send_req_bcs.
  ENDTRY.

  LOOP AT lt_destinatario INTO lw_destinatario.

    TRY.
        lo_recipient = cl_cam_address_bcs=>create_internet_address( lw_destinatario ).

        lo_send_request->add_recipient(
            EXPORTING
              i_recipient = lo_recipient
              i_express = 'X' ).
      CATCH cx_send_req_bcs INTO lx_send_req_bcs.
      CATCH cx_address_bcs INTO lx_address_bcs.
    ENDTRY.

  ENDLOOP.

  TRY.
      CALL METHOD lo_send_request->set_send_immediately
        EXPORTING
          i_send_immediately = 'X'. " Envio imediato
    CATCH cx_send_req_bcs INTO lx_send_req_bcs .
  ENDTRY.

  " Enviando Email
  TRY.
      lo_send_request->send(
      EXPORTING
        i_with_error_screen = 'X' ).
      COMMIT WORK.
      IF sy-subrc = 0. " Email enviado com sucesso
        "WRITE : 'Email enviado com sucesso.'.
      ENDIF.
    CATCH cx_send_req_bcs INTO lx_send_req_bcs .
  ENDTRY.

ENDFORM.                    " F_ENVIAR_EMAIL

*&---------------------------------------------------------------------*
*&      Form  F_SAVE_LOG
*&---------------------------------------------------------------------*
* Salva os registros de logs no banco de dados
*----------------------------------------------------------------------*
FORM f_save_log.

  " Salvar log no banco de dados
  CALL FUNCTION 'BAL_DB_SAVE'
    EXPORTING
      i_t_log_handle   = lt_log_handle
    EXCEPTIONS
      log_not_found    = 1
      save_not_allowed = 2
      numbering_error  = 3
      OTHERS           = 4.
  IF sy-subrc <> 0.
* Implement suitable error handling here
  ENDIF.

ENDFORM.                    " F_SAVE_LOG