FUNCTION zf_email_report_html.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(REPORT) TYPE  PROGRAMM
*"     REFERENCE(VARIANT) TYPE  RALDB_VARI OPTIONAL
*"     REFERENCE(SUBJECT) TYPE  SO_OBJ_DES
*"  TABLES
*"      TABLE_SENDERS TYPE  BCSY_SMTPA
*"      TABLE_PARAMS STRUCTURE  RSPARAMS OPTIONAL
*"      TEXT_BODY TYPE  BCSY_TEXT OPTIONAL
*"      RETURN TYPE  TABLE_OF_STRINGS OPTIONAL
*"  EXCEPTIONS
*"      ENTER_RECIPIENT
*"      ERROR_IMPORT_DATA
*"      ERROR_FORMAT_HTML
*"----------------------------------------------------------------------

  " Preparando email
  CLASS cl_bcs DEFINITION LOAD.
  DATA: lo_send_request TYPE REF TO cl_bcs VALUE IS INITIAL,
        lo_document     TYPE REF TO cl_document_bcs VALUE IS INITIAL,  " Document object
        lo_sender       TYPE REF TO if_sender_bcs VALUE IS INITIAL,    " Remetente
        lo_recipient    TYPE REF TO if_recipient_bcs VALUE IS INITIAL. " Destinatário

  DATA: lt_text TYPE bcsy_text,       " Tabela interna para o corpo do email
        lw_text LIKE LINE OF lt_text, " Estrutura para o corpo do email
        lt_destinatario TYPE bcsy_smtpa,
        lw_destinatario TYPE ad_smtpadr.

  " Parâmetros
  DATA: lt_params TYPE TABLE OF rsparams,
        lw_params TYPE rsparams.

  " Dados para o anexo
  DATA: lv_data(10)   TYPE c,
        lv_attsubject TYPE sood-objdes.

  " Classes de exceção
  DATA: lx_document_bcs TYPE REF TO cx_document_bcs,
        lx_address_bcs  TYPE REF TO cx_address_bcs,
        lx_send_req_bcs TYPE REF TO cx_send_req_bcs,
        lx_bcs_exception TYPE REF TO cx_bcs.

  " Variáveis globais
  DATA: lv_report      TYPE programm,
        lv_variante    TYPE raldb_vari,
        lv_lines       TYPE i,
        lv_bcs_message TYPE string,
        lv_conlengths  TYPE so_obj_len ,       " Para calcular o tamanho do arquivo html
        lt_listobject  TYPE TABLE OF abaplist, " Retorno dos dados da memória
        lt_html        TYPE TABLE OF w3html.   " Html container

*&--------------------------------------------------------------------*
*& START
*&--------------------------------------------------------------------*
  CLEAR: lv_report, lv_variante,
         lt_listobject[], lt_html[],
         lv_conlengths.

  lv_report = report.
  lv_variante = variant.

  MOVE-CORRESPONDING table_senders[] TO lt_destinatario[].
  MOVE-CORRESPONDING table_params[] TO lt_params[].

  TRANSLATE lv_report TO UPPER CASE.
  TRANSLATE lv_variante TO UPPER CASE.

  IF table_senders[] IS NOT INITIAL.

    IF lt_params[] IS NOT INITIAL.
      " Executa baseado em tabelas de parâmetros
      SUBMIT (lv_report) WITH SELECTION-TABLE lt_params EXPORTING LIST TO MEMORY AND RETURN.
    ELSE.
      " Executa baseado em uma vaiante de tela
      SUBMIT (lv_report) USING SELECTION-SET lv_variante EXPORTING LIST TO MEMORY AND RETURN.
    ENDIF.

    CALL FUNCTION 'LIST_FROM_MEMORY'
      TABLES
        listobject = lt_listobject
      EXCEPTIONS
        not_found  = 1
        OTHERS     = 2.
    IF sy-subrc <> 0.
      RAISE error_import_data.
    ENDIF.

    CALL FUNCTION 'WWW_HTML_FROM_LISTOBJECT'
      EXPORTING
        report_name = lv_report
      TABLES
        html        = lt_html
        listobject  = lt_listobject.
    IF sy-subrc <> 0.
      APPEND 'Erro ao formatar dados html' TO return ##NO_TEXT.
      RAISE error_format_html.
    ENDIF.

*--- Recuperando tamanho do documento html
    DESCRIBE TABLE lt_html LINES lv_lines.

    CLEAR: lo_send_request.

*--- Texto para o corpo do email
    MOVE-CORRESPONDING text_body[] TO lt_text[].

    TRY.

        lo_send_request = cl_bcs=>create_persistent( ).

        lo_document = cl_document_bcs=>create_document(   " Criação do documento
                      i_type = 'HTM'                      " Tipo do documento... XLS, TXT, HTML
                      i_text = lt_text                    " Tabela interna para o corpo do email
                      i_length = lv_conlengths            " Tamanho do documento
                      i_subject = subject ).              " Assunto do email

*--- Adicionando documento para enviar pedido
        CALL METHOD lo_send_request->set_document( lo_document ).

*--- Adicionando arquivo em anexo no email
        CLEAR: lv_attsubject, lv_data.
        CONCATENATE sy-datum+6(2) sy-datum+4(2) sy-datum+0(4) INTO lv_data.
        CONCATENATE lv_report '-' lv_data INTO lv_attsubject SEPARATED BY space. " Nome do arquivo em anexo

        lo_document->add_attachment( EXPORTING
                                   i_attachment_type = 'HTM'
                                   i_attachment_subject = lv_attsubject
                                   i_att_content_text = lt_html ).

        lo_sender = cl_sapuser_bcs=>create( sy-uname ). " Remetente é o usuário logado
        lo_send_request->set_sender(
          EXPORTING
            i_sender = lo_sender ).

*--- Enviando para todos os destinatários
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

        CALL METHOD lo_send_request->set_send_immediately
          EXPORTING
            i_send_immediately = 'X'. " Envio imediato

        " Enviando Email
        lo_send_request->send(
        EXPORTING
          i_with_error_screen = 'X' ).
        COMMIT WORK.
        IF sy-subrc = 0. " Email enviado com sucesso
          APPEND 'Email enviado com sucesso.' TO return ##NO_TEXT.
        ENDIF.

      CATCH cx_bcs INTO lx_bcs_exception.
        lv_bcs_message = lx_bcs_exception->get_text( ).
        APPEND lv_bcs_message TO return .
        EXIT.

    ENDTRY.

  ELSE.
    APPEND 'Informe pelo menos um destinatário' TO return ##NO_TEXT.
    RAISE enter_recipient.
  ENDIF.

ENDFUNCTION.