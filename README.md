# Phyton-code
Phyton code for automation at SAP GUI

# Este código é um EXEMPLO EDUCACIONAL.
    Valores reais de chaves, arquivos e sessões foram removidos
    para permitir publicação segura no GitHub.


from SapGuiLib import SapGuiBase, clipboardPreserverContextManager, sapToDataframe
from pyperclip import copy, paste
from datetime import datetime, timedelta
from pyautogui import confirm
from math import floor
from time import sleep

from math import floor

def converte_lista_para_string_compacta(lista, max_itens=6):
    """
    Converte uma lista em uma string compacta.

    - Se a lista tiver até `max_itens`, mostra todos os elementos.
    - Caso contrário, mostra os primeiros e últimos elementos,
      substituindo o meio por '...'.

    Exemplo:
    [1, 2, 3, 4, 5, 6, 7, 8] -> [1, 2, 3, ..., 6, 7, 8]
    """

    texto = "["

    # Caso a lista seja pequena
    if len(lista) <= max_itens:
        for elemento in lista:
            texto += f"{elemento}, "

    # Caso a lista seja grande
    else:
        metade = floor(max_itens / 2)

        # Primeiros elementos
        for i in range(metade):
            texto += f"{lista[i]}, "

        texto += "..., "

        # Últimos elementos
        for i in range(metade):
            texto += f"{lista[-metade + i]}, "

    # Remove a última vírgula e espaço
    texto = texto[:-2] + "]"

    return texto

    from datetime import datetime

    sap = SapGuiBase(permiteReutilizar=False)
sap.sapLogin(modo="pyautogui")

def executar_fp08m(session, lista_documentos):
    """
    Executa a transação FP08M no SAP para estorno de documentos.


    with open(caminho_arquivo_chaves, "r") as arquivo:
        lista_chaves = [
            linha.strip()
            for linha in arquivo.readlines()
            if linha.strip()
        ]

    # Geração de chaves genéricas de fallback (exemplo)
    # O padrão real foi removido
    lista_chaves += [
        f"CHAVE_PADRAO_EXEMPLO_{i:02d}"
        for i in range(1, 51)
    ]

    indice_chave = 0

    # ------------------------------------------------------------------
    # ACESSO À TRANSAÇÃO FP08M
    # ------------------------------------------------------------------
    session.findById("wnd[0]/tbar[0]/okcd").text = "fp08m"
    session.findById("wnd[0]").sendVKey(0)

    # Chave fixa de seleção (valor real removido)
    session.findById(
        "wnd[0]/usr/ctxtSC_FIKEY-LOW"
    ).text = "<CHAVE_FIXA_EXEMPLO>"

    # Chave de reconciliação utilizada no estorno
    session.findById(
        "wnd[0]/usr/ctxtP_FIKEY"
    ).text = lista_chaves[indice_chave]

    # ------------------------------------------------------------------
    # INSERÇÃO DA LISTA DE DOCUMENTOS
    # ------------------------------------------------------------------
    # Os documentos são inseridos via clipboard
    # (números reais removidos)
    with clipboardPreserverContextManager():
        copy("\r\n".join(str(doc) for doc in lista_documentos))
        session.findById(
            "wnd[0]/usr/btn%_SC_OPBEL_%_APP_%-VALU_PUSH"
        ).press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # ------------------------------------------------------------------
    # CONFIGURAÇÕES DE ESTORNO
    # ------------------------------------------------------------------
    # Marca a opção "Estornar documentos"
    session.findById("wnd[0]/usr/radP_RTSTOR").select()

    # Define a data de estorno como a data atual
    session.findById(
        "wnd[0]/usr/ctxtP_STODT"
    ).text = datetime.now().strftime("%d.%m.%Y")

    # Limpa o campo de filtro de chave
    session.findById(
        "wnd[0]/usr/ctxtSC_FIKEY-LOW"
    ).text = ""

    # Executa a transação (botão do relógio)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # ------------------------------------------------------------------
    # TRATAMENTO DE ERRO: CHAVE JÁ FECHADA
    # ------------------------------------------------------------------
    while True:
        try:
            mensagem_status = session.findById("wnd[0]/sbar").text
            print(mensagem_status)

            # Caso a chave esteja fechada, tenta a próxima
            if "já fechada" in mensagem_status:
                indice_chave += 1

                if indice_chave >= len(lista_chaves):
                    # Nenhuma chave disponível
                    session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    raise Exception("Nenhuma chave de reconciliação disponível.")

                # Define nova chave e tenta novamente
                session.findById(
                    "wnd[0]/usr/ctxtP_FIKEY"
                ).text = lista_chaves[indice_chave]

                session.findById("wnd[0]/tbar[1]/btn[8]").press()

            else:
                # Nenhum erro encontrado
                break

        except Exception:
            # Em caso de erro inesperado, sai do loop
            break

    # ------------------------------------------------------------------
    # CONFIRMAÇÃO FINAL
    # ------------------------------------------------------------------
    session.findById("wnd[1]/usr/btnBUTTON_1").press()

    # Retorna para a tela inicial
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

from datetime import datetime

def tratar_linha_fp(session, lista_chaves, dados_preenchidos=False):
    """
    Trata uma linha da lista de documentos no SAP,
    realizando tentativa de estorno com múltiplas chaves.

    Todas as chaves e identificadores reais foram removidos
    para permitir publicação segura no GitHub.
    """

    # Se não houver mais chaves disponíveis, encerra o processo
    if not lista_chaves:
        raise Exception("Nenhuma chave de reconciliação disponível.")

    # Preenchimento inicial dos dados de estorno
    if not dados_preenchidos:
        # Seleciona a linha atual
        session.findById("wnd[0]/usr/lbl[XXX,YY]").setFocus()
        session.findById("wnd[0]/usr/lbl[XXX,YY]").showContextMenu()

        # Seleciona a opção de estorno (ID real removido)
        session.findById("wnd[0]/usr").selectContextMenuItem("<OPCAO_ESTORNO>")

        # Define a data de estorno como a data atual
        session.findById(
            "wnd[0]/usr/ctxtRFK00-STODT"
        ).text = datetime.now().strftime("%d.%m.%Y")

    # Insere a chave de reconciliação
    session.findById(
        "wnd[0]/usr/ctxtRFK00-FIKEY"
    ).text = lista_chaves[0]

    # Salva (botão disquete)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

    # Lê mensagem de status do SAP
    mensagem_status = session.findById("wnd[0]/sbar").text

    # Caso a linha esteja bloqueada
    if "motivo de bloqueio" in mensagem_status:
        print("Linha bloqueada — estorno não permitido.")
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

    # Caso a chave esteja fechada, tenta a próxima
    elif "já fechada" in mensagem_status:
        return tratar_linha_fp(
            session,
            lista_chaves[1:],
            dados_preenchidos=True
        )

    # Caso não haja erro, o estorno foi realizado
    else:
        pass


def executar_fpl9(session, lista_conta_contratos, lista_chaves):
    """
    Executa a transação FPL9 no SAP para tratar múltiplas
    contas contrato e realizar estornos linha a linha.

    Código sanitizado para publicação no GitHub.
    """

    # ------------------------------------------------------------------
    # ACESSO À TRANSAÇÃO FPL9
    # ------------------------------------------------------------------
    session.findById("wnd[0]/tbar[0]/okcd").text = "fpl9"
    session.findById("wnd[0]").sendVKey(0)

    # Define tipo de lista (valor real removido)
    session.findById(
        "wnd[0]/usr/cmbFKKL1-LSTYP"
    ).key = "<TIPO_LISTA_EXEMPLO>"

    # ------------------------------------------------------------------
    # INSERÇÃO DAS CONTAS CONTRATO
    # ------------------------------------------------------------------
    with clipboardPreserverContextManager():
        copy("\r\n".join(str(ct) for ct in lista_conta_contratos))
        session.findById("wnd[0]/usr/btnPUSH_RVKT").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executa (botão verde)
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Expande a visão da lista
    session.findById("wnd[0]/tbar[1]/btn[39]").press()
    session.findById("wnd[0]/tbar[1]/btn[39]").press()

    # Interação manual do usuário (exemplo)
    resposta = confirm(
        "Filtre os dados e clique em OK para continuar",
        buttons=["OK", "Cancelar", "Nenhum caso"]
    )

    if resposta == "Cancelar":
        raise Exception("Execução cancelada pelo usuário.")

    if resposta == "Nenhum caso":
        # Retorna para a tela inicial
        for _ in range(10):
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return

    # ------------------------------------------------------------------
    # PERCORRER LINHAS DA LISTA
    # ------------------------------------------------------------------
    posicao_scroll = 0
    session.findById("wnd[0]/usr").verticalScrollbar.position = posicao_scroll

    while True:
        try:
            # Tenta ler a descrição da operação atual
            session.findById("wnd[0]/usr/lbl[XX,YY]").text
        except Exception:
            # Não há mais linhas
            break

        # Trata a linha atual
        tratar_linha_fp(session, lista_chaves)

        # Avança o scroll
        posicao_scroll += 1
        session.findById("wnd[0]/usr").verticalScrollbar.position = posicao_scroll

    # ------------------------------------------------------------------
    # RETORNO À TELA INICIAL
    # ------------------------------------------------------------------
    for _ in range(4):
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        
# Demonstração de como acessar propriedades de um elemento SAP GUI
# session.findById("<ELEMENT_ID>").type

from clipboard import copy, paste  # Exemplo: biblioteca para manipulação de clipboard

def executar_se16n(session, lista_conta_contratos):
    """
    Função para consultar dados no SAP usando a transação SE16N.

    Todos os IDs específicos de telas e tabelas foram removidos
    ou substituídos por placeholders para publicação segura.
    """

    # ------------------------------------------------------------------
    # ACESSO À TRANSAÇÃO SE16N
    # ------------------------------------------------------------------
    session.findById("wnd[0]/tbar[0]/okcd").text = "SE16N"
    session.findById("wnd[0]").sendVKey(0)

    # Definir tabela para consulta (nome real removido)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "<TABELA_EXEMPLO>"
    session.findById("wnd[0]").sendVKey(0)

    # Desmarcar todos os campos selecionáveis
    session.findById("wnd[0]/tbar[1]/btn[18]").press()

    # Selecionar apenas os campos de interesse (IDs removidos)
    session.findById("<CHECKBOX_CAMPO1>").selected = True  # exemplo: campo 1
    session.findById("<CHECKBOX_CAMPO2>").selected = True  # exemplo: campo 2

    # ------------------------------------------------------------------
    # INSERÇÃO DOS DADOS (clipboard)
    # ------------------------------------------------------------------
    with clipboardPreserverContextManager():
        copy("\r\n".join(str(item) for item in lista_conta_contratos))

        # Foco e envio dos dados via botão (ID real removido)
        session.findById("<BTN_INSERIR>").setFocus()
        session.findById("<BTN_INSERIR>").press()

        # Confirmar popups da transação (IDs genéricos)
        session.findById("<BTN_POPUP_OK>").press()
        session.findById("<BTN_POPUP_FECHAR>").press()

    # Limpa limite de linhas (campo real removido)
    session.findById("<CAMPO_MAX_LINHAS>").text = ""

    # Executa a consulta
    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # botão RUN

    # ------------------------------------------------------------------
    # EXPORTAÇÃO DOS RESULTADOS
    # ------------------------------------------------------------------
    # Abrir menu de exportação (IDs removidos)
    session.findById("<BTN_EXPORTAR>").pressToolbarContextButton("&MB_EXPORT")
    session.findById("<MENU_EXPORTAR>").selectContextMenuItem("&PC")

    # Selecionar opções de exportação via clipboard
    with clipboardPreserverContextManager():
        session.findById("<RADIO_SELECAO>").select()
        session.findById("<RADIO_SELECAO>").setFocus()
        session.findById("<BTN_CONFIRMAR>").press()
        resultado = paste()  # Conteúdo exportado para clipboard

    # Volta para tela inicial
    for _ in range(2):
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

    # Converte resultado para DataFrame (função de exemplo)
    return sapToDataframe(resultado)
def executar_zsvc20(session, lista_pns):
    """
    Função para consultar dados no SAP usando a transação ZSVC20.

    Todos os IDs específicos de telas, botões e campos foram removidos
    ou substituídos por placeholders para publicação segura.
    """

    # ------------------------------------------------------------------
    # ACESSO À TRANSAÇÃO ZSVC20
    # ------------------------------------------------------------------
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZSVC20"
    session.findById("wnd[0]").sendVKey(0)

    # ------------------------------------------------------------------
    # INSERÇÃO DOS DADOS (clipboard)
    # ------------------------------------------------------------------
    with clipboardPreserverContextManager():
        copy("\r\n".join(str(pn) for pn in lista_pns))

        # Botões de inserção e confirmação (IDs genéricos)
        session.findById("<BTN_INSERIR_PNS>").press()
        session.findById("<BTN_POPUP_OK>").press()
        session.findById("<BTN_POPUP_FECHAR>").press()

    # ------------------------------------------------------------------
    # DEFINIR PERÍODO DE CONSULTA
    # ------------------------------------------------------------------
    lim_max = datetime.now()
    lim_min = datetime.now() - timedelta(days=60)

    session.findById("<CAMPO_DATA_INICIAL>").text = lim_min.strftime("%d.%m.%Y")
    session.findById("<CAMPO_DATA_FINAL>").text = lim_max.strftime("%d.%m.%Y")

    # Executa a consulta
    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # botão executar

    # ------------------------------------------------------------------
    # FILTRO DE DANO (exemplo: código "0090")
    # ------------------------------------------------------------------
    session.findById("<TABELA_RESULTADOS>").setCurrentCell(-1, "FECOD")
    session.findById("<TABELA_RESULTADOS>").selectColumn("FECOD")
    session.findById("<TABELA_RESULTADOS>").contextMenu()
    session.findById("<TABELA_RESULTADOS>").selectContextMenuItem("&FILTER")
    session.findById("<CAMPO_FILTRO_DANO>").text = "0090"
    session.findById("<BTN_CONFIRMAR_FILTRO>").press()

    # ------------------------------------------------------------------
    # EXPORTAÇÃO DOS RESULTADOS
    # ------------------------------------------------------------------
    with clipboardPreserverContextManager():
        session.findById("<BTN_EXPORTAR>").pressToolbarContextButton("&MB_EXPORT")
        session.findById("<MENU_EXPORTAR>").selectContextMenuItem("&PC")
        session.findById("<RADIO_EXPORTAR>").select()
        session.findById("<RADIO_EXPORTAR>").setFocus()
        session.findById("<BTN_CONFIRMAR_EXPORT>").press()

        resultado = paste()  # Dados exportados para clipboard

    # Retorna à tela inicial
    for _ in range(2):
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

    # Converte resultado para DataFrame (função de exemplo)
    return sapToDataframe(resultado, linhaTitulo=4, primeiraLinhaDados=6)

    def processar_iw52(session, num_nota_restituicao, situacao_mensagem):
    """
    Função para processar uma nota de restituição na transação IW52 do SAP.
    
    Todos os IDs específicos de campos e abas foram substituídos por placeholders.
    """

    # ------------------------------------------------------------------
    # ACESSO À TRANSAÇÃO IW52
    # ------------------------------------------------------------------
    session.findById("wnd[0]/tbar[0]/okcd").text = "IW52"
    session.findById("wnd[0]").sendVKey(0)

    # Inserir número da nota de restituição
    session.findById("<CAMPO_NUM_NOTA>").text = str(num_nota_restituicao)

    # Clicar em "Nota" (ID real removido)
    session.findById("<BTN_NOTA>").press()

    # ------------------------------------------------------------------
    # INSERÇÃO DO TEXTO DE MENSAGEM
    # ------------------------------------------------------------------
    texto = ""
    if situacao_mensagem == "50006":
        texto = "MENSAGEM DE TEXTO PARA INSERIR NA NOTA DE SERVIÇO"
    elif situacao_mensagem == "50021":
        texto = "MENSAGEM DE TEXTO PARA INSERIR NA NOTA DE SERVIÇO"

    # Inserir o texto na aba de comentário (ID genérico)
    session.findById("<CAMPO_TEXTO>").text = texto

    # ------------------------------------------------------------------
    # SETAR STATUS DA NOTA PARA IMPR
    # ------------------------------------------------------------------
    session.findById("<BTN_STATUS>").press()
    aux_scroll = 0
    while True:
        session.findById("<TABELA_STATUS>").verticalScrollbar.position = aux_scroll
        status = session.findById("<TABELA_STATUS>/txtSTATUS[2,0]").text

        if status == "IMPR":
            break
        aux_scroll += 1

    # Selecionar e confirmar o status (IDs genéricos)
    session.findById("<RADIO_STATUS>").selected = True
    session.findById("<RADIO_STATUS>").setFocus()
    session.findById("<BTN_CONFIRMAR_STATUS>").press()

    # ------------------------------------------------------------------
    # CONFIRMAÇÕES FINAIS
    # ------------------------------------------------------------------
    session.findById("<BTN_BANDEIRA>").press()  # bandeira
    session.findById("<BTN_CONFIRMAR_DATA>").press()
    session.findById("<BTN_SIM>").press()
    
    # Volta para tela inicial
    session.findById("wnd[0]/tbar[0]/btn[3]").press()



    

