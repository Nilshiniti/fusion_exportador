# Exportador de arquivos do Fusion para o computador.
#
# Fluxo:
#   1. run() abre a janela (hub, projetos, formatos).
#   2. Ao clicar em Continuar, o execute guarda as escolhas em PEDIDO.
#   3. O destroy dispara um evento próprio, porque a API não aceita fechar documentos
#      dentro de eventos de comando (ver Document.close na API).
#   4. O evento próprio chama exportar_tudo(), que faz o trabalho.
#
# O script só LÊ da nuvem. As únicas chamadas que tocam em documentos são Documents.open
# e Document.close(False). Nenhuma salva, apaga, move ou renomeia nada na nuvem.
#
# Ajustes que podem ser mudados à mão estão logo abaixo.

import csv
import datetime
import os
import sys
import traceback

import adsk.core
import adsk.fusion

# ---------- Ajustes ----------
MAX_PASTA = 40            # tamanho máximo do nome de cada pasta (hub, projeto, subpastas)
MAX_ARQUIVO = 80          # tamanho máximo do nome de cada arquivo, sem a extensão
LIMITE_CAMINHO = 250      # o Windows para em 259; a margem evita problemas no Fusion
ABRIR_VISIVEL = True      # True: cada arquivo aparece aberto na tela enquanto exporta

# ---------- Nomes internos ----------
ID_COMANDO = "NilExportadorFusion_comando"
ID_EVENTO = "NilExportadorFusion_evento"
TITULO = "Exportador do Fusion"
EXTENSOES_FUSION = ("f3d", "f3z")
EXTENSOES_SAIDA = (".f3d", ".f3z", ".step", ".stl")
PROIBIDOS = '<>:"/\\|?*'
RESERVADOS = ("CON", "PRN", "AUX", "NUL",
              "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
              "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
COLUNAS = ["data_hora", "hub", "projeto", "pasta_nuvem", "arquivo", "id", "versao", "extensao_nuvem",
           "formato", "status", "caminho", "base_local", "motivo"]

# ---------- Estado do script ----------
HANDLERS = []             # o Fusion exige manter os handlers vivos enquanto o script roda
HUBS = []                 # hubs listados na janela, na mesma ordem da lista
PROJETOS = []             # projetos do hub escolhido, na mesma ordem da lista
PEDIDO = None             # escolhas da janela; None significa que a janela foi cancelada
ARQUIVO_LOG = None
ARQUIVO_CSV = None
LINHAS_PENDENTES = []     # linhas que não puderam ir ao CSV (por exemplo, aberto no Excel)


def limpar_nome(nome, tamanho):
    # Mantém acentos, hífens e o restante; troca só o que o Windows proíbe.
    saida = ""
    for letra in nome:
        if letra in PROIBIDOS or ord(letra) < 32:
            saida = saida + "_"
        else:
            saida = saida + letra
    saida = saida[:tamanho].rstrip(" .")
    if saida == "":
        saida = "sem_nome"
    if saida.split(".")[0].rstrip().upper() in RESERVADOS:
        saida = "_" + saida
    return saida


def escrever_log(mensagem):
    linha = datetime.datetime.now().strftime("%H:%M:%S") + " " + mensagem
    try:
        adsk.core.Application.get().log(linha)
    except Exception:
        pass
    if ARQUIVO_LOG is not None:
        try:
            with open(ARQUIVO_LOG, "a", encoding="utf-8") as arquivo:
                arquivo.write(linha + "\n")
        except Exception:
            pass


def gravar_linhas_pendentes():
    # Devolve True se tudo foi para o CSV. Se o CSV estiver aberto no Excel, as linhas
    # continuam guardadas e são tentadas de novo na próxima gravação.
    try:
        with open(ARQUIVO_CSV, "a", encoding="utf-8", newline="") as arquivo:
            escritor = csv.DictWriter(arquivo, fieldnames=COLUNAS, delimiter=";")
            for linha in LINHAS_PENDENTES:
                escritor.writerow(linha)
        LINHAS_PENDENTES.clear()
        return True
    except OSError:
        return False


def registrar(item, formato, status, caminho, motivo):
    linha = {
        "data_hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "hub": item["hub"],
        "projeto": item["projeto"],
        "pasta_nuvem": item["pasta_nuvem"],
        "arquivo": item["nome"],
        "id": item["id"],
        "versao": item["versao"],
        "extensao_nuvem": item["extensao"],
        "formato": formato,
        "status": status,
        "caminho": caminho,
        "base_local": item["base"],
        "motivo": motivo.replace("\r", " ").replace("\n", " | ")[:500],
    }
    LINHAS_PENDENTES.append(linha)
    if not gravar_linhas_pendentes():
        escrever_log("AVISO: o CSV está bloqueado; a linha fica guardada e será gravada depois.")
    escrever_log(status.upper() + " | " + formato + " | " + item["nome"] + " | " + caminho + " | " + motivo)


def exportar_tudo(pedido):
    global ARQUIVO_LOG, ARQUIVO_CSV
    app = adsk.core.Application.get()
    ui = app.userInterface
    botoes_sim_nao = adsk.core.MessageBoxButtonTypes.YesNoButtonType
    icone_aviso = adsk.core.MessageBoxIconTypes.WarningIconType

    # PASSO 1: escolher a pasta de destino.
    dialogo = ui.createFolderDialog()
    dialogo.title = "Pasta de destino (use um caminho curto; a mesma pasta retoma a exportação)"
    dialogo.initialDirectory = os.path.expanduser("~")
    if dialogo.showDialog() != adsk.core.DialogResults.DialogOK:
        return
    destino = dialogo.folder

    # PASSO 2: preparar a pasta de controle, o log e o CSV.
    pasta_controle = os.path.join(destino, "_exportador")
    pasta_temp = os.path.join(pasta_controle, "tmp")
    os.makedirs(pasta_temp, exist_ok=True)
    ARQUIVO_LOG = os.path.join(pasta_controle, "log_" + datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S") + ".txt")
    ARQUIVO_CSV = os.path.join(pasta_controle, "registro.csv")
    marcador = os.path.join(pasta_controle, "em_andamento.txt")

    escrever_log("Início. Fusion " + app.version + ", Python " + sys.version.split()[0])
    escrever_log("Licença de uso pessoal: " + str(app.data.personalUseLimits is not None))
    escrever_log("Hub: " + pedido["hub"].name + " | Formatos: " + ", ".join(pedido["formatos"]))
    escrever_log("Destino: " + destino)

    if not os.path.exists(ARQUIVO_CSV):
        with open(ARQUIVO_CSV, "w", encoding="utf-8-sig", newline="") as arquivo:
            csv.DictWriter(arquivo, fieldnames=COLUNAS, delimiter=";").writeheader()

    # O CSV precisa estar livre no início, senão a retomada fica sem registro.
    try:
        with open(ARQUIVO_CSV, "a", encoding="utf-8"):
            pass
    except OSError:
        ui.messageBox("O arquivo abaixo está aberto em outro programa. Feche e rode de novo.\n\n" + ARQUIVO_CSV,
                      TITULO, adsk.core.MessageBoxButtonTypes.OKButtonType, icone_aviso)
        return

    # PASSO 3: ler o que já foi feito em execuções anteriores.
    ultimo_status = {}        # (id, versão, formato) -> linha mais recente do CSV
    base_da_versao = {}       # (id, versão) -> nome local já reservado para aquela versão
    ids_conhecidos = set()    # arquivos que já têm alguma versão registrada
    bases_usadas = set()      # nomes locais já reservados, sem distinguir maiúsculas
    with open(ARQUIVO_CSV, "r", encoding="utf-8-sig", newline="") as arquivo:
        for linha in csv.DictReader(arquivo, delimiter=";"):
            ultimo_status[(linha.get("id"), linha.get("versao"), linha.get("formato"))] = linha
            if linha.get("base_local"):
                base_da_versao[(linha.get("id"), linha.get("versao"))] = linha["base_local"]
                ids_conhecidos.add(linha.get("id"))
                bases_usadas.add(linha["base_local"].casefold())

    # Se o Fusion fechou no meio da execução anterior, o marcador ficou para trás.
    # Esse arquivo é pulado uma vez, para a exportação não travar sempre no mesmo ponto.
    id_a_pular = None
    versao_a_pular = None
    if os.path.exists(marcador):
        with open(marcador, "r", encoding="utf-8") as arquivo:
            partes = arquivo.read().split("\t")
        if len(partes) >= 2:
            id_a_pular = partes[0]
            versao_a_pular = partes[1]
        os.remove(marcador)

    # PASSO 4: listar as pastas e montar a fila. Nada é aberto nesta etapa.
    total_ok = 0
    total_erro = 0
    total_ja_exportado = 0
    total_pulado = 0
    primeiros_erros = []
    fila = []
    longos = []
    cancelado = False

    progresso = ui.createProgressDialog()
    progresso.isCancelButtonShown = True
    progresso.cancelButtonText = "Cancelar"
    progresso.show(TITULO, "Listando o projeto %v de %m", 0, len(pedido["projetos"]), 0)

    for indice_projeto, projeto in enumerate(pedido["projetos"]):
        progresso.progressValue = indice_projeto
        adsk.doEvents()
        pilha = [(projeto.rootFolder, [limpar_nome(pedido["hub"].name, MAX_PASTA), limpar_nome(projeto.name, MAX_PASTA)], "")]

        while len(pilha) > 0 and not progresso.wasCancelled:
            pasta, partes_locais, pasta_nuvem = pilha.pop()
            adsk.doEvents()
            try:
                subpastas = pasta.dataFolders.asArray()
                arquivos = pasta.dataFiles.asArray()
            except Exception as erro:
                total_erro = total_erro + 1
                primeiros_erros.append("Pasta " + pasta_nuvem + ": " + str(erro))
                registrar({"hub": pedido["hub"].name, "projeto": projeto.name, "pasta_nuvem": pasta_nuvem, "nome": "",
                           "id": "", "versao": "", "extensao": "", "base": ""},
                          "LISTAGEM", "erro", "", str(erro))
                escrever_log(traceback.format_exc())
                continue

            for subpasta in subpastas:
                pilha.append((subpasta, partes_locais + [limpar_nome(subpasta.name, MAX_PASTA)],
                              pasta_nuvem + "/" + subpasta.name))

            for arquivo_nuvem in arquivos:
                id_arquivo = arquivo_nuvem.id
                versao = str(arquivo_nuvem.versionNumber)
                extensao = arquivo_nuvem.fileExtension.lower()
                item = {"arquivo": arquivo_nuvem, "hub": pedido["hub"].name, "projeto": projeto.name,
                        "pasta_nuvem": pasta_nuvem, "nome": arquivo_nuvem.name, "id": id_arquivo,
                        "versao": versao, "extensao": extensao, "base": "", "formatos": []}

                # Desenhos, PDFs e outros tipos não passam pelo exportador de modelos.
                if extensao not in EXTENSOES_FUSION:
                    total_pulado = total_pulado + 1
                    if (id_arquivo, versao, "-") not in ultimo_status:
                        registrar(item, "-", "pulado", "", "Tipo ." + extensao + " não é modelo do Fusion")
                    continue

                # Mesma versão usa o mesmo nome local, para o retry cair ao lado dos outros formatos.
                if (id_arquivo, versao) in base_da_versao:
                    base = base_da_versao[(id_arquivo, versao)]
                else:
                    # Versão nova de um arquivo já exportado não apaga a anterior: ganha " vN".
                    inicial = limpar_nome(arquivo_nuvem.name, MAX_ARQUIVO)
                    if id_arquivo in ids_conhecidos:
                        inicial = inicial + " v" + versao
                    inicial = os.path.join(*(partes_locais + [inicial]))

                    # Nomes parecidos recebem (2), (3)... e nada existente é sobrescrito.
                    base = inicial
                    numero = 2
                    while True:
                        ocupado = base.casefold() in bases_usadas
                        for extensao_saida in EXTENSOES_SAIDA:
                            if os.path.exists(os.path.join(destino, base + extensao_saida)):
                                ocupado = True
                        if not ocupado:
                            break
                        base = inicial + " (" + str(numero) + ")"
                        numero = numero + 1
                    base_da_versao[(id_arquivo, versao)] = base
                    bases_usadas.add(base.casefold())
                item["base"] = base

                # Fica na fila só o formato que ainda não tem arquivo confirmado no disco.
                for formato in pedido["formatos"]:
                    anterior = ultimo_status.get((id_arquivo, versao, formato))
                    if (anterior is not None and anterior.get("status") == "ok"
                            and os.path.exists(os.path.join(destino, anterior.get("caminho", "")))):
                        total_ja_exportado = total_ja_exportado + 1
                    else:
                        item["formatos"].append(formato)

                if len(item["formatos"]) == 0:
                    continue

                if id_arquivo == id_a_pular and versao == versao_a_pular:
                    total_erro = total_erro + 1
                    motivo = "A execução anterior parou neste arquivo (o Fusion pode ter fechado). Pulado só desta vez."
                    primeiros_erros.append(item["nome"] + ": " + motivo)
                    registrar(item, "TODOS", "erro", "", motivo)
                    continue

                # ".step" é a extensão mais longa, então vale como medida para todos os formatos.
                if len(os.path.join(destino, base) + ".step") > LIMITE_CAMINHO:
                    longos.append(item)
                else:
                    fila.append(item)

        if progresso.wasCancelled:
            cancelado = True
            break

    progresso.hide()
    adsk.doEvents()

    if cancelado:
        if not gravar_linhas_pendentes():
            escrever_log("AVISO: linhas sem gravar no CSV: " + str(len(LINHAS_PENDENTES)))
        escrever_log("Cancelado durante a listagem. Nenhum arquivo foi aberto.")
        ui.messageBox("Cancelado durante a listagem. Nenhum arquivo foi aberto.", TITULO)
        return

    # PASSO 5: avisar antes de começar, inclusive sobre caminhos longos.
    texto = ("Arquivos a abrir: " + str(len(fila)) + "\n"
             + "Saídas já exportadas (serão puladas): " + str(total_ja_exportado) + "\n"
             + "Itens que não são modelo do Fusion (serão pulados): " + str(total_pulado) + "\n")
    if len(longos) > 0:
        texto = texto + ("\nATENÇÃO: " + str(len(longos)) + " arquivo(s) passariam do limite de "
                         + str(LIMITE_CAMINHO) + " caracteres e NÃO serão exportados.\n"
                         + "Exemplo: " + os.path.join(destino, longos[0]["base"]) + ".step\n"
                         + "Para incluí-los, responda Não e escolha uma pasta de destino mais curta.\n")
        for item in longos:
            escrever_log("Caminho longo: " + os.path.join(destino, item["base"]))
    texto = texto + "\nFeche os arquivos com alterações não salvas antes de continuar.\nContinuar?"

    icone = icone_aviso if len(longos) > 0 else adsk.core.MessageBoxIconTypes.QuestionIconType
    if ui.messageBox(texto, TITULO, botoes_sim_nao, icone) != adsk.core.DialogResults.DialogYes:
        if not gravar_linhas_pendentes():
            escrever_log("AVISO: linhas sem gravar no CSV: " + str(len(LINHAS_PENDENTES)))
        escrever_log("Resposta Não ao aviso. Nenhum arquivo foi aberto.")
        return

    for item in longos:
        motivo = "Caminho passaria de " + str(LIMITE_CAMINHO) + " caracteres. Use uma pasta de destino mais curta."
        for formato in item["formatos"]:
            total_erro = total_erro + 1
            registrar(item, formato, "erro", "", motivo)
        primeiros_erros.append(item["nome"] + ": " + motivo)

    # PASSO 6: exportar um arquivo por vez.
    if len(fila) > 0:
        progresso.show(TITULO, "Exportando %v de %m (%p%)", 0, len(fila), 0)

    for indice, item in enumerate(fila):
        if progresso.wasCancelled:
            cancelado = True
            break

        progresso.progressValue = indice
        progresso.message = "Exportando %v de %m (%p%): " + item["nome"].replace("%", "")
        adsk.doEvents()

        documento = None
        aberto_aqui = False
        try:
            with open(marcador, "w", encoding="utf-8") as arquivo:
                arquivo.write(item["id"] + "\t" + item["versao"] + "\t" + item["nome"])

            # 6.1: arquivo já aberto no Fusion é usado sem fechar, e só se estiver igual à nuvem.
            for indice_doc in range(app.documents.count):
                aberto = app.documents.item(indice_doc)
                try:
                    arquivo_do_aberto = aberto.dataFile
                except Exception:
                    arquivo_do_aberto = None
                if arquivo_do_aberto is not None and arquivo_do_aberto.id == item["id"]:
                    documento = aberto
                    break

            if documento is not None:
                if documento.isModified:
                    raise RuntimeError("Arquivo aberto com alterações não salvas. Salve ou feche e rode de novo.")
                if documento.dataFile.versionNumber != item["arquivo"].versionNumber:
                    raise RuntimeError("Arquivo aberto em outra versão. Feche e rode de novo.")
            else:
                documento = app.documents.open(item["arquivo"], ABRIR_VISIVEL)
                if documento is None:
                    raise RuntimeError("Documents.open não abriu o arquivo.")
                aberto_aqui = True
                adsk.doEvents()

            documento_fusion = adsk.fusion.FusionDocument.cast(documento)
            if documento_fusion is None:
                raise RuntimeError("O arquivo aberto não é um design do Fusion.")
            design = documento_fusion.design
            gerenciador = design.exportManager
            raiz = design.rootComponent
            tem_referencias = documento.allDocumentReferences.count > 0

            # 6.2: exportar cada formato para a pasta temporária e só depois mover.
            # Assim, erro ou queda nunca deixa arquivo pela metade no lugar definitivo.
            for formato in item["formatos"]:
                if progresso.wasCancelled:
                    cancelado = True
                    break
                try:
                    for sobra in os.listdir(pasta_temp):
                        os.remove(os.path.join(pasta_temp, sobra))

                    if formato == "F3D":
                        extensao_pedida = ".f3z" if tem_referencias else ".f3d"
                        opcoes = gerenciador.createFusionArchiveExportOptions(os.path.join(pasta_temp, "x" + extensao_pedida), raiz)
                    elif formato == "STEP":
                        opcoes = gerenciador.createSTEPExportOptions(os.path.join(pasta_temp, "x.step"), raiz)
                    else:
                        opcoes = gerenciador.createSTLExportOptions(raiz, os.path.join(pasta_temp, "x.stl"))

                    if opcoes is None:
                        raise RuntimeError("O Fusion não criou as opções de exportação.")
                    if not gerenciador.execute(opcoes):
                        raise RuntimeError("ExportManager.execute devolveu False.")
                    adsk.doEvents()

                    # A extensão final é a do arquivo que o Fusion realmente gerou.
                    gerados = os.listdir(pasta_temp)
                    if len(gerados) != 1:
                        raise RuntimeError("Esperava 1 arquivo gerado, veio: " + str(gerados))
                    caminho = item["base"] + os.path.splitext(gerados[0])[1].lower()
                    caminho_final = os.path.join(destino, caminho)
                    os.makedirs(os.path.dirname(caminho_final), exist_ok=True)
                    if os.path.exists(caminho_final):
                        raise RuntimeError("Já existe um arquivo neste caminho; nada foi sobrescrito.")
                    os.rename(os.path.join(pasta_temp, gerados[0]), caminho_final)

                    total_ok = total_ok + 1
                    registrar(item, formato, "ok", caminho, "")
                except Exception as erro:
                    total_erro = total_erro + 1
                    primeiros_erros.append(item["nome"] + " (" + formato + "): " + str(erro))
                    registrar(item, formato, "erro", "", str(erro))
                    escrever_log(traceback.format_exc())

        except Exception as erro:
            # Falha ao abrir: conta para todos os formatos pedidos deste arquivo.
            for formato in item["formatos"]:
                total_erro = total_erro + 1
                registrar(item, formato, "erro", "", str(erro))
            primeiros_erros.append(item["nome"] + ": " + str(erro))
            escrever_log(traceback.format_exc())

        # 6.3: fechar sem salvar só o que este script abriu.
        if aberto_aqui:
            try:
                documento.close(False)
                adsk.doEvents()
            except Exception as erro:
                escrever_log("AVISO: não consegui fechar " + item["nome"] + ": " + str(erro))
        try:
            os.remove(marcador)
        except OSError:
            pass

        if cancelado:
            break

    progresso.hide()

    # PASSO 7: limpar e garantir que todas as linhas foram ao CSV.
    try:
        for sobra in os.listdir(pasta_temp):
            os.remove(os.path.join(pasta_temp, sobra))
        os.rmdir(pasta_temp)
    except OSError:
        pass

    while not gravar_linhas_pendentes():
        resposta = ui.messageBox("O registro.csv está aberto em outro programa. Feche e clique em Repetir.",
                                 TITULO, adsk.core.MessageBoxButtonTypes.RetryCancelButtonType, icone_aviso)
        if resposta == adsk.core.DialogResults.DialogCancel:
            reserva = os.path.join(pasta_controle, "registro_pendente_" + datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S") + ".csv")
            with open(reserva, "w", encoding="utf-8-sig", newline="") as arquivo:
                escritor = csv.DictWriter(arquivo, fieldnames=COLUNAS, delimiter=";")
                escritor.writeheader()
                for linha in LINHAS_PENDENTES:
                    escritor.writerow(linha)
            LINHAS_PENDENTES.clear()
            escrever_log("AVISO: linhas salvas à parte em " + reserva)
            break

    # PASSO 8: resumo final.
    resumo = ("CANCELADO.\n\n" if cancelado else "Concluído.\n\n")
    resumo = (resumo + "Exportados agora: " + str(total_ok) + "\n"
              + "Já exportados antes: " + str(total_ja_exportado) + "\n"
              + "Com erro: " + str(total_erro) + "\n"
              + "Pulados (não são modelo do Fusion): " + str(total_pulado) + "\n")
    if len(primeiros_erros) > 0:
        resumo = resumo + "\nPrimeiros erros:\n" + "\n".join(texto_erro[:150] for texto_erro in primeiros_erros[:5]) + "\n"
    resumo = resumo + "\nRegistro: " + ARQUIVO_CSV + "\nLog: " + ARQUIVO_LOG
    escrever_log(resumo)
    ui.messageBox(resumo, TITULO)


class AoRodarExportacao(adsk.core.CustomEventHandler):
    # Roda quando o Fusion fica ocioso, já sem comando ativo. Aqui fechar documentos é permitido.
    def __init__(self):
        super().__init__()

    def notify(self, args):
        ui = adsk.core.Application.get().userInterface
        try:
            exportar_tudo(PEDIDO)
        except Exception:
            escrever_log(traceback.format_exc())
            ui.messageBox("Erro inesperado:\n\n" + traceback.format_exc(), TITULO)
        adsk.terminate()


class AoCriarComando(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        try:
            comando = args.command
            comando.okButtonText = "Continuar"
            comando.isExecutedWhenPreEmpted = False
            comando.setDialogInitialSize(420, 300)
            entradas = comando.commandInputs

            HUBS.clear()
            for hub in app.data.dataHubs.asArray():
                HUBS.append(hub)

            lista_hubs = entradas.addDropDownCommandInput("hub", "Hub", adsk.core.DropDownStyles.TextListDropDownStyle)
            hub_ativo = app.data.activeHub
            indice_ativo = 0
            for indice, hub in enumerate(HUBS):
                if hub_ativo is not None and hub.id == hub_ativo.id:
                    indice_ativo = indice
            for indice, hub in enumerate(HUBS):
                lista_hubs.listItems.add(hub.name, indice == indice_ativo, "", -1)

            entradas.addBoolValueInput("todos", "Todos os projetos do hub", True, "", False)
            lista_projetos = entradas.addDropDownCommandInput("projetos", "Projetos", adsk.core.DropDownStyles.CheckBoxDropDownStyle)
            lista_projetos.maxVisibleItems = 20

            PROJETOS.clear()
            if len(HUBS) > 0:
                for projeto in HUBS[indice_ativo].dataProjects.asArray():
                    PROJETOS.append(projeto)
                    lista_projetos.listItems.add(projeto.name, False, "", -1)

            entradas.addBoolValueInput("f3d", "F3D / F3Z (arquivo do Fusion)", True, "", True)
            entradas.addBoolValueInput("step", "STEP (para o SolidCraft)", True, "", True)
            entradas.addBoolValueInput("stl", "STL", True, "", False)

            ao_mudar = AoMudarEntrada()
            comando.inputChanged.add(ao_mudar)
            HANDLERS.append(ao_mudar)
            ao_executar = AoExecutarComando()
            comando.execute.add(ao_executar)
            HANDLERS.append(ao_executar)
            ao_destruir = AoDestruirComando()
            comando.destroy.add(ao_destruir)
            HANDLERS.append(ao_destruir)
        except Exception:
            app.userInterface.messageBox("Erro ao abrir a janela:\n\n" + traceback.format_exc(), TITULO)


class AoMudarEntrada(adsk.core.InputChangedEventHandler):
    # Trocar de hub recarrega a lista de projetos.
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            if args.input.id != "hub":
                return
            indice_hub = args.inputs.itemById("hub").selectedItem.index
            lista_projetos = args.inputs.itemById("projetos")
            lista_projetos.listItems.clear()
            PROJETOS.clear()
            for projeto in HUBS[indice_hub].dataProjects.asArray():
                PROJETOS.append(projeto)
                lista_projetos.listItems.add(projeto.name, False, "", -1)
        except Exception:
            adsk.core.Application.get().userInterface.messageBox("Erro ao trocar de hub:\n\n" + traceback.format_exc(), TITULO)


class AoExecutarComando(adsk.core.CommandEventHandler):
    # Só guarda as escolhas. A exportação em si roda depois, fora do comando.
    def __init__(self):
        super().__init__()

    def notify(self, args):
        global PEDIDO
        ui = adsk.core.Application.get().userInterface
        try:
            entradas = args.command.commandInputs
            lista_hubs = entradas.itemById("hub")
            lista_projetos = entradas.itemById("projetos")

            projetos = []
            for indice, projeto in enumerate(PROJETOS):
                if entradas.itemById("todos").value or lista_projetos.listItems.item(indice).isSelected:
                    projetos.append(projeto)

            formatos = []
            if entradas.itemById("f3d").value:
                formatos.append("F3D")
            if entradas.itemById("step").value:
                formatos.append("STEP")
            if entradas.itemById("stl").value:
                formatos.append("STL")

            if len(projetos) == 0 or len(formatos) == 0:
                ui.messageBox("Escolha pelo menos um projeto e um formato. Nada foi exportado.", TITULO)
                return

            PEDIDO = {"hub": HUBS[lista_hubs.selectedItem.index], "projetos": projetos, "formatos": formatos}
        except Exception:
            ui.messageBox("Erro ao ler as escolhas:\n\n" + traceback.format_exc(), TITULO)


class AoDestruirComando(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        if PEDIDO is None:
            adsk.terminate()
        else:
            adsk.core.Application.get().fireCustomEvent(ID_EVENTO, "")


def run(context):
    app = adsk.core.Application.get()
    ui = app.userInterface
    try:
        if app.isOffLine:
            ui.messageBox("O Fusion está offline. Conecte e rode de novo.", TITULO)
            return

        # Limpa restos de uma execução anterior que tenha sido interrompida.
        definicao = ui.commandDefinitions.itemById(ID_COMANDO)
        if definicao is not None:
            definicao.deleteMe()
        try:
            app.unregisterCustomEvent(ID_EVENTO)
        except Exception:
            pass

        evento = app.registerCustomEvent(ID_EVENTO)
        ao_rodar = AoRodarExportacao()
        evento.add(ao_rodar)
        HANDLERS.append(ao_rodar)

        definicao = ui.commandDefinitions.addButtonDefinition(ID_COMANDO, TITULO, "Exporta arquivos do Fusion para o computador.", "")
        ao_criar = AoCriarComando()
        definicao.commandCreated.add(ao_criar)
        HANDLERS.append(ao_criar)

        definicao.execute()
        adsk.autoTerminate(False)
    except Exception:
        ui.messageBox("Erro ao iniciar:\n\n" + traceback.format_exc(), TITULO)


def stop(context):
    app = adsk.core.Application.get()
    try:
        definicao = app.userInterface.commandDefinitions.itemById(ID_COMANDO)
        if definicao is not None:
            definicao.deleteMe()
        app.unregisterCustomEvent(ID_EVENTO)
    except Exception:
        pass
    HANDLERS.clear()
