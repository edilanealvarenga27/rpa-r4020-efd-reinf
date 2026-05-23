import pyautogui
import openpyxl
import pyperclip
import time
import re
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import logging
import traceback
import threading
import queue

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.15

logging.basicConfig(
    filename="log_robo_r4020.txt",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8"
)

inicio_execucao = None
inicio_linha = None
parar_robo = False
fila_log = queue.Queue()


# =========================
# LOGS
# =========================

def log(msg):
    print(msg)
    logging.info(msg)
    fila_log.put(msg)


def log_erro(msg):
    print("❌ " + msg)
    logging.error(msg)
    logging.error(traceback.format_exc())
    fila_log.put("❌ " + msg)


def atualizar_log_tela():
    while not fila_log.empty():
        msg = fila_log.get()
        caixa_log.insert("end", msg + "\n")
        caixa_log.see("end")

    janela.after(200, atualizar_log_tela)


def tirar_print_erro(etapa):
    nome = f"erro_{etapa}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    pyautogui.screenshot(nome)
    logging.error(f"Print de erro salvo: {nome}")
    log(f"📸 Print de erro salvo: {nome}")


def verificar_parada():
    if parar_robo:
        raise Exception("Execução interrompida pelo usuário.")


def etapa(nome):
    verificar_parada()
    log(f"➡️ ETAPA: {nome}")


def iniciar_cronometro_execucao():
    global inicio_execucao
    inicio_execucao = time.time()
    log("⏱️ Cronômetro geral iniciado.")


def finalizar_cronometro_execucao():
    fim = time.time()
    duracao = fim - inicio_execucao
    log(f"🎉 Execução total finalizada em {duracao:.2f} segundos.")


def iniciar_cronometro_linha(numero_linha):
    global inicio_linha
    inicio_linha = time.time()
    log(f"⏱️ Iniciando contagem da linha {numero_linha}.")


def finalizar_cronometro_linha(numero_linha):
    fim = time.time()
    duracao = fim - inicio_linha
    log(f"✅ Linha {numero_linha} finalizada em {duracao:.2f} segundos.")


# =========================
# FUNÇÕES AUXILIARES
# =========================

def remover_acentos(texto):
    texto = str(texto)
    texto = unicodedata.normalize("NFD", texto)
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    return texto.replace("º", "o").replace("ª", "a").replace("ç", "c").replace("Ç", "C")


def apenas_numeros(valor):
    return re.sub(r"\D", "", str(valor))


def formatar_valor_excel(valor):
    if isinstance(valor, (int, float)):
        return f"{valor:.2f}".replace(".", "")

    if isinstance(valor, str):
        valor = valor.strip().replace(".", "").replace(",", ".")
        try:
            f = float(valor)
            return f"{f:.2f}".replace(".", "")
        except:
            return "000"

    return "000"


def formatar_data(valor):
    if isinstance(valor, datetime):
        return valor.strftime("%d%m%Y")

    texto = str(valor).strip()
    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]

    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt).strftime("%d%m%Y")
        except:
            pass

    return apenas_numeros(texto)


def formatar_periodo(valor):
    texto = str(valor).strip().lower()

    meses = {
        "janeiro": "01",
        "fevereiro": "02",
        "marco": "03",
        "março": "03",
        "abril": "04",
        "maio": "05",
        "junho": "06",
        "julho": "07",
        "agosto": "08",
        "setembro": "09",
        "outubro": "10",
        "novembro": "11",
        "dezembro": "12",
    }

    texto_sem_acento = remover_acentos(texto).lower()

    if "/" in texto_sem_acento:
        partes = texto_sem_acento.split("/")
        mes = partes[0].strip()
        ano = partes[1].strip()

        if mes in meses:
            return f"{meses[mes]}/{ano}"

        if mes.isdigit():
            return f"{mes.zfill(2)}/{ano}"

    return str(valor).strip()


def ler_referencia(sheet, referencia, linha):
    referencia = referencia.strip().upper()

    if re.match(r"^[A-Z]+[0-9]+$", referencia):
        return sheet[referencia].value

    return sheet[f"{referencia}{linha}"].value


def colar_texto(texto):
    pyperclip.copy(str(texto))
    time.sleep(0.1)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)


# =========================
# LEITURA DA PLANILHA
# =========================

def ler_dados_planilha(config):
    etapa("Abrindo planilha")
    wb = openpyxl.load_workbook(config["planilha"], data_only=True)
    sheet = wb[config["aba"]]

    dados = []

    for linha in range(config["linha_inicial"], config["linha_final"] + 1):
        verificar_parada()
        etapa(f"Lendo linha {linha}")

        periodo = formatar_periodo(ler_referencia(sheet, config["periodo"], linha))
        cnpj_unidade = apenas_numeros(ler_referencia(sheet, config["cnpj_unidade"], linha))
        cnpj_beneficiario = apenas_numeros(ler_referencia(sheet, config["cnpj_beneficiario"], linha))
        natureza_rendimento = str(ler_referencia(sheet, config["natureza"], linha)).strip()
        observacao_darf = remover_acentos(ler_referencia(sheet, config["observacao"], linha))
        ob_pagamento = formatar_data(ler_referencia(sheet, config["ob_pagamento"], linha))
        valor_doc_origem = formatar_valor_excel(ler_referencia(sheet, config["valor_doc"], linha))
        base_calculo = formatar_valor_excel(ler_referencia(sheet, config["base_calculo"], linha))
        valor_receita = formatar_valor_excel(ler_referencia(sheet, config["valor_receita"], linha))

        registro = (
            periodo,
            cnpj_unidade,
            cnpj_beneficiario,
            natureza_rendimento,
            observacao_darf,
            ob_pagamento,
            valor_doc_origem,
            base_calculo,
            valor_receita,
        )

        dados.append(registro)

        log(
            f"✅ Linha {linha} carregada | "
            f"Período={periodo} | Unidade={cnpj_unidade} | "
            f"Beneficiário={cnpj_beneficiario} | Natureza={natureza_rendimento} | "
            f"OB={ob_pagamento} | Valor={valor_doc_origem} | "
            f"Base={base_calculo} | Receita={valor_receita}"
        )

    return dados


# =========================
# CAPTURA DO BOTÃO
# =========================

def capturar_posicao_botao():
    messagebox.showinfo(
        "Capturar botão R-4020",
        "Após clicar em OK, você terá 5 segundos para posicionar o mouse sobre o botão R-4020."
    )

    for i in range(5, 0, -1):
        log(f"⏳ Capturando posição em {i}...")
        time.sleep(1)

    pos = pyautogui.position()
    log(f"✅ Posição capturada: {pos}")
    return pos


# =========================
# ROBÔ COM MESMO FLUXO
# =========================

def preencher_e_navegar(
    pos_botao,
    periodo,
    cnpj_unidade,
    cnpj_beneficiario,
    natureza_rendimento,
    observacao_darf,
    ob_pagamento,
    valor_doc_origem,
    base_calculo,
    valor_receita,
):
    try:
        etapa("Clicando no botão R-4020")
        pyautogui.moveTo(pos_botao[0], pos_botao[1], duration=0.5)
        pyautogui.click()
        time.sleep(3)

        etapa("Entrando no campo Período de Apuração")
        pyautogui.press("tab")

        etapa("Preenchendo Período")
        pyautogui.write(periodo, interval=0.1)
        pyautogui.press("tab")

        etapa("Preenchendo CNPJ Unidade")
        pyautogui.write(cnpj_unidade, interval=0.1)
        pyautogui.press("tab")

        etapa("Preenchendo CNPJ Beneficiário")
        pyautogui.write(cnpj_beneficiario, interval=0.1)

        etapa("Avançando até botão Continuar")
        pyautogui.press("tab", presses=2, interval=0.1)
        pyautogui.press("enter")
        time.sleep(2)

        etapa("Segunda tela - confirmando")
        pyautogui.press("tab", presses=2, interval=0.1)
        pyautogui.press("enter")
        time.sleep(1)

        etapa("Avançando 6 TABs até campo Código")
        pyautogui.press("tab", presses=6, interval=0.1)
        time.sleep(2)

        etapa("Digitando código 17")
        pyautogui.write("17", interval=0.1)
        pyautogui.press("tab")
        time.sleep(2)

        etapa("Preenchendo Natureza do Rendimento")
        pyautogui.write(natureza_rendimento, interval=0.1)
        pyautogui.press("tab")

        etapa("Colando Observação Pré-Doc")
        colar_texto(observacao_darf)
        pyautogui.press("tab")

        etapa("Enviando dados preliminares")
        pyautogui.press("enter")
        time.sleep(1)

        etapa("Voltando até link complementar com 5 Shift+Tab")
        for i in range(5):
            verificar_parada()
            log(f"Shift+Tab {i + 1}/5")
            pyautogui.hotkey("shift", "tab")
            time.sleep(0.1)

        etapa("Entrando no link complementar")
        pyautogui.press("enter")
        time.sleep(2)

        etapa("Avançando 6 TABs até OB Pagamento")
        pyautogui.press("tab", presses=6, interval=0.1)

        etapa("Preenchendo Data OB Pagamento")
        pyautogui.write(ob_pagamento, interval=0.1)
        pyautogui.press("tab")

        etapa("Preenchendo Valor Documento de Origem")
        pyautogui.write(valor_doc_origem, interval=0.1)

        etapa("Avançando 4 TABs até Observação Final")
        pyautogui.press("tab", presses=4, interval=0.1)

        etapa("Colando Observação Final")
        colar_texto(observacao_darf)

        etapa("Avançando 3 TABs até Base de Cálculo")
        pyautogui.press("tab", presses=3, interval=0.1)

        etapa("Preenchendo Base de Cálculo")
        pyautogui.write(base_calculo, interval=0.1)
        pyautogui.press("tab")

        etapa("Preenchendo Valor Receita")
        pyautogui.write(valor_receita, interval=0.1)

        etapa("Avançando 7 TABs até botão Finalizar")
        pyautogui.press("tab", presses=7, interval=0.1)

        etapa("Finalizando envio")
        pyautogui.press("enter")
        time.sleep(1)

        etapa("Fechamentos finais - 3 Shift+Tab")
        for i in range(3):
            verificar_parada()
            log(f"Shift+Tab final {i + 1}/3")
            pyautogui.hotkey("shift", "tab")
            time.sleep(0.1)

        etapa("Enter fechamento 1")
        pyautogui.press("enter")

        etapa("2 TABs + Enter fechamento 2")
        pyautogui.press("tab", presses=2, interval=0.1)
        pyautogui.press("enter")

        log("✅ Linha enviada com sucesso!")

    except Exception:
        log_erro("Erro durante preenchimento/navegação")
        tirar_print_erro("preenchimento")
        raise


# =========================
# CONTROLES DA INTERFACE
# =========================

def selecionar_planilha():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[
            ("Arquivos Excel", "*.xlsx *.xlsm"),
            ("Todos os arquivos", "*.*")
        ]
    )

    if caminho:
        entrada_planilha.delete(0, tk.END)
        entrada_planilha.insert(0, caminho)


def parar_execucao():
    global parar_robo
    parar_robo = True
    log("🛑 Solicitação de parada enviada. O robô vai parar na próxima etapa segura.")


def animacao_trabalhando():
    if status_animando.get():
        atual = label_status.cget("text")
        if atual.endswith("..."):
            label_status.config(text="🤖 Robô trabalhando")
        else:
            label_status.config(text=atual + ".")
        janela.after(500, animacao_trabalhando)


def iniciar_robo_thread():
    thread = threading.Thread(target=iniciar_robo, daemon=True)
    thread.start()


def iniciar_robo():
    global parar_robo
    parar_robo = False

    try:
        config = {
            "planilha": entrada_planilha.get(),
            "aba": entrada_aba.get(),
            "linha_inicial": int(entrada_linha_inicial.get()),
            "linha_final": int(entrada_linha_final.get()),

            "periodo": entrada_periodo.get(),
            "cnpj_unidade": entrada_cnpj_unidade.get(),
            "cnpj_beneficiario": entrada_cnpj_beneficiario.get(),
            "natureza": entrada_natureza.get(),
            "observacao": entrada_observacao.get(),
            "ob_pagamento": entrada_ob_pagamento.get(),
            "valor_doc": entrada_valor_doc.get(),
            "base_calculo": entrada_base_calculo.get(),
            "valor_receita": entrada_valor_receita.get(),
        }

        if not config["planilha"]:
            messagebox.showerror("Erro", "Selecione a planilha.")
            return

        botao_iniciar.config(state="disabled")
        botao_parar.config(state="normal")

        status_animando.set(True)
        barra_progresso.start(10)
        animacao_trabalhando()

        iniciar_cronometro_execucao()

        dados = ler_dados_planilha(config)

        log(f"✅ Total de linhas carregadas: {len(dados)}")
        log("⚠️ Primeiro, vamos capturar a posição do botão R-4020.")

        pos_botao = capturar_posicao_botao()

        for i, linha in enumerate(dados, start=1):
            verificar_parada()

            log(f"\n🚀 ENVIANDO LINHA {i}/{len(dados)}")
            iniciar_cronometro_linha(i)

            preencher_e_navegar(pos_botao, *linha)

            finalizar_cronometro_linha(i)
            time.sleep(2)

        finalizar_cronometro_execucao()
        log("🎉 TODAS AS LINHAS PROCESSADAS COM SUCESSO!")

        messagebox.showinfo("Concluído", "Todas as linhas foram processadas com sucesso.")

    except Exception as e:
        log_erro(f"Execução interrompida: {e}")
        tirar_print_erro("execucao_geral")
        messagebox.showerror("Execução interrompida", str(e))

    finally:
        status_animando.set(False)
        barra_progresso.stop()
        label_status.config(text="✅ Robô parado/finalizado")

        botao_iniciar.config(state="normal")
        botao_parar.config(state="disabled")


# =========================
# JANELA
# =========================

janela = tk.Tk()
janela.title("Robô R-4020 - Configuração e Acompanhamento")
janela.geometry("840x830")

status_animando = tk.BooleanVar(value=False)

tk.Label(
    janela,
    text="Configuração do Robô R-4020",
    font=("Arial", 16, "bold")
).pack(pady=10)

frame = tk.Frame(janela)
frame.pack(padx=20, pady=10, fill="x")

tk.Label(frame, text="Planilha:").grid(row=0, column=0, sticky="w")
entrada_planilha = tk.Entry(frame, width=75)
entrada_planilha.grid(row=0, column=1, padx=5)
tk.Button(frame, text="Procurar", command=selecionar_planilha).grid(row=0, column=2)

tk.Label(frame, text="Aba:").grid(row=1, column=0, sticky="w")
entrada_aba = tk.Entry(frame, width=25)
entrada_aba.insert(0, "Relatório Base")
entrada_aba.grid(row=1, column=1, sticky="w", padx=5)

tk.Label(frame, text="Linha inicial:").grid(row=2, column=0, sticky="w")
entrada_linha_inicial = tk.Entry(frame, width=10)
entrada_linha_inicial.insert(0, "9")
entrada_linha_inicial.grid(row=2, column=1, sticky="w", padx=5)

tk.Label(frame, text="Linha final:").grid(row=3, column=0, sticky="w")
entrada_linha_final = tk.Entry(frame, width=10)
entrada_linha_final.insert(0, "10")
entrada_linha_final.grid(row=3, column=1, sticky="w", padx=5)

tk.Label(
    janela,
    text="Informe a coluna ou célula fixa de cada informação:",
    font=("Arial", 12, "bold")
).pack(pady=10)

frame_cols = tk.Frame(janela)
frame_cols.pack(padx=20)

campos = [
    ("Período de Apuração", "AG"),
    ("Estabelecimento / CNPJ Unidade", "C6"),
    ("Beneficiário / Recolhedor DARF", "L"),
    ("Natureza do Rendimento", "AP"),
    ("Observação Pré-Doc", "E"),
    ("Data OB Pagamento", "Z"),
    ("Valor Doc de Origem", "W"),
    ("Base de Cálculo DARF", "S"),
    ("Valor Receita DARF", "T"),
]

entradas = []

for i, (label, padrao) in enumerate(campos):
    tk.Label(frame_cols, text=label + ":").grid(row=i, column=0, sticky="w", pady=4)
    e = tk.Entry(frame_cols, width=12)
    e.insert(0, padrao)
    e.grid(row=i, column=1, sticky="w", padx=8)
    entradas.append(e)

(
    entrada_periodo,
    entrada_cnpj_unidade,
    entrada_cnpj_beneficiario,
    entrada_natureza,
    entrada_observacao,
    entrada_ob_pagamento,
    entrada_valor_doc,
    entrada_base_calculo,
    entrada_valor_receita,
) = entradas

frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=15)

botao_iniciar = tk.Button(
    frame_botoes,
    text="Iniciar Robô",
    command=iniciar_robo_thread,
    bg="#1f5fa8",
    fg="white",
    font=("Arial", 12, "bold"),
    padx=20,
    pady=8
)
botao_iniciar.grid(row=0, column=0, padx=10)

botao_parar = tk.Button(
    frame_botoes,
    text="Parar Robô",
    command=parar_execucao,
    bg="#b91c1c",
    fg="white",
    font=("Arial", 12, "bold"),
    padx=20,
    pady=8,
    state="disabled"
)
botao_parar.grid(row=0, column=1, padx=10)

label_status = tk.Label(
    janela,
    text="✅ Robô parado",
    font=("Arial", 11, "bold"),
    fg="#1f5fa8"
)
label_status.pack(pady=5)

barra_progresso = ttk.Progressbar(
    janela,
    orient="horizontal",
    mode="indeterminate",
    length=500
)
barra_progresso.pack(pady=5)

tk.Label(
    janela,
    text="Acompanhamento da execução:",
    font=("Arial", 11, "bold")
).pack(pady=(10, 2))

caixa_log = tk.Text(janela, height=14, width=100)
caixa_log.pack(padx=20, pady=5)

tk.Label(
    janela,
    text="Use preferencialmente o Google Chrome, pois a navegação por TAB pode mudar no Firefox.",
    fg="gray"
).pack(pady=5)

atualizar_log_tela()

janela.mainloop()