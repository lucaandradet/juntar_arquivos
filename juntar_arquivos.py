"""
Junta todos os arquivos de uma pasta que tenham a mesma extensão em um único
arquivo, mantendo apenas um cabeçalho.

Como usar:
    1. Execute: python juntar_arquivos.py  (ou dê dois cliques no arquivo)
    2. Aguarde a tela de carregamento e selecione a pasta na janela que abrir.
    3. Se a pasta tiver arquivos de mais de uma extensão, escolha qual juntar.
    4. Acompanhe o andamento na janela de progresso.
    5. Ao terminar, responda se deseja juntar outra pasta. A nova seleção
       abre na última pasta escolhida. Responder "Não" encerra o programa.

O arquivo gerado fica dentro da própria pasta e recebe o nome dela.
Ex.: pasta "Vendas_2026" com arquivos .csv  ->  Vendas_2026/Vendas_2026.csv

Arquivos texto (csv, txt, tsv, ...) são concatenados linha a linha, sem
reinterpretar os dados. Arquivos Excel (xlsx, xls) são lidos com pandas
(necessário: pip install pandas openpyxl).
"""

import queue
import sys
import threading
import time
import tkinter as tk
from collections import Counter
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

EXTENSOES_EXCEL = {".xlsx", ".xlsm", ".xls"}
TEMPO_MINIMO_SPLASH = 1.5   # segundos que a tela inicial fica visível


# --------------------------------------------------------------------------
# Lógica de junção
# --------------------------------------------------------------------------

def fmt(numero: int) -> str:
    """Formata o número com ponto como separador de milhar: 23.461.002"""
    return f"{numero:,}".replace(",", ".")


def contar_extensoes(pasta: Path) -> Counter:
    """Conta quantos arquivos de cada extensão existem na pasta,
    ignorando o arquivo consolidado (nome da pasta) e temporários do Excel."""
    contagem = Counter()
    for p in pasta.iterdir():
        if (p.is_file() and p.suffix
                and p.stem != pasta.name
                and not p.name.startswith("~$")):
            contagem[p.suffix.lower()] += 1
    return contagem


def listar_arquivos(pasta: Path, extensao: str, saida: Path):
    return sorted(
        p for p in pasta.iterdir()
        if p.is_file()
        and p.suffix.lower() == extensao
        and not p.name.startswith("~$")
        and p.resolve() != saida.resolve()
    )


def decodificar(linha: bytes) -> str:
    """Decodifica uma linha em UTF-8 e, se falhar, em Latin-1."""
    try:
        return linha.decode("utf-8")
    except UnicodeDecodeError:
        return linha.decode("latin-1")


def juntar_texto(arquivos, saida: Path, log, progresso):
    """Concatena arquivos texto. O andamento é calculado pelos bytes lidos,
    então a porcentagem avança também dentro de arquivos grandes."""
    total_bytes = sum(a.stat().st_size for a in arquivos) or 1
    lidos = 0
    cabecalho_ref = None
    total = 0
    qtd = len(arquivos)

    with open(saida, "w", encoding="utf-8-sig", newline="") as out:
        for i, arq in enumerate(arquivos, start=1):
            progresso(i, qtd, arq.name, lidos / total_bytes)
            with open(arq, "rb") as f:
                bruto = f.readline()
                lidos += len(bruto)
                if not bruto.strip():
                    log.append(f"[vazio] {arq.name}")
                    continue
                cab = decodificar(bruto.removeprefix(b"\xef\xbb\xbf")).rstrip("\r\n")
                if cabecalho_ref is None:
                    cabecalho_ref = cab
                    out.write(cab + "\n")
                elif cab != cabecalho_ref:
                    log.append(f"[AVISO] cabeçalho diferente em {arq.name}")

                linhas = 0
                for bruto in f:
                    lidos += len(bruto)
                    linha = decodificar(bruto).rstrip("\r\n")
                    if linha.strip():
                        out.write(linha + "\n")
                        linhas += 1
                    if linhas % 5000 == 0:
                        progresso(i, qtd, arq.name, lidos / total_bytes)
                total += linhas
                log.append(f"[ok] {arq.name}: {fmt(linhas)} linhas")

    progresso(qtd, qtd, "", 1.0)
    return total


def juntar_excel(arquivos, saida: Path, log, progresso):
    qtd = len(arquivos)
    progresso(0, qtd, "Carregando bibliotecas do Excel...", 0.0)
    import pandas as pd

    frames, colunas_ref = [], None
    for i, arq in enumerate(arquivos, start=1):
        # leitura = 90% do trabalho, gravação final = 10%
        progresso(i, qtd, arq.name, (i - 1) / qtd * 0.9)
        df = pd.read_excel(arq, dtype=str)
        if colunas_ref is None:
            colunas_ref = list(df.columns)
        elif list(df.columns) != colunas_ref:
            log.append(f"[AVISO] colunas diferentes em {arq.name}")
        frames.append(df)
        log.append(f"[ok] {arq.name}: {fmt(len(df))} linhas")

    progresso(qtd, qtd, "Gravando arquivo final...", 0.9)
    consolidado = pd.concat(frames, ignore_index=True)
    consolidado.to_excel(saida, index=False)
    progresso(qtd, qtd, "", 1.0)
    return len(consolidado)


def juntar(pasta: Path, extensao: str, progresso=None):
    """Junta os arquivos e devolve (caminho_saida, qtd_arquivos, total_linhas, log).

    progresso(arquivo_atual, total_arquivos, nome, fracao) é chamado durante a
    execução, com fracao entre 0 e 1.
    """
    progresso = progresso or (lambda *a: None)
    ext_saida = ".xlsx" if extensao in EXTENSOES_EXCEL else extensao
    saida = pasta / f"{pasta.name}{ext_saida}"
    arquivos = listar_arquivos(pasta, extensao, saida)
    log = []
    if extensao in EXTENSOES_EXCEL:
        total = juntar_excel(arquivos, saida, log, progresso)
    else:
        total = juntar_texto(arquivos, saida, log, progresso)
    return saida, len(arquivos), total, log


# --------------------------------------------------------------------------
# Interface
# --------------------------------------------------------------------------

def centralizar(janela, largura=None, altura=None):
    """Posiciona a janela no centro da tela."""
    janela.update_idletasks()
    largura = largura or janela.winfo_reqwidth()
    altura = altura or janela.winfo_reqheight()
    x = (janela.winfo_screenwidth() - largura) // 2
    y = (janela.winfo_screenheight() - altura) // 2
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


def criar_janela_base():
    """Cria a janela principal invisível no centro da tela.

    As caixas de diálogo (seleção de pasta e mensagens) usam essa janela
    como "mãe" e por isso abrem centralizadas sobre ela, ou seja, no meio
    da tela.
    """
    root = tk.Tk()
    root.title("Juntar arquivos")
    centralizar(root, 1, 1)
    root.attributes("-alpha", 0.0)      # invisível
    root.attributes("-topmost", True)   # diálogos aparecem na frente
    root.update()
    return root


def mostrar_splash(root):
    """Tela inicial de carregamento, sem bordas, no centro da tela."""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.attributes("-topmost", True)

    moldura = tk.Frame(splash, bg="#1f4e79", padx=2, pady=2)
    moldura.pack(fill="both", expand=True)
    conteudo = tk.Frame(moldura, bg="white", padx=40, pady=25)
    conteudo.pack(fill="both", expand=True)

    tk.Label(conteudo, text="Juntar Arquivos", bg="white", fg="#1f4e79",
             font=("Segoe UI", 16, "bold")).pack()
    tk.Label(conteudo, text="Carregando, aguarde...", bg="white", fg="#555555",
             font=("Segoe UI", 10)).pack(pady=(5, 15))
    barra = ttk.Progressbar(conteudo, mode="indeterminate", length=260)
    barra.pack()
    barra.start(12)

    centralizar(splash)
    inicio = time.time()
    while time.time() - inicio < TEMPO_MINIMO_SPLASH:
        splash.update()
        time.sleep(0.02)
    barra.stop()
    splash.destroy()


TITULO_SELECAO = "Selecione a pasta com os arquivos a juntar"


def centralizar_dialogo_windows(titulo: str, timeout: float = 5.0):
    """Centraliza uma janela nativa do Windows pelo título.

    A janela de seleção de pasta é criada pelo próprio Windows, que a abre na
    posição que memorizou e ignora a janela-base. Esta função roda em segundo
    plano, espera a janela aparecer e a move para o centro da tela.
    """
    if sys.platform != "win32":
        return

    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.FindWindowW.restype = wintypes.HWND
    SWP_NOSIZE, SWP_NOZORDER = 0x0001, 0x0004

    def mover():
        fim = time.time() + timeout
        while time.time() < fim:
            hwnd = user32.FindWindowW("#32770", titulo)
            if hwnd and user32.IsWindowVisible(hwnd):
                ret = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(ret))
                larg, alt = ret.right - ret.left, ret.bottom - ret.top
                x = (user32.GetSystemMetrics(0) - larg) // 2   # largura da tela
                y = (user32.GetSystemMetrics(1) - alt) // 2    # altura da tela
                user32.SetWindowPos(hwnd, None, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER)
                return
            time.sleep(0.01)

    threading.Thread(target=mover, daemon=True).start()


def selecionar_pasta(root, initialdir=None) -> str:
    """Abre a seleção de pasta centralizada na tela."""
    centralizar_dialogo_windows(TITULO_SELECAO)
    return filedialog.askdirectory(parent=root, title=TITULO_SELECAO,
                                   initialdir=initialdir)


def escolher_extensao(root, contagem: Counter):
    """Janela para escolher a extensão quando a pasta tem mais de uma."""
    janela = tk.Toplevel(root)
    janela.title("Escolha a extensão")
    janela.resizable(False, False)
    janela.attributes("-topmost", True)

    ttk.Label(janela, text="A pasta tem arquivos de mais de uma extensão.\n"
                           "Qual deseja juntar?").pack(padx=20, pady=(15, 10))

    opcoes = [f"{ext}  ({qtd} arquivos)" for ext, qtd in contagem.most_common()]
    escolha = tk.StringVar(value=opcoes[0])
    ttk.Combobox(janela, textvariable=escolha, values=opcoes,
                 state="readonly", width=30).pack(padx=20)

    resultado = {"ext": None}

    def confirmar():
        resultado["ext"] = escolha.get().split()[0]
        janela.destroy()

    botoes = ttk.Frame(janela)
    botoes.pack(pady=15)
    ttk.Button(botoes, text="Juntar", command=confirmar).pack(side="left", padx=5)
    ttk.Button(botoes, text="Cancelar", command=janela.destroy).pack(side="left", padx=5)

    centralizar(janela)
    janela.grab_set()
    root.wait_window(janela)
    return resultado["ext"]


def fmt_tempo(segundos: float) -> str:
    """Formata uma duração: 45 s / 2 min 05 s / 1 h 03 min."""
    s = int(round(segundos))
    if s < 60:
        return f"{s} s"
    if s < 3600:
        return f"{s // 60} min {s % 60:02d} s"
    return f"{s // 3600} h {(s % 3600) // 60:02d} min"


def juntar_com_progresso(root, pasta: Path, extensao: str, qtd_arquivos: int):
    """Executa a junção em segundo plano mostrando uma janela de progresso.

    Devolve (resultado de juntar(), duração em segundos) ou levanta a
    exceção que ocorreu.
    """
    janela = tk.Toplevel(root)
    janela.title("Juntando arquivos")
    janela.resizable(False, False)
    janela.attributes("-topmost", True)
    janela.protocol("WM_DELETE_WINDOW", lambda: None)   # não deixa fechar no meio

    quadro = ttk.Frame(janela, padding=20)
    quadro.pack(fill="both", expand=True)

    ttk.Label(quadro, text=f"Pasta: {pasta.name}",
              font=("Segoe UI", 10, "bold")).pack(anchor="w")
    ttk.Label(quadro, text=f"{fmt(qtd_arquivos)} arquivo(s) {extensao} para juntar"
              ).pack(anchor="w", pady=(2, 12))

    txt_arquivo = tk.StringVar(value="Preparando...")
    ttk.Label(quadro, textvariable=txt_arquivo, width=50).pack(anchor="w")

    barra = ttk.Progressbar(quadro, mode="determinate", maximum=100, length=360)
    barra.pack(pady=(8, 4))
    txt_pct = tk.StringVar(value="0%")
    ttk.Label(quadro, textvariable=txt_pct,
              font=("Segoe UI", 11, "bold")).pack()

    txt_decorrido = tk.StringVar(value="Tempo decorrido: 0 s")
    txt_restante = tk.StringVar(value="Tempo restante: calculando...")
    ttk.Label(quadro, textvariable=txt_decorrido).pack(anchor="w", pady=(12, 0))
    ttk.Label(quadro, textvariable=txt_restante).pack(anchor="w")

    centralizar(janela)
    janela.grab_set()

    fila = queue.Queue()
    ultimo = {"pct": -1, "arq": None}

    def progresso(atual, total, nome, fracao):
        # Chamado pela thread de trabalho: só envia quando algo mudou,
        # para não sobrecarregar a interface.
        pct = int(fracao * 1000)    # resolução de 0,1% para a estimativa
        if pct != ultimo["pct"] or nome != ultimo["arq"]:
            ultimo["pct"], ultimo["arq"] = pct, nome
            fila.put(("progresso", atual, total, nome, fracao))

    def trabalhar():
        try:
            fila.put(("fim", juntar(pasta, extensao, progresso)))
        except Exception as e:      # repassa o erro para a interface
            fila.put(("erro", e))

    resultado = {}
    estado = {"fracao": 0.0, "ultima_estimativa": 0.0}
    inicio = time.time()

    def atualizar_tempo():
        """Atualiza o tempo decorrido e a estimativa de término.

        A estimativa usa a velocidade média desde o início:
        restante = decorrido x (parte que falta / parte já feita).
        """
        agora = time.time()
        decorrido = agora - inicio
        txt_decorrido.set(f"Tempo decorrido: {fmt_tempo(decorrido)}")

        # recalcula a cada meio segundo para o número não ficar "pulando"
        if agora - estado["ultima_estimativa"] < 0.5:
            return
        estado["ultima_estimativa"] = agora

        f = estado["fracao"]
        if f < 0.01 or decorrido < 2:
            txt_restante.set("Tempo restante: calculando...")
            return
        restante = decorrido * (1 - f) / f
        termino = time.strftime("%H:%M:%S", time.localtime(agora + restante))
        txt_restante.set(f"Tempo restante: ~{fmt_tempo(restante)}  "
                         f"(término previsto às {termino})")

    def atualizar_tela():
        try:
            while True:
                msg = fila.get_nowait()
                if msg[0] == "progresso":
                    _, atual, total, nome, fracao = msg
                    if nome and atual:
                        txt_arquivo.set(f"Arquivo {fmt(atual)} de {fmt(total)}: {nome}")
                    elif nome:
                        txt_arquivo.set(nome)
                    estado["fracao"] = fracao
                    pct = int(fracao * 100)
                    barra["value"] = pct
                    txt_pct.set(f"{pct}%")
                else:
                    resultado[msg[0]] = msg[1]
                    janela.destroy()
                    return
        except queue.Empty:
            pass
        atualizar_tempo()
        janela.after(50, atualizar_tela)

    threading.Thread(target=trabalhar, daemon=True).start()
    atualizar_tela()
    root.wait_window(janela)
    duracao = time.time() - inicio

    if "erro" in resultado:
        raise resultado["erro"]
    return resultado["fim"], duracao


def processar_pasta(root, pasta: Path):
    """Junta os arquivos de uma pasta e mostra o resultado ao usuário."""
    contagem = contar_extensoes(pasta)
    if not contagem:
        messagebox.showwarning("Nada a juntar",
                               f"Nenhum arquivo encontrado em:\n{pasta}", parent=root)
        return

    if len(contagem) == 1:
        extensao = next(iter(contagem))
    else:
        extensao = escolher_extensao(root, contagem)
        if not extensao:
            return

    try:
        (saida, qtd, total, log), duracao = juntar_com_progresso(
            root, pasta, extensao, contagem[extensao])
    except ImportError:
        messagebox.showerror("Dependência ausente",
                             "Para juntar arquivos Excel instale:\n\n"
                             "pip install pandas openpyxl", parent=root)
        return
    except Exception as e:
        messagebox.showerror("Erro ao juntar", str(e), parent=root)
        return

    avisos = [l for l in log if l.startswith("[AVISO]") or l.startswith("[vazio]")]
    msg = (f"{fmt(qtd)} arquivo(s) {extensao} juntado(s).\n"
           f"{fmt(total)} linhas de dados + 1 cabeçalho.\n"
           f"Tempo total: {fmt_tempo(duracao)}\n\n"
           f"Arquivo gerado:\n{saida}")
    if avisos:
        msg += "\n\nAtenção:\n" + "\n".join(avisos)
        messagebox.showwarning("Concluído com avisos", msg, parent=root)
    else:
        messagebox.showinfo("Concluído", msg, parent=root)


def main():
    root = criar_janela_base()
    mostrar_splash(root)

    ultima_pasta = None
    while True:
        pasta_sel = selecionar_pasta(root, initialdir=ultima_pasta)
        # Cancelou a seleção (na primeira vez ou depois de responder "Sim"
        # para juntar mais uma pasta): encerra o programa.
        if not pasta_sel:
            break
        pasta = Path(pasta_sel)
        ultima_pasta = str(pasta)

        processar_pasta(root, pasta)

        if not messagebox.askyesno("Continuar?",
                                   "Deseja juntar os arquivos de mais alguma pasta?",
                                   parent=root):
            break

    root.destroy()


if __name__ == "__main__":
    main()
