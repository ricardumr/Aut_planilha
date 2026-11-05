"""
Ponto automático: ao ligar o notebook, pergunta "Colocar ponto às HH:MM?"
com a hora em que ele foi ligado menos 3 minutos, e grava na coluna Entrada.

    Ligou às 07:45  ->  "Colocar ponto às 07:42?"  ->  Sim: grava 07:42

Como funciona
- Lê no log do Windows (Sistema) a primeira vez que o PC foi ligado hoje.
  Funciona mesmo com a "Inicialização Rápida" do Windows ativada.
- Planilha do mês: Documentos/Ponto/Horas Extra - <mês> <ano>.xlsm
  Se ela ainda não existe, cria a partir do modelo.xlsm com as datas do mês.
  As planilhas dos meses anteriores ficam guardadas na mesma pasta.
- Só pergunta se a Entrada de hoje está em branco. Respondeu "Não": não
  pergunta de novo no mesmo dia. Nunca sobrescreve um horário já preenchido.

A SAÍDA é perguntada na hora de desligar o PC (ver saida.py).

Uso
    python ponto.py             pergunta e lança a entrada de hoje
    python ponto.py --simular   só mostra o que seria perguntado, sem salvar
    python instalar.py          faz rodar sozinho (entrada ao ligar, saída ao desligar)
"""

import argparse
import calendar
import ctypes
import logging
import os
import re
import subprocess
import sys
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

# ============================ Configuração ============================
PASTA = Path(__file__).resolve().parent
PASTA_PLANILHAS = Path.home() / "Documents" / "Ponto"  # todas as planilhas ficam aqui
MODELO = PASTA / "modelo.xlsm"               # base para criar a planilha de cada mês
NOME_ARQUIVO = "Horas Extra - {mes} {ano}.xlsm"
MINUTOS_ANTES = 3   # entrada = hora que ligou o PC menos isso
MINUTOS_DEPOIS = 0  # saída = hora que mandou desligar mais isso (usado pelo saida.py)
IGNORAR_ANTES_DE = time(5, 0)  # PC ligado de madrugada (ex.: atualização do Windows) não conta
ARQUIVO_LOG = PASTA / "ponto.log"
ARQUIVO_RECUSA = PASTA / "entrada_recusada.txt"  # dia em que respondeu "Não" (não pergunta de novo)

# Layout da planilha (igual ao modelo)
PRIMEIRA_LINHA = 6  # linha do dia 1
COL_DATA, COL_DIA, COL_ENTRADA, COL_SAIDA, COL_HORARIO = 1, 2, 3, 4, 5
# ======================================================================

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
DIAS_SEMANA = ["Segunda-Feira", "Terça-Feira", "Quarta-Feira", "Quinta-Feira",
               "Sexta-Feira", "Sábado", "Domingo"]

# Eventos do log "Sistema" que indicam que o PC foi ligado.
EVENTOS_LIGAR = [
    ("Microsoft-Windows-Kernel-Boot", 27),          # boot (inclusive Inicialização Rápida)
    ("Microsoft-Windows-Kernel-General", 12),       # boot completo
    ("Microsoft-Windows-Power-Troubleshooter", 1),  # voltou da suspensão/hibernação
    ("Microsoft-Windows-Winlogon", 7001),           # logon do usuário
]

log = logging.getLogger("ponto")


def horarios_que_ligou(desde: datetime) -> dict[date, datetime]:
    """Primeira vez que o PC foi ligado em cada dia, a partir de `desde` (hora local)."""
    desde_utc = desde.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    filtros = " or ".join(f"(Provider[@Name='{p}'] and EventID={i})" for p, i in EVENTOS_LIGAR)
    consulta = f"*[System[({filtros}) and TimeCreated[@SystemTime>='{desde_utc}']]]"
    saida = subprocess.run(
        ["wevtutil", "qe", "System", f"/q:{consulta}", "/f:xml"],
        capture_output=True, timeout=60, check=True,
        creationflags=subprocess.CREATE_NO_WINDOW,
    ).stdout.decode("latin-1")

    primeiros: dict[date, datetime] = {}
    for carimbo in re.findall(r"SystemTime=['\"]([^'\"]+)['\"]", saida):
        quando = utc_para_local(carimbo)
        if quando.time() < IGNORAR_ANTES_DE:
            continue
        dia = quando.date()
        if dia not in primeiros or quando < primeiros[dia]:
            primeiros[dia] = quando
    return primeiros


def utc_para_local(carimbo: str) -> datetime:
    """'2026-11-05T10:47:42.1234567Z' -> datetime local (o Python só aceita 6 casas decimais)."""
    m = re.match(r"(\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d)(?:\.(\d+))?", carimbo)
    micro = int((m.group(2) or "0")[:6].ljust(6, "0"))
    utc = datetime.fromisoformat(m.group(1)).replace(microsecond=micro, tzinfo=timezone.utc)
    return utc.astimezone().replace(tzinfo=None)


def caminho_planilha(ano: int, mes: int) -> Path:
    """Ex.: Documentos/Ponto/Horas Extra - novembro 2026.xlsm"""
    return PASTA_PLANILHAS / NOME_ARQUIVO.format(mes=MESES[mes - 1], ano=ano)


def nova_planilha(ano: int, mes: int):
    """Planilha do mês a partir do modelo: datas, dias da semana e fórmulas do mês, horários em branco."""
    if not MODELO.exists():
        raise FileNotFoundError(f"Modelo não encontrado: {MODELO}")
    wb = load_workbook(MODELO, keep_vba=True)  # keep_vba mantém o formato .xlsm
    ws = wb.worksheets[0]
    ws.title = MESES[mes - 1].capitalize()

    dias_no_mes = calendar.monthrange(ano, mes)[1]
    for i in range(31):
        linha = PRIMEIRA_LINHA + i
        if i < dias_no_mes:
            dia = datetime(ano, mes, i + 1)
            ws.cell(linha, COL_DATA).value = dia
            ws.cell(linha, COL_DIA).value = DIAS_SEMANA[dia.weekday()]
            ws.cell(linha, COL_HORARIO).value = f"=SUM(D{linha}-C{linha})"
        else:
            for col in (COL_DATA, COL_DIA, COL_HORARIO):
                ws.cell(linha, col).value = None
        ws.cell(linha, COL_ENTRADA).value = None
        ws.cell(linha, COL_SAIDA).value = None
    return wb


def abrir_planilha(ano: int, mes: int):
    """(workbook, caminho, nova?) da planilha do mês; cria em memória a partir do modelo se não existir."""
    caminho = caminho_planilha(ano, mes)
    if caminho.exists():
        return load_workbook(caminho, keep_vba=True), caminho, False
    return nova_planilha(ano, mes), caminho, True


def linhas_por_data(ws) -> dict[date, int]:
    linhas = {}
    for linha in range(PRIMEIRA_LINHA, PRIMEIRA_LINHA + 31):
        valor = ws.cell(linha, COL_DATA).value
        if isinstance(valor, datetime):
            linhas[valor.date()] = linha
    return linhas


def salvar(wb, caminho: Path) -> None:
    """Salva num arquivo temporário e troca no fim, para nunca deixar a planilha corrompida."""
    caminho.parent.mkdir(parents=True, exist_ok=True)
    temp = caminho.with_suffix(".tmp")
    try:
        wb.save(temp)
        os.replace(temp, caminho)  # falha com PermissionError se a planilha estiver aberta no Excel
    finally:
        temp.unlink(missing_ok=True)


def garantir_planilha(ano: int, mes: int) -> None:
    """Cria a planilha do mês se ainda não existe (no 1º dia de uso do mês)."""
    wb, caminho, nova = abrir_planilha(ano, mes)
    if nova:
        salvar(wb, caminho)
        log.info("Planilha do mês criada: %s", caminho)


def em_branco(dia: date, coluna: int) -> bool:
    """True se a célula (Entrada/Saída) do dia está em branco na planilha."""
    wb, _, _ = abrir_planilha(dia.year, dia.month)
    ws = wb.worksheets[0]
    linha = linhas_por_data(ws).get(dia)
    return linha is not None and ws.cell(linha, coluna).value in (None, "")


def gravar(dia: date, coluna: int, horario: time) -> bool:
    """Grava o horário na coluna do dia, só se estiver em branco. True se gravou."""
    wb, caminho, _ = abrir_planilha(dia.year, dia.month)  # relê agora: a planilha pode ter mudado
    ws = wb.worksheets[0]
    linha = linhas_por_data(ws).get(dia)
    if linha is None:
        raise LookupError(f"Dia {dia:%d/%m/%Y} não encontrado na planilha {caminho.name}")
    celula = ws.cell(linha, coluna)
    if celula.value not in (None, ""):
        return False  # já preenchido: nunca sobrescreve
    celula.value = horario
    salvar(wb, caminho)
    return True


def entrada_recusada(dia: date) -> bool:
    try:
        return ARQUIVO_RECUSA.read_text().strip() == dia.isoformat()
    except FileNotFoundError:
        return False


def lancar_entrada(simular: bool = False, agora: datetime | None = None) -> None:
    """Pergunta "Colocar ponto às HH:MM?" (hora que ligou - 3 min) e grava a Entrada de hoje."""
    agora = agora or datetime.now()
    hoje = agora.date()
    if agora.time() < IGNORAR_ANTES_DE:
        return
    if not simular:
        garantir_planilha(hoje.year, hoje.month)
    if not em_branco(hoje, COL_ENTRADA):
        log.info("Entrada de hoje já preenchida")
        return
    if entrada_recusada(hoje):
        log.info("Entrada de hoje já foi recusada: não pergunta de novo")
        return

    # Primeira vez que ligou hoje. Sem boot hoje (PC ficou ligado desde ontem), vale o momento atual.
    ligou = horarios_que_ligou(datetime.combine(hoje, time())).get(hoje, agora)
    entrada = (ligou - timedelta(minutes=MINUTOS_ANTES)).replace(second=0, microsecond=0)
    pergunta = f"Colocar ponto às {entrada:%H:%M}?\n\n(notebook ligado às {ligou:%H:%M})"
    if simular:
        log.info("Simulação: perguntaria \"%s\"", pergunta.replace("\n\n", " "))
        return

    if not perguntar(pergunta):
        ARQUIVO_RECUSA.write_text(hoje.isoformat())
        log.info("%s: Entrada %s recusada", f"{hoje:%d/%m}", f"{entrada:%H:%M}")
    elif gravar(hoje, COL_ENTRADA, entrada.time()):
        log.info("%s: ligou às %s -> Entrada %s", f"{hoje:%d/%m}", f"{ligou:%H:%M:%S}", f"{entrada:%H:%M}")
    else:
        log.info("Entrada de hoje foi preenchida enquanto a pergunta estava aberta: mantida")


def horario_saida(quando: datetime) -> datetime:
    return (quando + timedelta(minutes=MINUTOS_DEPOIS)).replace(second=0, microsecond=0)


def saida_pendente(dia: date) -> bool:
    """True se a Saída do dia ainda está em branco na planilha."""
    return em_branco(dia, COL_SAIDA)


def lancar_saida(quando: datetime) -> None:
    """Grava a Saída do dia de `quando` (se ainda estiver em branco)."""
    saida = horario_saida(quando)
    if gravar(quando.date(), COL_SAIDA, saida.time()):
        log.info("%s: Saída %s registrada", f"{quando:%d/%m}", f"{saida:%H:%M}")
    else:
        log.info("%s: Saída já estava preenchida, mantida", f"{quando:%d/%m}")


TITULO = "Ponto automático"
MB_YESNO, MB_ICONQUESTION, MB_ICONWARNING = 0x4, 0x20, 0x30
MB_TOPMOST, MB_SETFOREGROUND, IDYES = 0x40000, 0x10000, 6


def perguntar(mensagem: str) -> bool:
    """Janela Sim/Não por cima de tudo. True se respondeu Sim."""
    estilo = MB_YESNO | MB_ICONQUESTION | MB_TOPMOST | MB_SETFOREGROUND
    return ctypes.windll.user32.MessageBoxW(None, mensagem, TITULO, estilo) == IDYES


def avisar(mensagem: str) -> None:
    ctypes.windll.user32.MessageBoxW(None, mensagem, TITULO, MB_ICONWARNING | MB_TOPMOST)


def configurar_log() -> None:
    handlers: list[logging.Handler] = [logging.FileHandler(ARQUIVO_LOG, encoding="utf-8")]
    if sys.stderr:  # no pythonw (sem console) não existe stderr
        handlers.append(logging.StreamHandler())
    logging.basicConfig(level=logging.INFO, handlers=handlers,
                        format="%(asctime)s %(levelname)s %(message)s", datefmt="%d/%m/%Y %H:%M:%S")


def main() -> int:
    parser = argparse.ArgumentParser(description="Pergunta e lança a hora de entrada (PC ligado - 3 min).")
    parser.add_argument("--simular", action="store_true", help="só mostra o que seria perguntado, sem salvar")
    args = parser.parse_args()
    configurar_log()
    try:
        lancar_entrada(args.simular)
    except PermissionError:
        log.exception("Sem permissão para salvar a planilha")
        avisar("Não consegui salvar a planilha de ponto. Ela está aberta no Excel?\n\n"
               "Feche a planilha: a pergunta aparece de novo no próximo desbloqueio do PC.")
        return 1
    except Exception:
        log.exception("Erro ao lançar o ponto")
        avisar(f"Erro ao lançar o ponto. Detalhes em:\n{ARQUIVO_LOG}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
