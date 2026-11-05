"""
Pergunta, ao desligar o PC, se deseja registrar o horário de SAÍDA na planilha.

Fica rodando escondido (o instalar.py faz ele iniciar junto com o Windows).
Quando você manda desligar/reiniciar e a Saída de hoje ainda está em branco:

  1. O Windows mostra a tela "Ponto automático está impedindo o desligamento".
     Clique em "Cancelar" nessa tela.
  2. Aparece a pergunta "Registrar a Saída de hoje às HH:MM?"
        Sim       -> grava a Saída e desliga
        Não       -> desliga sem gravar
        Cancelar  -> não desliga

Se a Saída de hoje já está preenchida, não pergunta nada e o PC desliga normal.
"""

import ctypes
import logging
import re
import subprocess
import sys
import threading
from ctypes import wintypes
from datetime import datetime, timedelta, timezone

import ponto

TITULO = "Ponto automático"
CLASSE_JANELA = "PontoAutomaticoSaida"

WM_CLOSE, WM_DESTROY = 0x0010, 0x0002
WM_QUERYENDSESSION, WM_ENDSESSION = 0x0011, 0x0016
WM_PERGUNTAR = 0x8000 + 1  # WM_APP + 1
ENDSESSION_LOGOFF = 0x80000000
WS_POPUP, WS_EX_TOOLWINDOW, WS_EX_NOACTIVATE = 0x80000000, 0x00000080, 0x08000000
SW_SHOWNA = 8
MB_YESNOCANCEL, MB_ICONQUESTION, MB_TOPMOST, MB_SETFOREGROUND = 0x3, 0x20, 0x40000, 0x10000
IDYES, IDNO = 6, 7
ERROR_ALREADY_EXISTS = 183

LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [("style", wintypes.UINT), ("lpfnWndProc", WNDPROC),
                ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int),
                ("hInstance", wintypes.HINSTANCE), ("hIcon", wintypes.HICON),
                ("hCursor", wintypes.HANDLE), ("hbrBackground", wintypes.HBRUSH),
                ("lpszMenuName", wintypes.LPCWSTR), ("lpszClassName", wintypes.LPCWSTR)]


user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.DefWindowProcW.restype = LRESULT
user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
user32.RegisterClassW.restype = wintypes.ATOM
user32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
                                   ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                   wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID]
user32.CreateWindowExW.restype = wintypes.HWND
user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]
user32.ShutdownBlockReasonCreate.argtypes = [wintypes.HWND, wintypes.LPCWSTR]
user32.ShutdownBlockReasonDestroy.argtypes = [wintypes.HWND]
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE
kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
kernel32.CreateMutexW.restype = wintypes.HANDLE

log = logging.getLogger("ponto")

ACOES = {"desligar": "desliga", "reiniciar": "reinicia", "sair": "sai da conta"}
COMANDOS = {
    "desligar": ["shutdown", "/s", "/t", "0"],
    "reiniciar": ["shutdown", "/r", "/t", "0"],
    "sair": ["shutdown", "/l"],
}


def tipo_de_desligamento(lparam: int) -> str:
    """'sair', 'reiniciar' ou 'desligar' (o Windows não diz direto se é reiniciar)."""
    if lparam & ENDSESSION_LOGOFF:
        return "sair"
    # Quando dá, o Windows registra o pedido no evento User32 1074 ("reiniciar" / "Desligado").
    desde = (datetime.now() - timedelta(minutes=2)).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    consulta = f"*[System[Provider[@Name='User32'] and EventID=1074 and TimeCreated[@SystemTime>='{desde}']]]"
    try:
        xml = subprocess.run(["wevtutil", "qe", "System", f"/q:{consulta}", "/f:xml", "/rd:true", "/c:1"],
                             capture_output=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW,
                             ).stdout.decode("utf-8", "replace")
    except (OSError, subprocess.SubprocessError):
        return "desligar"
    tipo = re.search(r"<Data Name=['\"]param5['\"]>([^<]*)<", xml)
    if tipo and re.search(r"reinici|restart", tipo.group(1), re.IGNORECASE):
        return "reiniciar"
    return "desligar"


def executar(tipo: str) -> None:
    """Roda o shutdown.exe. Ele só termina depois que o Windows pergunta a todos os apps se podem fechar."""
    codigo = subprocess.run(COMANDOS[tipo], creationflags=subprocess.CREATE_NO_WINDOW).returncode
    if codigo:
        log.warning("Comando para %s terminou com código %s (cancelado ou bloqueado por outro app?)", tipo, codigo)


class Vigia:
    """Janela invisível que recebe o aviso de desligamento do Windows."""

    def __init__(self, desligar=executar):
        self.desligar = desligar
        self.liberado = False      # True = deixa o próximo desligamento passar sem perguntar
        self.perguntando = False
        self.quando = None
        self.tipo = "desligar"
        self._wndproc = WNDPROC(self._tratar)  # guardar a referência (senão o Python apaga)
        instancia = kernel32.GetModuleHandleW(None)
        classe = WNDCLASSW(lpfnWndProc=self._wndproc, hInstance=instancia, lpszClassName=CLASSE_JANELA)
        user32.RegisterClassW(ctypes.byref(classe))
        # Janela "visível" mas fora da tela e fora da barra de tarefas: o Windows só mostra o
        # aviso "impedindo o desligamento" (e espera a resposta) para apps com janela visível.
        self.hwnd = user32.CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, CLASSE_JANELA, TITULO,
                                           WS_POPUP, -32000, -32000, 1, 1, None, None, instancia, None)
        if not self.hwnd:
            raise ctypes.WinError(ctypes.get_last_error())
        user32.ShowWindow(self.hwnd, SW_SHOWNA)

    def rodar(self) -> None:
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    def _tratar(self, hwnd, msg, wparam, lparam):
        try:
            if msg == WM_QUERYENDSESSION:
                return self._pedido_de_desligar(lparam)
            if msg == WM_ENDSESSION:
                if not wparam:          # desligamento foi cancelado
                    self.liberado = False
                return 0
            if msg == WM_PERGUNTAR:
                self._perguntar()
                return 0
            if msg == WM_DESTROY:
                user32.PostQuitMessage(0)
                return 0
        except Exception:
            log.exception("Erro no vigia de saída")
            if msg == WM_QUERYENDSESSION:
                return 1  # na dúvida, nunca impede o desligamento
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _pedido_de_desligar(self, lparam) -> int:
        """Retorna 1 para deixar desligar, 0 para segurar e perguntar."""
        if self.liberado:
            return 1
        if self.perguntando:
            return 0
        agora = datetime.now()
        if not ponto.saida_pendente(agora.date()):
            return 1
        self.quando, self.tipo, self.perguntando = agora, tipo_de_desligamento(lparam), True
        user32.ShutdownBlockReasonCreate(self.hwnd, "Clique em Cancelar para registrar a saída do ponto.")
        user32.PostMessageW(self.hwnd, WM_PERGUNTAR, 0, 0)  # pergunta depois de responder ao Windows
        log.info("Pedido para %s às %s: perguntando a saída", self.tipo, f"{agora:%H:%M:%S}")
        return 0

    def _perguntar(self) -> None:
        acao = ACOES[self.tipo]
        texto = (f"Registrar a Saída de hoje ({self.quando:%d/%m}) às "
                 f"{ponto.horario_saida(self.quando):%H:%M}?\n\n"
                 f"Sim: registra e {acao}\nNão: {acao} sem registrar\nCancelar: continua no Windows")
        resposta = user32.MessageBoxW(None, texto, TITULO,
                                      MB_YESNOCANCEL | MB_ICONQUESTION | MB_TOPMOST | MB_SETFOREGROUND)
        self.perguntando = False
        user32.ShutdownBlockReasonDestroy(self.hwnd)

        if resposta == IDYES:
            try:
                ponto.lancar_saida(self.quando)
            except PermissionError:
                log.exception("Sem permissão para salvar a planilha")
                ponto.avisar("Não consegui salvar a Saída: a planilha está aberta no Excel?\n\n"
                             "Feche a planilha e mande desligar de novo.")
                return
            except Exception:
                log.exception("Erro ao registrar a saída")
                ponto.avisar(f"Erro ao registrar a Saída. Detalhes em:\n{ponto.ARQUIVO_LOG}")
                return
        elif resposta != IDNO:
            log.info("Desligamento cancelado pelo usuário")
            return

        self.liberado = True
        log.info("Executando: %s", self.tipo)
        # Em outra thread: o shutdown.exe espera o Windows consultar todas as janelas, inclusive
        # esta. Se esperássemos aqui, esta janela não responderia e o desligamento falharia.
        threading.Thread(target=self.desligar, args=(self.tipo,), daemon=True).start()


def main() -> int:
    ponto.configurar_log()
    kernel32.CreateMutexW(None, False, f"Local\\{CLASSE_JANELA}")
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        return 0  # já está rodando
    kernel32.SetProcessShutdownParameters(0x3FF, 0)  # receber o aviso de desligamento antes dos outros apps
    vigia = Vigia()
    log.info("Vigia de saída iniciado")
    vigia.rodar()
    return 0


if __name__ == "__main__":
    sys.exit(main())
