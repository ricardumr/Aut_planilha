"""
Faz o ponto rodar sozinho:
  - ENTRADA (ponto.py): tarefa no Agendador de Tarefas, roda ao fazer logon
    (quando liga o PC) e ao desbloquear a tela.
  - SAÍDA (saida.py): inicia junto com o Windows e fica esperando você mandar
    desligar, para perguntar se registra a saída.

    python instalar.py            instala (ou atualiza)
    python instalar.py --remover  remove tudo
"""

import argparse
import ctypes
import os
import subprocess
import sys
import tempfile
import time
import winreg
from pathlib import Path
from xml.sax.saxutils import escape

NOME_TAREFA = "Ponto Automatico"
NOME_INICIALIZACAO = "PontoAutomaticoSaida"  # valor na chave Run do usuário
CHAVE_RUN = r"Software\Microsoft\Windows\CurrentVersion\Run"
CLASSE_JANELA_SAIDA = "PontoAutomaticoSaida"  # mesma do saida.py
PASTA = Path(__file__).resolve().parent
SCRIPT_ENTRADA = PASTA / "ponto.py"
SCRIPT_SAIDA = PASTA / "saida.py"
PYTHONW = Path(sys.executable).with_name("pythonw.exe")  # pythonw = roda sem abrir janela


def xml_tarefa() -> str:
    usuario = escape(f"{os.environ['USERDOMAIN']}\\{os.environ['USERNAME']}")
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Pergunta "Colocar ponto às HH:MM?" (hora que ligou o PC menos 3 min) e grava a entrada.</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{usuario}</UserId>
    </LogonTrigger>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <StateChange>SessionUnlock</StateChange>
      <UserId>{usuario}</UserId>
    </SessionStateChangeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{usuario}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT12H</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{escape(str(PYTHONW))}</Command>
      <Arguments>"{escape(str(SCRIPT_ENTRADA))}"</Arguments>
      <WorkingDirectory>{escape(str(PASTA))}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def parar_vigia_saida() -> None:
    """Fecha o saida.py se estiver rodando (espera até 5 s)."""
    user32 = ctypes.windll.user32
    for _ in range(50):
        hwnd = user32.FindWindowW(CLASSE_JANELA_SAIDA, None)
        if not hwnd:
            return
        user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
        time.sleep(0.1)


def instalar() -> None:
    if not PYTHONW.exists():
        sys.exit(f"pythonw.exe não encontrado ao lado de {sys.executable}")

    # Entrada: tarefa agendada
    with tempfile.TemporaryDirectory() as pasta_temp:
        arquivo = Path(pasta_temp) / "tarefa.xml"
        arquivo.write_text(xml_tarefa(), encoding="utf-16")
        subprocess.run(["schtasks", "/create", "/tn", NOME_TAREFA, "/xml", str(arquivo), "/f"], check=True)

    # Saída: iniciar junto com o Windows, e já iniciar agora (reinicia se já estava rodando)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CHAVE_RUN, 0, winreg.KEY_SET_VALUE) as chave:
        winreg.SetValueEx(chave, NOME_INICIALIZACAO, 0, winreg.REG_SZ, f'"{PYTHONW}" "{SCRIPT_SAIDA}"')
    parar_vigia_saida()
    subprocess.Popen([str(PYTHONW), str(SCRIPT_SAIDA)], cwd=PASTA,
                     creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)

    print(f'Entrada: tarefa "{NOME_TAREFA}" roda ao ligar o PC e ao desbloquear a tela.')
    print("Saída: pergunta ao desligar já está ativa e inicia junto com o Windows.")
    print(f'Para testar a entrada agora: schtasks /run /tn "{NOME_TAREFA}"  (resultado em ponto.log)')


def remover() -> None:
    subprocess.run(["schtasks", "/delete", "/tn", NOME_TAREFA, "/f"])
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CHAVE_RUN, 0, winreg.KEY_SET_VALUE) as chave:
            winreg.DeleteValue(chave, NOME_INICIALIZACAO)
    except FileNotFoundError:
        pass
    parar_vigia_saida()
    print("Ponto automático removido.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Faz o ponto rodar sozinho ao ligar e ao desligar o PC.")
    parser.add_argument("--remover", action="store_true", help="remove a tarefa e a inicialização automática")
    remover() if parser.parse_args().remover else instalar()
