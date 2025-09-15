Instalação do serviço Windows usando nssm

1) Pré-requisitos
- nssm (Non-Sucking Service Manager) baixado e disponível. Ex.: extraia nssm.exe para C:\nssm\nssm.exe
- Ter Python instalado e `python` disponível no PATH, ou ajustar o caminho no passo de instalação.

2) Criar script runner
- Já existe `run_leads.bat` na pasta do projeto que inicia `leads.py`.

3) Instalar o serviço (PowerShell como Administrador)
- Abrir PowerShell como Administrador e executar:

  # criar serviço chamado 'CSLeads'
  C:\nssm\nssm.exe install CSLeads "C:\Windows\System32\cmd.exe" "/c C:\Users\alex\Documents\Python\Leads\run_leads.bat"

  # definir diretório de trabalho (opcional)
  C:\nssm\nssm.exe set CSLeads AppDirectory "C:\Users\alex\Documents\Python\Leads"

  # configurar reinício automático
  C:\nssm\nssm.exe set CSLeads Start SERVICE_AUTO_START
  C:\nssm\nssm.exe set CSLeads AppRestartDelay 5000

  # iniciar serviço
  Start-Service CSLeads

4) Logs
- Para ver logs do serviço via nssm (se configurado), olhe o diretório AppDirectory ou configure AppStdout/AppStderr no nssm:
  C:\nssm\nssm.exe set CSLeads AppStdout "C:\Users\alex\Documents\Python\Leads\logs\stdout.log"
  C:\nssm\nssm.exe set CSLeads AppStderr "C:\Users\alex\Documents\Python\Leads\logs\stderr.log"

5) Teste manual
- Antes de instalar como serviço, teste com:
  cd C:\Users\alex\Documents\Python\Leads
  .\run_leads.bat

Observações
- Ajuste caminhos (nssm.exe, python, run_leads.bat) conforme seu ambiente.
- Se usar um venv, use o python do venv no lugar de `python` no .bat ou aponte AppPath diretamente para o python.exe do venv e passe `leads.py` como argumento.
