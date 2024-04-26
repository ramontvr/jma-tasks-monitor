from flask import Flask, render_template
from flask_bootstrap import Bootstrap
import win32com.client
import pythoncom

from twilio.rest import Client
##CREDENCIAIS##

account_sid = 'Sua Account aqui'
auth_token = 'SeuTokenAqui'
client = Client(account_sid, auth_token)


def send_whatsapp_message(to, message):
    message = client.messages.create(
        from_='whatsapp:+ WHATSDOTWILIO', #Whatsapp do Twilio
        body=message,
        to='whatsapp:+5547999999999' #Numero que será enviado a msg
    )




app = Flask(__name__)
Bootstrap(app)

# Codigo de cada informação da Task
TASK_STATE = {0: "Desconhecido", 1: "Desabilitado", 2: "Queued", 3: "Pronto", 4: "Desconectado", 5: "Em execução"}
TASK_RESULT = {0: "Não iniciado", 1: "Erro", 2: "Desconhecido", 3: "Desativado", 4: "Satisfeito", 5: "Finalizado"}
TRIGGER_TYPE = {0: "Evento", 1: "Hora", 2: "Diário", 3: "Semanal", 4: "Mensal", 5: "MensalDOW", 6: "Idle", 7: "Registro", 8: "Logon", 9: "Boot", 10: "Terminado", 11: "Custom"}

def get_tasks():
    pythoncom.CoInitialize()
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()

    folders_queue = [scheduler.GetFolder('\\Microsoft\\Office')]
    tasks_info = []
    while folders_queue:
        folder = folders_queue.pop(0)
        folders_queue += list(folder.GetFolders(0))
        tasks = list(folder.GetTasks(0))
   
        for task in tasks:
            task_info = {
                "Nome da Tarefa": task.Name,
                "Status": TASK_STATE[task.State],
                "Próxima Execução": task.NextRunTime,
                "Última Execução": task.LastRunTime,
                "Resultado da Última Execução": TASK_RESULT[task.LastTaskResult],  # Use o dicionário de mapeamento aqui
                "Disparadores": [TRIGGER_TYPE[trigger.Type] for trigger in task.Definition.Triggers],  # Use o dicionário de mapeamento aqui,
               ##"Autor da Tarefa": task.Author,
               ##"Autor da Tarefa": task.RegistrationInfo.Author if hasattr(task.RegistrationInfo, 'Author') else 'Desconhecido',

               ## OBS: De alguma forma não está acessando o autor da task

              
            }
            tasks_info.append(task_info)

            ##MENSAGEM TWILIO##
            if  TASK_STATE[task.State] != 3:
            # Se o estado mudou, envie uma mensagem do WhatsApp
                send_whatsapp_message('+5596984155024', f'O estado da tarefa {task.Name} mudou para {TASK_STATE[task.State]}')
    return tasks_info

@app.route('/tasks')
def tasks():
    tasks = get_tasks()
    return render_template('tasks.html', tasks=tasks)

if __name__ == "__main__":
    app.run(debug=True)
