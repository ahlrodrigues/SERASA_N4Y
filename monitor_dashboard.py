
import time
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess

WATCHED_FOLDER = './download'
SCRIPT_TO_RUN = 'gerar_dashboard.py'

class ChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.src_path.endswith('.xlsx'):
            print(f'[INFO] Detecção de modificação: {event.src_path}')
            executar_script()

    def on_created(self, event):
        if event.src_path.endswith('.xlsx'):
            print(f'[INFO] Novo arquivo detectado: {event.src_path}')
            executar_script()

def executar_script():
    print('[INFO] Executando script de geração do dashboard...')
    result = subprocess.run(['python', SCRIPT_TO_RUN])
    if result.returncode == 0:
        print('[SUCESSO] Script executado com sucesso.')
    else:
        print('[ERRO] Houve uma falha na execução do script.')

if __name__ == '__main__':
    print(f'[MONITOR] Observando a pasta: {WATCHED_FOLDER}')
    event_handler = ChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, WATCHED_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
