

import threading


threads = []

def run_task(task, on_complete = None, on_error = None):
    
    def run():
        try:
            task()
            if on_complete:
                on_complete()
        except Exception as e:
            if on_error:
                on_error(e)
                
    thread = threading.Thread(target=run)
    thread.daemon = True
    thread.start()