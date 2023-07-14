import logging




class CustomHandler(logging.Handler):
    def __init__(self, log_writer):
        super().__init__()
        self.log_writer = log_writer

    def emit(self, record):
        log_message = self.format(record)
        self.log_writer(log_message)
        print(log_message)