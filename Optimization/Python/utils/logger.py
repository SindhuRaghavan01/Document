from datetime import datetime
# from datetime import now, strptime, strftime

# from time import clock

class Logger():
    DEBUG, INFO, WARNING, ERROR = [2 ** i for i in range(4)]

    def __init__(self, log_file, log_level=DEBUG):
        self.log_level = log_level
        self.log_file = log_file
        self.fp = open(log_file, "w")

    @property
    def time(self):
        return datetime.strftime(datetime.now(), "%F %T.%f")[:-3]

    def write_log(self, message):
        if message:
            self.fp.write(message + "\n")
            self.fp.flush()

    def info(self, message: str):
        if self.log_level <= self.INFO:
            message = self.time + " INFO: " + message
            self.write_log(message)

    def debug(self, message: str):
        if self.log_level <= self.DEBUG:
            message = self.time + " DEBUG: " + message
            self.write_log(message)

    def warning(self, message: str):
        if self.log_level <= self.WARNING:
            message = self.time + " WARNING: " + message
            self.write_log(message)

    def error(self, message: str):
        if self.log_level <= self.ERROR:
            message = self.time + " ERROR: " + message
            self.write_log(message)

    def close(self):
        self.fp.close()
