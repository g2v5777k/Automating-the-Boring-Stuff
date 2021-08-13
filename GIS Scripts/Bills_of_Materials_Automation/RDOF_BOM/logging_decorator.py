import logging
import os
import functools
import datetime

# reference https://medium.com/swlh/add-log-decorators-to-your-python-project-84094f832181 , https://dev.to/aldo/implementing-logging-in-python-via-decorators-1gje

root = os.path.dirname(os.path.abspath(__file__))
logDir = os.path.join(root, 'test_logs') # whatever dir you want to store logs in
today = datetime.datetime.now()
def makelogger(logName):
    '''
    Creates a log file, and returns a log object
    :return: log object
    '''
    logName = logName + '_{}'.format(today.strftime('%m.%d.%Y'))
    logPath = os.path.join(logDir, str(logName) + '.log') if os.path.exists(os.path.join(logDir, str(logName) + '.log')) else os.path.join(logDir, str(logName) + '.log')
    logger = logging.getLogger(logName)
    logger.setLevel(logging.INFO)
    # create file handler, log format and add formatting to file handler
    fileHandler = logging.FileHandler(logPath)
    log_format = '%(levelname)s %(asctime)s %(message)s'
    formatter = logging.Formatter(log_format)
    if not logger.handlers:
        logger.addHandler(fileHandler)
    return logger

def logError(logName = 'Log'):
    def errorLog(func=None):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # runs function, writes to log for function and logs args + errors, if any.
            try:
                logfunc = makelogger(logName)
                # logfunc.info('Run Date: {}'.format(today.strftime('%m/%d/%Y %H:%M:%S')))
                logfunc.info('Running function' + ' ' + func.__name__ + f' with arguments {args}')
                logfunc.info(func(*args, **kwargs))
                return func(*args, **kwargs)
            except Exception as e:
                logfunc = makelogger(logName)
                message = 'Exception has occured at --> ' + func.__name__+'\n'
                logfunc.exception(message)
                return e
        return wrapper
    return errorLog


