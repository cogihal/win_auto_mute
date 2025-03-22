import os
import sys
import datetime


def mute_log(msg: str):
    log_file = ''
    if hasattr(sys, '_MEIPASS'):
        log_file  = os.path.join(sys._MEIPASS, 'win_auto_mute.log') # type: ignore
    else:
        full_path_name = os.path.abspath(sys.argv[0])
        folder_name = os.path.dirname(full_path_name)
        log_file  = os.path.join(folder_name, 'win_auto_mute.log')

    # Time
    now = datetime.datetime.now()
    logtime = now.strftime('%Y/%m/%d %H:%M:%S')

    try:
        with open(log_file, 'a') as f:
            print(f'{logtime} : {msg}', file=f)
    except Exception as e:
        raise e

