import shutil
import time
import os

def ensure_dir(file_path):
    '''Создаёт папку для хранения '''
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

Directory = os.getcwd()
Directory_BackUp = Directory + '_' + time.strftime(("%m.%d(%I)"), time.gmtime())
ensure_dir(Directory)

shutil.copytree(Directory, Directory_BackUp)