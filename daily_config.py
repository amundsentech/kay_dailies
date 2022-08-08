import importlib
import subprocess

import sys

pkgs = {'pandas': 'pd', 'tqdm': 'tqdm','numpy':'np','openpyxl':'openpyxl','openpyxl':'openpyxl'}
def check_packages():
    for p in pkgs:
        s = pkgs[p]
        try:
            print(f'check for {p}')
            s = importlib.import_module(p)
            
        except ImportError:
            print(f'{p} is not installed and has to be installed')
            subprocess.call([sys.executable, '-m', 'pip', 'install', p])
        finally:
            s = importlib.import_module(p)
            print(f'{p} is properly installed')
    return
try:
    check_packages()
except:
    print('upgrade pip then try again')
    subprocess.check_call([sys.executable,'-m','pip','install','--upgrade','pip'])
    check_packages()

path='/Volumes/GoogleDrive/Shared drives/AMC Projects/_AZ_Kay/_Drilling/_Shift Reports/'




