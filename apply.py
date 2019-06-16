import sys

import sys
import re
import subprocess
from sys import argv

if __name__ == '__main__':

    if argv[1] == 'en':
        bashCommand = 'python3  setup.py en'
        #bashCommand1 = 'python3  send.py en'
        bashCommand1 = 'python3  draft.py en'

    elif argv[1] == 'gis':
        bashCommand = 'python3  setup.py gis'
        bashCommand1 = 'python3  draft.py gis'
        #bashCommand1 = 'python3  send.py en'

    elif argv[1] == 'data':
        bashCommand = 'python3  setup.py data'
        bashCommand1 = 'python3  draft.py data'
        #bashCommand1 = 'python3  send.py en'

    elif argv[1] == 'w':
        bashCommand = 'python3  setup.py w'
        bashCommand1 = 'python3  draft.py w'
        #bashCommand1 = 'python3  send.py w'

    elif argv[1] == 'c':
        bashCommand = 'python3  setup.py c'
        bashCommand1 = 'python3  draft.py c'
        #bashCommand1 = 'python3  send.py c'

    else:
        print ('You need to give an argument!')

    process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()

    process = subprocess.Popen(bashCommand1.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()
