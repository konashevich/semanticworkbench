import platform
import subprocess
import sys

if platform.system() == 'Windows':
    subprocess.run([
        'powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', 
        '-File', 'run_shared_venv.ps1', '--transport', 'stdio'
    ])
else:
    sys.exit(0)
