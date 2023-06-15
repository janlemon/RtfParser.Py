import subprocess

def install_dependencies():
    subprocess.check_call(['pip', 'install', '-r', 'requirements.txt'])

if __name__ == '__main__':
    install_dependencies()
