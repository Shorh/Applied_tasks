from subprocess import call
from os.path import join
from os import environ, name, getcwd


def install(package):
    call(['pip', 'install', package], shell=True)


path = join('data', 'requirements.txt')

with open(path, encoding='utf8') as fin:
    for line in fin:
        install(line[:-1])

if name == 'nt':
    if 'PROGRAMFILES(X86)' in environ:
        file = join('data', 'msodbcsql_17.4.1.1_x64.msi')
    else:
        file = join('data', 'msodbcsql_17.4.1.1_x86.msi')
    call(['msiexec', '/i', file],
         shell=True)