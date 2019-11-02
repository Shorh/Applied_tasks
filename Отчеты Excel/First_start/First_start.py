from subprocess import call
from os.path import join
from os import environ, name


def install(package):
    call(['pip', 'install', package], shell=True)


path = join('requirements.txt')

with open(path, encoding='utf8') as fin:
    for line in fin:
        install(line[:-1])

if name == 'nt':
    if 'PROGRAMFILES(X86)' in environ:
        file = join('msodbcsql_17.4.1.1_x64.msi')
    else:
        file = join('msodbcsql_17.4.1.1_x86.msi')
    call(['msiexec', '/i', file],
         shell=True)
