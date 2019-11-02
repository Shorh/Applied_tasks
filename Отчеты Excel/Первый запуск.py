from subprocess import call
from os.path import join


def install(package):
    call(['pip', 'install', package], shell=True)


path = join('data', 'requirements.txt')

with open(path, encoding='utf8') as fin:
    for line in fin:
        install(line[:-1])
