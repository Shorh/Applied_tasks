def task_a():
    """
    Камни и украшения
    Даны две строки строчных латинских символов: строка J и строка S.
    Символы, входящие в строку J, — «драгоценности», входящие в строку S —
    «камни». Нужно определить, какое количество символов из S одновременно
    являются «драгоценностями». Проще говоря, нужно проверить, какое количество
    символов из S входит в J.
    """

    j = input()
    s = input()

    result = 0
    for ch in s:
        if ch in j:
            result += 1

    print(result)


def task_b():
    """
    Последовательно идущие единицы
    Требуется найти в бинарном векторе самую длинную последовательность единиц и
    вывести её длину.
    Желательно получить решение, работающее за линейное время и при этом
    проходящее по входному массиву только один раз.
    """

    with open('input.txt') as fin:
        n = int(fin.readline())

        count = 0
        count_i = 0

        for i in range(n):
            if int(fin.readline()) == 1:
                count_i += 1
            else:
                count_i = 0

            if count < count_i:
                count = count_i

    print(count)


def task_c():
    """
    Удаление дубликатов
    Дан упорядоченный по неубыванию массив целых 32-разрядных чисел. Требуется
    удалить из него все повторения.
    Желательно получить решение, которое не считывает входной файл целиком в
    память, т.е., использует лишь константный объем памяти в процессе работы.
    """

    with open('input.txt') as fin:
        n = int(fin.readline())

        with open('output.txt', 'w') as fout:
            prev = fin.readline()

            fout.write(prev)

            for i in range(1, n):
                now = fin.readline()
                if int(now) != int(prev):
                    fout.write(now)

                prev = now


def task_d():
    """
    Генерация скобочных последовательностей
    Дано целое число n. Требуется вывести все правильные скобочные
    последовательности длины 2 ⋅ n, упорядоченные лексикографически.
    В задаче используются только круглые скобки.
    Желательно получить решение, которое работает за время, пропорциональное
    общему количеству правильных скобочных последовательностей в ответе, и при
    этом использует объём памяти, пропорциональный n.
    """

    with open('input.txt') as fin:
        n = int(fin.readline())

    s = '(' * n + ')' * n

    def next_seq(string, num):
        c_open = 0
        c_close = 0

        for i in range(num - 1, -1, -1):
            if string[i] == '(':
                c_open += 1
                if c_close > c_open:
                    break
            else:
                c_close += 1

        ans = string[:num - c_open - c_close]

        if ans:
            ans = ans + ')' + '(' * c_open + ')' * (c_close - 1)

        return ans

    with open('output.txt', 'w') as fout:
        while s:
            fout.write(s + '\n')
            s = next_seq(s, 2 * n)


def task_e():
    """
    Анаграммы
    Даны две строки, состоящие из строчных латинских букв. Требуется определить,
    являются ли эти строки анаграммами, т. е. отличаются ли они только порядком
    следования символов.
    """

    with open('input.txt') as fin:
        a1, a2 = [line.rstrip('\n') for line in fin]

    res = 0

    if len(a1) == len(a2):
        dct1 = {}
        dct2 = {}

        for i in range(len(a1)):
            if a1[i] in dct1:
                dct1[a1[i]] += 1
            else:
                dct1[a1[i]] = 1

            if a2[i] in dct2:
                dct2[a2[i]] += 1
            else:
                dct2[a2[i]] = 1

        if dct1 == dct2:
            res = 1

    with open('output.txt', 'w') as fout:
        fout.write(str(res))


def task_f_1():
    """
    Слияние k сортированных списков
    Даны k отсортированных в порядке неубывания массивов неотрицательных целых
    чисел, каждое из которых не превосходит 100. Требуется построить результат
    их слияния: отсортированный в порядке неубывания массив, содержащий все
    элементы исходных k массивов.
    Длина каждого массива не превосходит 10 ⋅ k.
    Постарайтесь, чтобы решение работало за время k ⋅ log(k) ⋅ n, если считать,
    что входные массивы имеют длину n.
    """

    from collections import deque

    with open('input.txt') as fin:
        n = int(fin.readline())
        arr = []
        for i in range(n):
            x = fin.readline().rstrip('\n')
            x = x.split(' ')
            arr.append(deque(x[1:]))

    with open('output.txt', 'w') as fout:
        if n == 1:
            fout.write(' '.join(arr[0]))
        else:
            lst = {i for i in range(n)}
            while lst:
                minimum = 101
                index = -1
                for i in lst:
                    if int(arr[i][0]) < minimum:
                        minimum = int(arr[i][0])
                        index = i

                if len(arr[index]) == 1:
                    lst.discard(index)

                arr[index].popleft()

                if lst:
                    fout.write(str(minimum) + ' ')
                else:
                    fout.write(str(minimum))


def task_f_2():
    """
    Слияние k сортированных списков
    Даны k отсортированных в порядке неубывания массивов неотрицательных целых
    чисел, каждое из которых не превосходит 100. Требуется построить результат
    их слияния: отсортированный в порядке неубывания массив, содержащий все
    элементы исходных k массивов.
    Длина каждого массива не превосходит 10 ⋅ k.
    Постарайтесь, чтобы решение работало за время k ⋅ log(k) ⋅ n, если считать,
    что входные массивы имеют длину n.
    """

    with open('input.txt') as fin:
        n = int(fin.readline())
        arr = []
        for i in range(n):
            x = fin.readline().rstrip('\n')
            x = x.split(' ')
            for itm in range(1, int(x[0]) + 1):
                arr.append(int(x[itm]))

    arr.sort()

    with open('output.txt', 'w') as fout:
        fout.write(' '.join(str(itm) for itm in arr))


# task_a()
# task_b()
# task_c()
# task_d()
# task_e()
task_f_1()
# task_f_2()
