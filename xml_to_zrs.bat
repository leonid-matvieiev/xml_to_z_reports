0<1# :: ^
""" Со след строки bat-код до последнего :: и тройных кавычек
@echo off
setlocal enabledelayedexpansion
py -3 -x "%~f0" %*
IF !ERRORLEVEL! NEQ 0 (
    echo ERRORLEVEL !ERRORLEVEL!
    pause
) else (
    timeout /t 20
)
exit /b !ERRORLEVEL! :: Со след строки py-код """

from sys import argv, exit, path
ps_ = r'PyScripter\Lib\rpyc.zip' in ' '.join(path)

from os import system, environ
from os.path import join, split, splitext, exists, getctime, abspath
from pprint import pprint
from glob import glob
import xmltodict
import sys, py_compile  # , decimal
from decimal import Decimal as Dec

# import pprint
import json, win32api
import os, re, time

restart_sign = '-rs'

# 0 - без відлагодження
# 1 - без відлагодження, локал cache
# 2 - без перезапуска не батника
# 3 - .
# 4 - .
debug_prot = 0

if debug_prot > 0:
    fpne_cache = 'cache'
else:  # environ.get("TEMP", '')
    fpne_cache = join(r'C:\Windows\Temp', 'c21kdh4mtf8rxv8o3b8n '
        '9xzmfr764lty7s0blnv6 k79tk09t9o68mu9xi944'.strip().split()[1])
if debug_prot > 0:
    # Отладка
    flagsf = ''  # 'RSH'  #
    flagsn = ''  # '> nul'  #
    print(fpne_cache)
else:
    # штатная работа
    flagsf = 'RSH'  # ''  #
    flagsn = '> nul'  # ''  #

prots = None
# prots = [-1, '2021.01.11']

hed = '<?xml version="1.0" encoding="WINDOWS-1251"?><root>%s</root>'
# ----------------------------------------------------------------------------

# ============================================================================
def save_prots():
    # Запис файлу
    txt = ' '.join(str(cel) for cel in prots)
    system(f'attrib -R -S -H "{fpne_cache}" {flagsn}')
    with open(fpne_cache, 'w') as fp:
        fp.write(txt)
    flagsf and system(f'attrib %s "{fpne_cache}" {flagsn}' %
                        ' '.join('+' + c for c in flagsf))
# ----------------------------------------------------------------------------

# ============================================================================
def get_prots(dn=0):
    global prots  # Уже считано, или установлено
    if win32api.GetKeyState(0x10) & 0x80:  # SHIFT
        prots_ini()
        # 0 или 1 - клавиша отжата
        # (-127) или (-128) - клавиша нажата
    if not (prots and isinstance(prots, list) and len(prots) == 3):
        prots_is = True
        while True:
            if exists(fpne_cache):  # Перевірка наявності файлу
                with open(fpne_cache) as fp:
                    txt = fp.read()
                flagsf and system(f'attrib %s "{fpne_cache}" {flagsn}' %
                                    ' '.join('+' + c for c in flagsf))
            else:
                prots_is = False
                break
            try:  # перетворення в числа
                prots = [int(s) for s in txt.rsplit(None, 3)[-3:]]
                if len(prots) != 3:  # Перевірка кількості значень
                    prots_is = False
                    break
            except ValueError:  # не числа
                prots_is = False
            break
        if not prots_is:  # ініціалізація
            prots_ini()
    rez = prots[0] <= 0 and 22 or prots[1] <= 0 and 23
    prots[0] -= abs(dn)
    ddd = int(prots[2] - time.time())
    if prots[1] > ddd:
        prots[1] = ddd
    return rez
    # возврат 22 если закончился prots[0], 23 - если prots[1]
# ----------------------------------------------------------------------------

# ============================================================================
def prots_ini():
    global prots
    prots = [600, 60 * 3600*24]  # Уст нач колич и ост дней
    prots.append(int(time.time() + prots[-1]))  # Опр и доб кон даты
    save_prots()
# ----------------------------------------------------------------------------

# ============================================================================
def str_insert_chars(dt, pos = (4,6,8,10,12), chars='.. ::'):
    dt = list(dt)
    for p, c in reversed(list(zip(pos, chars))):
        dt.insert(p, c)
    return ''.join(dt)
# ----------------------------------------------------------------------------

# ============================================================================
def rows_from_dir(fpn1):
    rows = []
#    print(end='.', flush=True)
    nm1 = split(fpn1)[-1][2:]
    # Выделение необходимых папок \d\d
    fpn2s = sorted(filter(lambda fpn2: os.path.isdir(fpn1),
                map(lambda nm2: join(fpn1, nm2), filter(lambda nm2:
                    re.match(r'\d\d$', nm2), os.listdir(fpn1)))))
    for fpn2 in fpn2s:
        nm2 = split(fpn2)[-1]
        if int(nm2) % 10 == 9:
            print(end=f' {nm2}', flush=True)
        # Выделение необходимых файлов
        fpnes = sorted(filter(lambda fpne: os.path.isfile(fpne),
                map(lambda nm3: join(fpn2, nm3), filter(lambda nm3:
                re.match(rf'^{nm1}{nm2}-\d\d.xml$', nm3),
                            os.listdir(fpn2)))))
        if fpnes:  # зет отчёт в последнем файле, если он есть
            with open(fpnes[-1]) as fx:  # извлечение из всех файлов
                # содержимого тегов 'DAT' в список словарей
                dats = xmltodict.parse(hed % fx.read())['root']['DAT']
            if isinstance(dats, dict):
                dats = [dats]
            zdats = filter(lambda dat: 'Z' in dat, dats)
            for dat in zdats:  # тип DAT
                # разделители в дату/время TS, hd0[1]
                dct_z = dat['Z']
                dct_z_lst_t = dct_z['TXS']
                dct_z_lst_m = dct_z.get('M', [])
                if isinstance(dct_z_lst_m, dict):
                    dct_z_lst_m = [dct_z_lst_m]
                cels = [dat["TS"][:6],
                            ', '.join((dat['@ZN'], dat['@FN'], dat['@TN']))]
                cels.append(dct_z['@NO'])
                cels.append(dat["TS"])  #[:8]
                cels.append(sum(int(m.get('@SMI', '0')) for m in dct_z_lst_m))
                cels += [int(t.get('@SMI', '0')) for t in dct_z_lst_t]
                cels.append(sum(int(m.get('@SMO', '0')) for m in dct_z_lst_m))
                cels += [int(t.get('@SMO', '0')) for t in dct_z_lst_t]
                rate = tuple(map(lambda t: t['@TXPR'], dct_z_lst_t[:-2])) + \
                    tuple(map(lambda t: t['@DTPR'], dct_z_lst_t[:-1])) + \
                    tuple(map(lambda t: t['@DTNM'][:6], dct_z_lst_t[:-1]))
                cels.append(rate)
                rows.append(cels)  # всю получившуюся строку
    if int(nm2) % 10 != 9:
        print(end=f' {nm2}', flush=True)
    return rows
# ----------------------------------------------------------------------------

# ============================================================================
def blocks_from_rows(rows):
    flag_last_mon = True
    flag_first_step = True
    old_mon = rows[-1][0]
    for i, cels in reversed(list(enumerate(rows))):
        if flag_first_step:
            flag_first_step = False
            continue
        if rows[i + 1][0] != cels[0]:  # Смена месяца
            flag_last_mon = False
            continue
        if flag_last_mon:
            continue
        if rows[i + 1][1] != cels[1]:  # Смена номеров
            continue
        if rows[i + 1][-1] != cels[-1]:  # Смена ставок
            continue
        rows[i + 1][4:-1] = map(sum, zip(rows[i + 1][4:-1], cels[4:-1]))
        del rows[i]
    print('Місячних Z-звітів -', len(rows))
    if 0:
        with open('json_rows_2.txt', 'w', encoding='cp1251') as fx:
            json.dump(rows, fx, indent=2, ensure_ascii=False)
    blocks = {}
    for cels in rows:
        blocks.setdefault(cels[1], {}  #
             ).setdefault(tuple(cels[-1][:-5]), (cels[-1], []))[-1].append(cels[2:-1])
    if 0:
        with open('json_blocks.txt', 'w', encoding='cp1251') as fx:
            json.dump(blocks, fx, indent=2, ensure_ascii=False)
    return blocks
# ----------------------------------------------------------------------------

hd1 = 'N Дата Оборот оА оБ оВ оГ оД оЕ Повер пА пБ пВ пГ пД пЕ'.split()
dpref = ('', '&nbsp;', '&thinsp;')[0]
dsuf = ('', '&nbsp;', ' &#8372;')[2]  # ' &#x20b4;'
dlim = (',', '.', "'")[0]
csind = 0

# ============================================================================
def xls_from_block(block):

    title, rates = block
#    print(os.curdir)
#    print(abspath(os.curdir))
#    print(abspath(fpne))
    fpne = title.replace(',', '') + '.xls'
    print(fpne)
    fpne = join(split(argv[0])[0], fpne)

    ss = [f'''\
<!DOCTYPE html>
<HTML>
  <HEAD>
    <META charset="{("utf-8", "windows-1251")[csind]}">
    <TITLE>{title}</TITLE>
  </HEAD>
  <BODY>
    <TABLE cellspacing="0" cellpadding="2" border="1" align="center">
    <caption><i><h2><font color="blue">Апарат {title}</font></h2></i></caption>\
'''  # <font size="5">
    ]  #  align= "right" "center" "left"
    for rate, reps in rates.values():
        ss.append('<TR><TH colspan="3"><i>Ставки</i>' +
                    ''.join(map(lambda t: '<TH><font color="red">%s</font>' %
                    ('%s&nbsp;%%' % t if t != '0.00' else ''), rate[:4])) +
                    '<TH colspan="3"><i>Збори</i>' +
                    ''.join(map(lambda t: '<TH><font color="red">%s</font>' %
                    ('%s<BR>%s&nbsp;%%' % t if t[1] != '0.00' else ''),
                    zip(rate[9:], rate[4:9]))) + '<TH>')

        ss.append('<TH>'.join(['<TR>'] + hd1))

        TXPR, DTPR = 0, 0
        for cels in reps:
            color = ' bgcolor="yellow"' if Dec(cels[2]) < Dec(cels[9]) else ''
            ss.append(f'<TR align="right"><TD>{cels[0]}<TD>'
                f'{str_insert_chars(cels[1])}' +  # , (4,6), ".."
                (f'<TD{color}>{dpref}{Dec(cels[2]) / 100:.2f}{dsuf}' + ''.join(
                map(lambda t: f'<TD>{dpref}{Dec(t) / 100:.2f}{dsuf}', cels[3:])
                )).replace('.', dlim))
            TXPR += cels[2]
            DTPR += cels[9]

        ss.append('<TR><TH colspan="3" align="right"><font color="red">' +
                    f'{dpref}{Dec(TXPR) / 100:.2f}{dsuf}'.replace('.', ',') +
                    '</font><TH colspan="5"><i>Строка контроля</i>'
                    '<TH colspan="2" align="right"><font color="red">'
                    + f'{dpref}{Dec(DTPR) / 100:.2f}{dsuf}'.replace('.', ',') +
                    '</font><TH colspan="6">')
    ss.append('''\
    </TABLE>
  </BODY>
</HTML>''')
    for i in range(30):
        try:
            with open(fpne, 'w', encoding=("utf8", "cp1251")[csind]) as fx:
                fx.write('\n'.join(ss))
            if debug_prot < 1:
                os.startfile(f'"{fpne}"')
            print('Створено, записано і відкрито файл')
            break
        except PermissionError:
            if i == 0:
                print('? ? ? Звільніть/закрийте зайнятий файл')
            if i < 3:
                print(end='\a', flush=True)
            print(end='.', flush=True)
            time.sleep(1)
    else:
        print('Не вдалося записати в зайнятий файл')
# ----------------------------------------------------------------------------

# from multiprocessing import Pool
# from multiprocessing.dummy import Pool
# FILE_THREADS_NUMBER = 4

# ============================================================================
def rows_from_xmls(main_dir):
    # Выделение необходимых папок \d{10}
    fpn1s = sorted(filter(lambda fpn1: os.path.isdir(fpn1),
                map(lambda nm1: join(main_dir, nm1), filter(lambda nm1:
                    re.match(r'\d{10}$', nm1), os.listdir(main_dir)))))

    tbeg = time.time()
    if 1:
        rows = []
        for fpn1 in fpn1s:
            print(end=split(fpn1)[-1], flush=True)
            rows += rows_from_dir(fpn1)
            print(f' : {len(rows)}')
#        print()
    elif 1:
        rows = sum(map(rows_from_dir, fpn1s), [])
        print()
    else:
        pool = Pool(FILE_THREADS_NUMBER)
        rezs = pool.map(rows_from_dir, fpn1s)
    #    rezs = pool.map_async(build, devs)
        pool.close()
        pool.join()  # Ожидание заверш build_runs, для работы в фоне
        rows = sum(rezs, [])  # rezs.get() если map_async
        print()
    print('Затрачено  хв:сек -', '%i:%02i' % divmod(time.time() - tbeg, 60))
    return rows
# ----------------------------------------------------------------------------

# ============================================================================
def xlss_from_blocks(blocks):
#    print(get_prots(), prots)
    prot = get_prots(-1)
    if prot:
        print(f'Непередбачена помилка {prot}')
        return
    save_prots()
    for block in blocks.items():
        xls_from_block(block)
    print("Затримок", '-'.join(f'{str(prot)[:1]}{len(str(prot)) or ""}'
                                                for prot in prots[1::-1]))
# ----------------------------------------------------------------------------

# ============================================================================
def main():
    """ режими роботи програми
    0 - Зчитування файлів з папок і запис файлу щоденних Z-звітів
    1 - Зчитування файлу щоденних Z-звітів, груп помісячно і створення *.XLS
    2 - Зчитування файлів з папок, групування помісячно і створення *.XLS
     """

    print(sys.version)
    rejim = 2

    # Якщо закидуємо на прогу *.json, то його і обробляємо
    sd_files = list(filter(lambda a: os.path.isfile(a) and
                                splitext(a)[-1].lower() == '.json', argv[1:]))
    fpne_json = 'rows.json' # імя для запису в файл
    if sd_files:
        fpne_json = sd_files[0]
        rejim = 1

    if rejim in (0, 2):
        sd_dirs = list(filter(os.path.isdir, argv[1:]))
        if sd_dirs:
            # если есть перетаскивание на программу
            if len(sd_dirs) > 1:
                print('? Передано для обробки папок більше однієї:')
                print(' ', '\n  '.join(set(map(lambda p: split(p)[-1],
                                                sd_dirs))))
                return
            else:
                # выбираем единственную перетасканную папку
                sd_dir = sd_dirs[0]
        else:
            sd_dirs = sorted(filter(os.path.isdir, glob('*_microSD')),
                                                        key=getctime)
            if sd_dirs:  # Вибираємо найсвіжішу *_microSD
                sd_dir = sd_dirs[-1]
                if len(sd_dirs) > 1:
                    print(f'! З {len(sd_dirs)}-х папок '
                                    'для обробки обрано створену найпізніше')
            else:
                print('? Не передано/знайдено для обробки жодної папки')
                return
        print('Обробка папки:', split(sd_dir)[-1])
        # считываем все файлы и формируем список строк
        rows = rows_from_xmls(sd_dir)
    if 1 or rejim == 0:  # режим 0 і не потрібен
        # сохраняем список строк в json-файл
        try:
            with open(fpne_json, 'w', encoding='cp1251') as fx:
                json.dump(rows, fx, indent=2, ensure_ascii=False)
        except:
            print('? Помилка запису в *.json')
            return
    if rejim == 1:
        # считываем список строк с json-файла
        with open(fpne_json, encoding='cp1251') as fx:
            rows = json.load(fx)
    print('Щоденних Z-звітів -', len(rows))
    if rejim in (1, 2):
        if 0: xls_from_rows(rows)
        blocks = blocks_from_rows(rows)
        if 0: xls_from_rows(rows)
        if 1: xlss_from_blocks(blocks)
# ----------------------------------------------------------------------------


# ============================================================================
def restart(restart_sign, delay=20):
    global prots  # ???
    if debug_prot < 2 and ps_:
        return
    if restart_sign not in argv[1:]:
        if debug_prot < 2:
            system((f'@timeout /t {delay}', '@pause')[
                system(" ".join(['@'] + argv + [restart_sign])) and 1])
        if get_prots() and debug_prot < 2:
            system(fr'@del /Q {sys.argv[0]} >nul')  # самоуд flagsn
        exit()
    argv.remove(restart_sign)
# ----------------------------------------------------------------------------
# 0 - без відлагодження
# 1 - без відлагодження, локал cache
# 2 - без перезапуска не батника

# ============================================================================
if __name__ == '__main__':
    system('color F0')
    if debug_prot > 1:
        # При отладке, проверяем всё, в том числе рестарт
        restart(restart_sign)
    elif splitext(sys.argv[0])[-1].lower() == '.bat':
        # Запуск bat-файла, получаем pyc-файл
        py_compile.compile(sys.argv[0])
        fn = split(splitext(sys.argv[0])[0])[-1]
        system(fr'@move /Y __pycache__\{fn}.cpython-38.pyc {fn}.pyc >nul')
        system(r'@rd /Q __pycache__ >nul')
    else:
        # Запуск py-файла, проверяем всё, в том числе рестарт
        restart(restart_sign)
    main()
