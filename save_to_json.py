import sys
from save_to_json_funcs import *


def main(argv):
    """
    Основная функция утилиты.
    Выполняет проверку введенных данных.
    После чего запускает функцию сохранения данных.
    """
    print('Performing checks...')
    if len(argv) > 2:
        return print(f'Error: To many arguments:{len(argv)}. Required: 1-2.\n'
                     f'Please check if the file path have no spaces.')
    slash = define_slash()
    try:
        in_file = argv[0]
    except IndexError:
        return print('Error: Enter excel filename with path.')
    if slash in in_file:
        in_f_path = in_file.rsplit(slash, 1)[0] + slash
    else:
        return print('Error: No path for input file.')
    if not file_exists(in_file):
        return print(f'Error: No such file at {in_f_path}.')

    try:
        out_file = argv[1]
        out_file = define_outfile(out_file, in_file, in_f_path, slash)
        if not out_file:
            return print(f'Error: incorrect output file path or name {argv[1]}')
    except IndexError:
        out_file = in_file.rsplit('.', 1)[0] + '.json'
    if file_exists(out_file):
        return print(f'Error: Such json-file already exists: {out_file}')
    if not file_is_valid(out_file):
        return print(f'Error: Invalid output file: {out_file}')
    print('Saving data to json. Please wait...')
    save_to_json(in_file, out_file)
    print(f'Data was successfully saved at:\n {out_file}')


if __name__ == '__main__':
    main(sys.argv[1:])
