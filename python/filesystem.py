# -*- coding,  utf-8 -*-

import os
from pathlib import Path

def getcwd():
    return os.getcwd()

def chdir(path):
    os.chdir(path)

def is_exist_dir(path):
    return os.path.isdir(path)

def is_exist_file(path):
    return os.path.isfile(path)

def make_dir(path):
    if not is_exist_dir(path):
        os.makedirs(path)

def remove_dir(path):
    if is_exist_dir(path):
        os.removedirs(path)

def get_file_list_in_dir(path, recursive=True, condition=None):
    file_ptr = Path(path)

    if condition is not None:
        return list(file_ptr.glob(condition))

    if recursive:
        return list(file_ptr.glob('**/*'))
    else:
        return list(file_ptr.glob('*'))

def rename_file(from_path, to_path, force_rewrite=False):
    if is_exist_file(from_path) \
        and (force_rewrite or not is_exist_file(to_path)):
        os.rename(from_path, to_path)

def write_string_to_file(
    target_string, file_path, encoding='utf-8'):
    with open(file_path, mode='w', encoding=encoding) as f:
        print(target_string, file=f)