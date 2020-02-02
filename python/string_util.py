# -*- coding,  utf-8 -*-

import re

def replace(target_string, from_string, to_string,
            is_regex=False, flags=(re.MULTILINE | re.DOTALL)):
    if is_regex:
        return re.compile(from_string, flags=flags).sub(
            to_string, target_string)
    else:
        return target_string.replace(from_string, to_string)

def split(target_string, delimiter, is_regex=False,
          flags=(re.MULTILINE | re.DOTALL)):
    if is_regex:
        return re.compile(delimiter, flags=flags).split(target_string)
    else:
        return target_string.split(delimiter)

def convert_string_half_to_full(target_string):
    FULL_STRING = ''.join(chr(0xff01 + i) for i in range(94))
    HALF_STRING = ''.join(chr(0x21 + i) for i in range(94))
    HALF_TO_FULL = str.maketrans(HALF_STRING, FULL_STRING)
    return target_string.translate(HALF_TO_FULL)

def convert_string_full_to_half(target_string):
    FULL_STRING = ''.join(chr(0xff01 + i) for i in range(94))
    HALF_STRING = ''.join(chr(0x21 + i) for i in range(94))
    FULL_TO_HALF = str.maketrans(FULL_STRING, HALF_STRING)
    return target_string.translate(FULL_TO_HALF)
