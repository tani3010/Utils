# -*- coding,  utf-8 -*-

import datetime
import tzlocal

def convert_string_to_datetime(target_str, string_format):
    return datetime.datetime.strptime(target_str, string_format)

def convert_datetime_to_string(dt, string_format):
    return dt.strftime(string_format)

def get_local_timezone():
    return tzlocal.get_localzone()

def get_current_time():
    return datetime.now()