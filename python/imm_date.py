# -*- coding,  utf-8 -*-

import QuantLib as ql

def get_end_of_friday(target_date, cal=ql.UnitedStates()):
    tmp_date = cal.endOfMonth(target_date)
    while not (cal.isBusinessDay(tmp_date) and tmp_date.weekday() == ql.Friday):
        tmp_date = cal.advance(tmp_date, -1, ql.Days, ql.Preceding)
    return tmp_date

def get_nth_weekday(n, weekday, m, y):
    return ql.Date.nthWeekday(n, weekday, m, y)

def get_first_business_date(m, y, cal):
    return cal.adjust(ql.Date(1, m, y))

def get_nth_weeks_weekday(n, weekday, m, y, cal):
    first_biz_date = get_first_business_date(m, y, cal)
    if first_biz_date.weekday() <= weekday:
        return get_nth_weekday(n, weekday, m, y)
    else:
        return get_nth_weekday(n-1, weekday, m, y)

def get_next_imm_date(target_date=ql.Date.todaysDate(), is_major=False):
    return ql.IMM.nextDate(target_date, is_major)

def get_next_imm_code(target_date=ql.Date.todaysDate(), is_major=False):
    return ql.IMM.nextCode(target_date, is_major)