# -*- coding,  utf-8 -*-

import QuantLib as ql

def convert_datetime_to_qldate(target_date):
    return ql.Date(target_date.day, target_date.month, target_date.year)

def convert_qldate_to_datetime(target_date):
    return target_date.to_date()

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

def create_schedule_by_termination_date(
    effective_date, termination_date, tenor, cal,
    convention=ql.ModifiedFollowing,
    termination_date_convention=ql.ModifiedFollowing,
    date_generation=ql.DateGeneration.Forward,
    end_of_month=False,
    first_Date=ql.Date(), next_to_lastdate=ql.Date()):

    sche = ql.Schedule(
        effective_date,
        termination_date,
        tenor,
        cal,
        convention,
        termination_date_convention,
        date_generation,
        end_of_month,
        first_Date,
        next_to_lastdate
    )
    return [iter for iter in sche]

def create_schedule_by_termination_period(
    effective_date, termination_period, tenor, cal,
    convention=ql.ModifiedFollowing,
    termination_date_convention=ql.ModifiedFollowing,
    date_generation=ql.DateGeneration.Forward,
    end_of_month=False,
    first_Date=ql.Date(), next_to_lastdate=ql.Date()):

    termination_date = cal.advance(
        effective_date, termination_period,
        termination_date_convention, end_of_month)

    return create_schedule_by_termination_date(
        effective_date, termination_date, tenor, cal,
        convention, termination_date_convention,
        date_generation, end_of_month,
        first_Date, next_to_lastdate
    )
