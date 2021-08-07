import xlwings as xw
import pandas as pd
import QuantLib as ql
import json

_xlsx_path = json.load(
    open('settings.json'), encoding='utf-8')["Data Getter Path"]
book = xw.Book(_xlsx_path)
term_structure = book.sheets['期限结构']


def create_swap(name: str,
                direction: str,
                nominal: float,
                trade_date: str,
                maturity: str,
                first_reset_date: str,
                fixed_rate: float,
                payment_frequency: str,
                fixed_daycount: str,
                spread: int,
                reset_frequency: str,
                floating_daycount: str):

    if payment_frequency != reset_frequency:
        return create_compounding_swap(direction,
                                       nominal,
                                       trade_date,
                                       maturity,
                                       first_reset_date,
                                       fixed_rate,
                                       payment_frequency,
                                       fixed_daycount,
                                       spread,
                                       reset_frequency,
                                       floating_daycount)

    direction = ql.VanillaSwap.Payer if direction == '支付固定' else ql.VanillaSwap.Receiver
    trade_date = ql.Date(trade_date, '%Y/%m/%d')
    maturity = ql.Period(maturity)
    first_reset_date = ql.Date(first_reset_date, '%Y/%m/%d')
    payment_frequency = convert_frequency(payment_frequency)
    fixed_daycount = convert_daycount(fixed_daycount)
    spread = spread * 0.0001
    reset_frequency = convert_frequency(reset_frequency)
    floating_daycount = convert_daycount(floating_daycount)

    calendar = ql.China(ql.China.IB)
    accrural_start_date = calendar.advance(first_reset_date, ql.Period('1D'))
    termination_date = calendar.advance(accrural_start_date, maturity)

    # fixed_schedule = ql.MakeSchedule(first_reset_date,
    #                                  termination_date,
    #                                  frequency=payment_frequency,
    #                                  calendar=calendar,
    #                                  convention=ql.ModifiedFollowing)

    # floating_schedule = ql.MakeSchedule(first_reset_date,
    #                                     termination_date,
    #                                     frequency=payment_frequency,
    #                                     calendar=calendar,
    #                                     convention=ql.ModifiedFollowing)

    # build yield curve
    yts = ql.RelinkableYieldTermStructureHandle()
    index = ql.IborIndex('MyIndex',
                         ql.Period('3m'),
                         1,
                         ql.EURCurrency(),
                         ql.China(ql.China.IB),
                         ql.ModifiedFollowing,
                         True,
                         ql.Actual365Fixed(),
                         yts)

    helpers = ql.RateHelperVector()
    swap_rate = get_swap_curve(name, first_reset_date.ISO().replace('-', '/'))
    for tenor in swap_rate:
        swap_index = ql.EuriborSwapIsdaFixA(ql.Period(tenor))
        rate = swap_rate[tenor]
        helpers.append(ql.SwapRateHelper(rate, swap_index))
    curve = ql.PiecewiseLinearForward(0, calendar, helpers, fixed_daycount)
    yts.linkTo(curve)
    engine = ql.DiscountingSwapEngine(yts)

    swap = ql.MakeVanillaSwap(swapType=direction,
                              effectiveDate=accrural_start_date,
                              terminationDate=termination_date,

                              fixedLegTenor=payment_frequency,
                              floatingLegTenor=payment_frequency,
                              fixedLegCalendar=calendar,
                              floatingLegCalendar=calendar,
                              fixedLegDayCount=fixed_daycount,
                              floatingLegDayCount=floating_daycount,
                              floatingLegSpread=spread,
                              fixedLegConvention=ql.ModifiedFollowing,
                              floatingLegConvention=ql.ModifiedFollowing,
                              floatingLegDateGenRule=ql.DateGeneration.Forward,
                              fixedLegDateGenRule=ql.DateGeneration.Forward,

                              swapTenor=maturity,
                              iborIndex=index,
                              fixedRate=fixed_rate,
                              forwardStart=ql.Period('1D'),
                              Nominal=nominal,
                              pricingEngine=engine)
    swap.yts = yts
    engine = ql.DiscountingSwapEngine(swap.yts)
    swap.setPricingEngine(engine)
    return swap


def create_compounding_swap(direction: str,
                            nominal: float,
                            trade_date: str,
                            maturity: str,
                            first_reset_date: str,
                            fixed_rate: float,
                            payment_frequency: str,
                            fixed_daycount: str,
                            spread: int,
                            reset_frequency: str,
                            floating_daycount: str):
    pass


def convert_frequency(freq: str):
    if freq == 'Q':
        return ql.Period('3M')
    elif freq == 'M':
        return ql.Period('1M')
    elif freq == 'Y':
        return ql.Period('1Y')


def convert_daycount(daycount: str):
    if daycount == '360':
        return ql.Actual360()
    elif daycount == '365':
        return ql.Actual365Fixed()


def get_swap_curve(name: str, date: str):
    book.app.calculation = 'manual'
    term_structure.range('A2').value = name
    term_structure.range('B2').value = date
    book.app.calculate()
    book.app.calculation = 'automatic'

    curve = term_structure.range('A1').expand().options(pd.DataFrame).value
    return curve.loc[:, '6M':'10Y'].T.iloc[:, 0].to_dict()


def get_discount_curve(name: str, date: str):
    pass


def get_fixing_rate(name: str, date: str):
    fixing_rate_sheet = book.sheets['重置利率']
    book.app.calculation = 'manual'
    fixing_rate_sheet.range('A2').value = name
    fixing_rate_sheet.range('B2').value = date
    book.app.calculate()
    book.app.calculation = 'automatic'

    fixing_rate = fixing_rate_sheet.range('C2').value
    return fixing_rate
