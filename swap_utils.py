import xlwings as xw
import pandas as pd
import QuantLib as ql
import json

_xlsx_path = json.load(
    open('settings.json'), encoding='utf-8')["Data Getter Path"]
book = xw.Book(_xlsx_path)
term_structure = book.sheets['期限结构']
curve_building_method = ql.PiecewiseFlatForward


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

    # 'M' < 'Q' < 'Y'
    if payment_frequency < reset_frequency:
        return create_constant_swap(name,
                                    direction,
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

    if payment_frequency > reset_frequency:
        return create_compounding_swap(name,
                                       direction,
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

    # build yield curve
    yts = ql.RelinkableYieldTermStructureHandle()
    index = ql.IborIndex('LPR',
                         min(payment_frequency, reset_frequency),
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
    curve = curve_building_method(0, calendar, helpers, fixed_daycount)
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
    swap.reset_frequency = reset_frequency
    swap.payment_frequency = payment_frequency
    swap.is_compounded = False
    return swap


def create_compounding_swap(name: str,
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
    # reset_frequency > payment_frequency
    fixed_leg = create_swap(name,
                            direction,
                            nominal,
                            trade_date,
                            maturity,
                            first_reset_date,
                            fixed_rate,
                            payment_frequency,
                            fixed_daycount,
                            spread,
                            payment_frequency,
                            floating_daycount).fixedLeg()

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
    last_reset_date = calendar.advance(first_reset_date, maturity)

    floating_schedule = ql.MakeSchedule(accrural_start_date,
                                        termination_date,
                                        payment_frequency,
                                        calendar=calendar,
                                        forwards=True,
                                        convention=ql.ModifiedFollowing)

    fixing_schedule = ql.MakeSchedule(first_reset_date,
                                      last_reset_date,
                                      reset_frequency,
                                      calendar=calendar,
                                      forwards=True,
                                      convention=ql.ModifiedFollowing)

    yts = ql.RelinkableYieldTermStructureHandle()
    index = ql.IborIndex('LPR',
                         min(payment_frequency, reset_frequency),
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
    curve = curve_building_method(0, calendar, helpers, fixed_daycount)
    yts.linkTo(curve)
    engine = ql.DiscountingSwapEngine(yts)

    floating_leg = ql.SubPeriodsLeg(nominals=[nominal],
                                    schedule=floating_schedule,
                                    index=index,
                                    paymentCalendar=ql.China(),
                                    paymentLag=0,
                                    paymentDayCounter=floating_daycount,
                                    rateSpreads=[spread],
                                    averagingMethod=ql.RateAveraging.Compound)

    swap = ql.Swap(fixed_leg, floating_leg)
    swap.yts = yts
    engine = ql.DiscountingSwapEngine(swap.yts)
    swap.setPricingEngine(engine)
    swap.reset_frequency = reset_frequency
    swap.payment_frequency = payment_frequency
    swap.is_compounded = True
    return swap


def create_constant_swap(name: str,
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
    # reset_frequency < payment_frequency
    fixed_leg = create_swap(name,
                            direction,
                            nominal,
                            trade_date,
                            maturity,
                            first_reset_date,
                            fixed_rate,
                            payment_frequency,
                            fixed_daycount,
                            spread,
                            payment_frequency,
                            floating_daycount).fixedLeg()
    # reset frequency > payment frequency
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
    first_payment_date = calendar.advance(
        accrural_start_date, payment_frequency)
    termination_date = calendar.advance(accrural_start_date, maturity)
    last_reset_date = calendar.advance(first_reset_date, maturity)

    floating_schedule = ql.MakeSchedule(first_payment_date,
                                        termination_date,
                                        payment_frequency,
                                        calendar=calendar,
                                        forwards=True,
                                        convention=ql.ModifiedFollowing)

    fixing_schedule = ql.MakeSchedule(first_reset_date,
                                      last_reset_date,
                                      reset_frequency,
                                      calendar=calendar,
                                      forwards=True,
                                      convention=ql.ModifiedFollowing)

    yts = ql.RelinkableYieldTermStructureHandle()
    index = ql.IborIndex('LPR',
                         min(payment_frequency, reset_frequency),
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
    curve = curve_building_method(0, calendar, helpers, fixed_daycount)
    yts.linkTo(curve)
    engine = ql.DiscountingSwapEngine(yts)

    floating_coupons = []
    for payment_date in floating_schedule:
        last_fixing_date = [y for y in fixing_schedule if y <
                            calendar.advance(payment_date, ql.Period('-7D'))][-1]
        startDate = payment_date - payment_frequency
        endDate = payment_date
        print(endDate)
        coupon = ql.IborCoupon(endDate, nominal, startDate, endDate,
                               calendar.businessDaysBetween(
                                   last_fixing_date, startDate),
                               index)
        coupon.setPricer(ql.BlackIborCouponPricer())
        floating_coupons.append(ql.as_floating_rate_coupon(coupon))

    floating_leg = ql.Leg(floating_coupons)

    swap = ql.Swap(fixed_leg, floating_leg)
    swap.yts = yts
    engine = ql.DiscountingSwapEngine(swap.yts)
    swap.setPricingEngine(engine)
    swap.reset_frequency = reset_frequency
    swap.payment_frequency = payment_frequency
    swap.is_compounded = False
    return swap


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

    curve_sheet = book.sheets[name+'_swap']
    curve = curve_sheet.range('A1').expand().options(pd.DataFrame).value
    curve = curve.loc[:date].iloc[-1].to_frame().T / 100
    curve.columns = [x.split(':')[-1] for x in curve.columns]
    curve.index = [get_fixing_rate(name, date)]
    curve.index.name = '3M'

    term_structure.range('C1').expand().value = curve

    book.app.calculate()
    book.app.calculation = 'automatic'

    curve = term_structure.range('B1').expand().options(pd.DataFrame).value
    return curve.T.iloc[:, 0].astype(float).to_dict()


def get_discount_curve(name: str, date: str):
    pass


def get_fixing_rate(name: str, date: str):
    fixing_rate_sheet = book.sheets['重置利率']
    book.app.calculation = 'manual'
    fixing_rate_sheet.range('A2').value = name
    fixing_rate_sheet.range('B2').value = date

    rate_sheet = book.sheets[name]
    rates = rate_sheet.range('A1').expand().options(pd.DataFrame).value
    fixing_rate_sheet.range(
        'C2').value = rates.loc[:date].iloc[-1].iloc[0] / 100

    book.app.calculate()
    book.app.calculation = 'automatic'

    fixing_rate = fixing_rate_sheet.range('C2').value
    return fixing_rate
