import QuantLib as ql
import swap_utils
from copy import deepcopy


def calculate_vanilla(swap, name, calc_date):
    swap_rate = swap_utils.get_swap_curve(
        name, calc_date)  # export data to exce
    origin_curve = swap.yts.currentLink()

    calc_date = ql.Date(calc_date, '%Y/%m/%d')
    ql.Settings.instance().evaluationDate = calc_date
    calendar = ql.China(ql.China.IB)
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

    for fixing_date in [cf.fixingDate() for cf in map(ql.as_floating_rate_coupon, swap.floatingLeg())]:
        if fixing_date > calc_date:
            break
        date = calc_date.ISO().replace('-', '/')
        index.addFixing(fixing_date, swap_utils.get_fixing_rate(name, date))

    helpers = ql.RateHelperVector()
    for tenor in swap_rate:
        swap_index = ql.EuriborSwapIsdaFixA(ql.Period(tenor))
        rate = swap_rate[tenor]
        helpers.append(ql.SwapRateHelper(rate, swap_index))
    curve = ql.PiecewiseLinearForward(
        0, calendar, helpers, ql.Actual365Fixed())
    swap.yts.linkTo(curve)

    npv = swap.NPV()

    # get DV01
    shift = 0.00005
    helpers = ql.RateHelperVector()
    for tenor in swap_rate:
        swap_index = ql.EuriborSwapIsdaFixA(ql.Period(tenor))
        rate = swap_rate[tenor] + shift
        helpers.append(ql.SwapRateHelper(rate, swap_index))
    up_curve = ql.PiecewiseLinearForward(
        0, calendar, helpers, ql.Actual365Fixed())
    swap.yts.linkTo(up_curve)
    up_swap_npv = swap.NPV()

    helpers = ql.RateHelperVector()
    for tenor in swap_rate:
        swap_index = ql.EuriborSwapIsdaFixA(ql.Period(tenor))
        rate = swap_rate[tenor] - shift
        helpers.append(ql.SwapRateHelper(rate, swap_index))
    down_curve = ql.PiecewiseLinearForward(
        0, calendar, helpers, ql.Actual365Fixed())
    swap.yts.linkTo(down_curve)
    down_swap_npv = swap.NPV()

    dv01 = (up_swap_npv - down_swap_npv) / 10
    swap.yts.linkTo(curve)
    return npv, dv01


def calculate_curvature_adj():
    pass
