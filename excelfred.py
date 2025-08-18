"""
excelfred (alias "xl") - A Python package recreating Excel 514 functions
`Author: Samuel Raj P (FRED)` https://www.linkedin.com/in/samuel-raj23
"""

def __getattr__(name):
    """
    returns all function names starting with that letter in alphabetical order.

    **CODE**:

     # import excelfred as xl; xl.
     # from excelfred import ...
    """
    if len(name) == 1 and name.isalpha(): 
        letter = name.upper(); funcs = []
        for obj_name, obj_value in globals().items():
            if callable(obj_value) and obj_name.upper().startswith(letter): funcs.append(obj_name)
        return sorted(funcs) if funcs else f"No functions starting with '{letter}'"
    raise AttributeError(f"module 'excelfred' has no attribute '{name}'")

#A
def ABS(*args: int | float | str) -> int | float:
    """**=ABS(number)** Returns an Absolute value of a number by taking Modulus. A number without its sign
    
    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ABS(-5)                 # -> 5.0
     ABS(5+2)                # -> 7.0
     ABS(11,7)               # -> 18.0
     ABS("3+2-7")            # -> 2.0
     ABS(-5,"54",-(3.2+6.8)) # -> 69.00
    """
    try:
        total = 0
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"🚫 String Error: Invalid Formula `{arg}` is not evaluatable.")
            try: number = float(arg)
            except: raise TypeError(f"🚫 ABS() only works on numbers or math-like strings. Got `{arg}`.")
            total += abs(number)
        assert args, "Value Error: 🚫 ABS() requires at least one numeric input."
        return total
    except AssertionError as ae: raise ValueError(str(ae))

def ACCRINT(issue: str, first_interest: str, settlement: str, rate: float, par: float = 15000.00, frequency: int = 1, basis: int = 0, calc_method: bool = True) -> int | float:
    """
    `=ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])` Calculates accrued interest for a security that pays periodic interest.

    Parameters:
        issue (str): Issue date (format: 'DD-MM-YYYY')
        first_interest (str): First interest date (format: 'DD-MM-YYYY')
        settlement (str): Settlement date (format: 'DD-MM-YYYY')
        rate (float): Annual interest rate (e.g. 0.1 for 10%)
        par (float): Par value of the security
        frequency (int): Payments per year (1=Annual, 2=Semiannual, 4=Quarterly)
        basis (int, optional): Day count basis (default is 0):
            0 = US (NASD) 30/360
            1 = Actual/actual
            2 = Actual/360
            3 = Actual/365
            4 = European 30/360
        calc_method (bool, optional): 
            If TRUE (default), calculates from issue to settlement date. 
            If FALSE, from first interest to settlement date.

    Returns:
        float: The accrued interest
    """
    import pandas as pd
    try:
        issue_date = pd.to_datetime(issue, dayfirst=True)
        first_date = pd.to_datetime(first_interest, dayfirst=True)
        settle_date = pd.to_datetime(settlement, dayfirst=True)
    except Exception: raise ValueError("🚫 Could not parse one or more dates. Try using common formats like 'DD-MM-YYYY'")
    if settle_date <= issue_date: raise ValueError("🚫 Settlement date must be after issue date.")
    if frequency not in [1, 2, 4]: raise ValueError("🚫 Frequency must be 1 (Annual), 2 (Semi-annual), or 4 (Quarterly).")
    if basis not in range(5): raise ValueError("🚫 Basis must be between 0 and 4.")
    start_date = issue_date if calc_method else first_date
    if settle_date <= start_date: raise ValueError("🚫 Settlement date must be after the start date based on calc_method.")
    delta_days = (settle_date - start_date).days
    if basis == 0 or basis == 4:
        d1 = start_date
        d2 = settle_date
        days = ((d2.year - d1.year) * 360 +
                (d2.month - d1.month) * 30 +
                (d2.day - d1.day))
        year_basis = 360
    elif basis == 1:
        days = delta_days
        leap_start = int(pd.Timestamp(start_date).is_leap_year)
        leap_end = int(pd.Timestamp(settle_date).is_leap_year)
        year_basis = 366 if leap_start or leap_end else 365
    elif basis == 2:
        days = delta_days
        year_basis = 360
    elif basis == 3:
        days = delta_days
        year_basis = 365
    accrued = par * rate * days / year_basis
    return accrued

def ACCRINTM(issue: str, maturity: str, rate: float, par: float = 15000.00, basis: int = 0) -> int | float:
    """
    `=ACCRINTM(issue, settlement maturity, rate, par, [basis])` Returns the accrued interest for a security that pays interest at maturity (no periodic coupons).

    Parameters:
        issue (str): Issue date in 'DD-MM-YYYY' format
        maturity (str): Maturity date in 'DD-MM-YYYY' format
        rate (float): Annual interest rate (e.g., 0.1 for 10%)
        par (float): Par value of the security (default = 15000.00)
        basis (int): Day count basis (0–4). Default is 0:
            0 = US (NASD) 30/360
            1 = Actual/actual
            2 = Actual/360
            3 = Actual/365
            4 = European 30/360

    Returns:
        float: Accrued interest at maturity
    """
    import pandas as pd
    try:
        issue_date = pd.to_datetime(issue, dayfirst=True)
        maturity_date = pd.to_datetime(maturity, dayfirst=True)
    except: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY'.")
    if maturity_date <= issue_date: raise ValueError("Logic Error: 🚫 Maturity date must be after the issue date.")
    delta_days = (maturity_date - issue_date).days
    if basis == 0 or basis == 4:
        d1 = issue_date; d2 = maturity_date
        days = ((d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (d2.day - d1.day))
        year_basis = 360
    elif basis == 1:
        days = delta_days
        leap_start = int(issue_date.is_leap_year)
        leap_end = int(maturity_date.is_leap_year)
        year_basis = 366 if leap_start or leap_end else 365
    elif basis == 2:
        days = delta_days
        year_basis = 360
    elif basis == 3:
        days = delta_days
        year_basis = 365
    else: raise ValueError("String Error: 🚫 Basis must be an integer between 0 and 4.")
    interest = par * rate * days / year_basis
    return interest

def ACOS(*args: int | float | str) -> float:
    """
    **=ACOS(number)** Returns the arccosine (inverse cosine) of a number, in a radians in the range of 0 to pi. Returns angle
    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     print(ACOS(1))           # ➜ 0.0
     print(ACOS(0))           # ➜ 1.5707963268
     print(ACOS(-1))          # ➜ 3.1415926536
     print(ACOS("0.5"))       # ➜ 1.0471975512
     print(ACOS(0.5+0.3))     # ➜ 0.927295218

    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ACOS() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 ACOS() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try:  arg = eval(arg)
                except:  raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Number Error: 🚫 ACOS expects a numeric value or a math-like string. Got `{arg}`")
            if number < -1 or number > 1: raise ValueError("Range Error: 🚫 Input must be between -1 to 1 (inclusive).")
            return np.arccos(number)
    except AssertionError as ae:
        raise ValueError(str(ae))

def ACOSH(*args: int | float | str) -> float:
    """
    **=ACOSH(number)** Returns the inverse hyperbolic cosine of a number. The number must be greater than or equal to 1. 

    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ACOSH(1)            # -> 0.0
     ACOSH("2 + 3")      # -> 2.2924316696
     ACOSH(2 + 3)        # -> 2.2924316696
     ACOSH("5")          # -> 2.2924316696
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ACOSH() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 ACOSH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Evaluate Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 ACOSH expects a numeric value or a math-like string. Got `{arg}`")
            if number < 1: raise ValueError("Number Error: 🚫 Input must be greater than or equal to 1.")
            return np.arccosh(number)
    except AssertionError as ae: raise ValueError(str(ae))

def ACOT(*args: int | float | str) -> float:
    """
    **=ACOT(number)** Returns the arccotangent (inverse cotangent) of a number, if radians ranges 0 to Pi. It returns 1/number.

    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ACOT(1)           # -> 0.7853981634
     ACOT(0.5)         # -> 1.1071487178
     ACOT("1 / 3")     # -> 1.2490457724
     ACOT(0)           # -> 1.5707963268
     ACOT(-5 + 2)      # -> 1.8925468812
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ACOT() requires one numeric input."
        if len(args) != 1: raise ValueError("Value Error: 🚫 ACOT() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Value Error: 🚫 ACOT expects a numeric value or a math-like string. Got `{arg}`")
            if number == 0: return np.pi / 2  # ACOT(0) = π/2

            return np.arctan(1 / number)
    except AssertionError as ae: raise ValueError(str(ae))

def ACOTH(*args: int | float | str) -> float:
    """
    **=ACOTH(number)** Returns the inverse hyperbolic cotangent of a number.
    ACOTH is only defined for values **greater than 1 or less than -1**.

    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ACOTH(2)           # -> 0.5493061443
     ACOTH(-2)          # -> -0.5493061443
     ACOTH(3+2)         # -> 0.5493061443
     ACOTH(-1.5)        # -> -0.8047189562
     ACOTH("5 - 0.5")   # -> 0.2027325541
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ACOTH() requires one numeric input."
        if len(args) != 1: raise ValueError("Argument Error: 🚫 ACOTH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Evaluate Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 ACOTH expects a numeric value or a math-like string. Got `{arg}`")
            if -1 < number < 1: raise ValueError("Number Error: 🚫 Input must be less than or equal to -1 or greater than or equal to 1.")
            
            return 0.5 * np.log((number + 1) / (number - 1))
    except AssertionError as ae: raise ValueError(str(ae))

def ADDRESS(row_num: int, col_num: int, abs_num=1, a1=True, sheet_name=False) -> str:
    """
    `=ADDRESS(row_num, col_num, [abs_num], [a1], [sheet_name])` Create a cell reference as **text**, given specified row and column numbers.
    Parameters:
        row_num (must >=1): Row number
        col_num (must >=1): Column number
        abs_num (optional): Determines the style of absolute/mixed references:
            1 = absolute row & column
            2 = absolute row, relative column 
            3 = relative row, absolute column 
            4 = relative row & column 
        a1 (optional): If True (default A1 Excel-Style ), Else will use R1C1 Style 
        sheet_name (bool, optional): If True, adds 'ExcelFred!' prefix

    *Example Input*:

     print(ADDRESS(1, 1))                             # -> $A$1
     print(ADDRESS(3, 4, abs_num=2))                  # -> D$3
     print(ADDRESS(2, 2, abs_num=4))                  # -> B2
     print(ADDRESS(4, 3, a1=False))                   # -> R4C3
     print(ADDRESS(4, 3, abs_num=4, a1=False))        # -> R[4]C[3]
     print(ADDRESS(5, 5, a1=True, sheet_name=True))   # -> ExcelFred!$E$5
     print(ADDRESS(10, 10, abs_num=3, a1=False))      # -> R[10]C10
    """
    if not isinstance(row_num, int) or not isinstance(col_num, int): raise ValueError("#VALUE!")
    if row_num < 1 or col_num < 1: raise ValueError("🚫 #VALUE!")
    sheet = "ExcelFred!" if sheet_name else ""
    if a1:
        col_letter = ""
        col = col_num
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        
        if abs_num == 1: cell = f"${col_letter}${row_num}"
        elif abs_num == 2: cell = f"{col_letter}${row_num}"
        elif abs_num == 3: cell = f"${col_letter}{row_num}"
        elif abs_num == 4: cell = f"{col_letter}{row_num}"
        else: raise ValueError("🚫 #VALUE!")
    else:
        if abs_num == 1: r = f"R{row_num}"; c = f"C{col_num}"
        elif abs_num == 2: r = f"R{row_num}"; c = f"C[{col_num}]"
        elif abs_num == 3: r = f"R[{row_num}]"; c = f"C{col_num}"
        elif abs_num == 4: r = f"R[{row_num}]"; c = f"C[{col_num}]"
        else: raise ValueError("🚫 #VALUE!")
        cell = r + c
    return sheet + cell

def AGGREGATE(function_num: int, options: int, array, k: float = None, *, hidden: None = None, is_subtotal: None = None) -> float:
    """
    `=AGGREGATE(function_num, options, array, [k])` Returns an **Aggregate** in a list or database.

    Parameters:
        function_num : int  (1..19)  - selects operation from AVERAGE..QUARTILE.EXC, check in examples :)
        options      : int  (0..7)   - ignore controls:
                       0 = ignore nothing, 1 = ignore hidden rows, 2 = ignore errors
                       3 = ignore hidden rows & errors, 4 = ignore nested SUBTOTAL/AGGREGATE
                       5 = ignore hidden rows & nested SUBTOTAL/AGGREGATE
                       6 = ignore errors & nested SUBTOTAL/AGGREGATE
                       7 = ignore hidden rows, errors, nested SUBTOTAL/AGGREGATE
        array        : iterable (list / numpy array / pandas Series) of values (numbers or np.nan for errors)
        k            : optional numeric, used by LARGE(14), SMALL(15), PERCENTILE/QUARTILE (16-19)
        hidden       : optional boolean mask (same length as array). True = hidden row
        is_subtotal  : optional boolean mask (same length as array). True = value produced by SUBTOTAL/AGGREGATE

    Notes:
        - Values that represent "errors" should be np.nan (or produced as np.nan).
        - If options does NOT include 'ignore errors' and any np.nan remains, the function raises ValueError("#VALUE! 🚫 encountered error in data")
        - For exact SUBTOTAL/AGGREGATE ignoring: pass `is_subtotal` mask marking those positions.
        - For QUARTILE/PERCENTILE, `k` semantics:
            * For PERCENTILE.INC (16) and PERCENTILE.EXC (18): k should be between 0..1 (fraction)
            * For QUARTILE.INC (17) and QUARTILE.EXC (19): k should be 0..4 (quart index)
    
    *Example*:

     data = [82, 88, 91, 37, 56, 74, 82, 44, 95, 99, 70, 63, 56, 72, 84, 65, 88, 42, 77]
     print("Average =", AGGREGATE(1, 6, data))                      # 71.84210526315789
     print("Count =", AGGREGATE(2, 6, data))                        # 19
     print("CountA =", AGGREGATE(3, 6, data))                       # 19
     print("Max =", AGGREGATE(4, 6, data))                          # 99
     print("Min =", AGGREGATE(5, 6, data))                          # 37
     print("Product =", AGGREGATE(6, 6, data))                      # 9.439721027123594e+34
     print("Sum =", AGGREGATE(9, 6, data))                          # 1365
     print("Median =", AGGREGATE(10, 6, data))                      # 336.5847953216374
     print("Large(1) =", AGGREGATE(14, 6, data, 1))                 # 99
     print("Small(1) =", AGGREGATE(15, 6, data, 1))                 # 37
     print("Stdev.S =", AGGREGATE(7, 6, data))                      # 18.346247445230794
     print("Stdev.P =", AGGREGATE(8, 6, data))                      # 17.856925997891764
     print("Var.S =", AGGREGATE(11, 6, data))                       # 318.8698060941828
     print("Var.P =", AGGREGATE(12, 6, data))                       # 74
     print("Mode =", AGGREGATE(13, 6, data))                        # 56
     print("Percentile Inc (0.5) =", AGGREGATE(16, 6, data, 0.5))   # 74
     print("Quartile Inc (1) =", AGGREGATE(17, 6, data, 1))         # 59.5
     print("Percentile Exc (0.55) =", AGGREGATE(18, 6, data, 0.55)) # 77
     print("Quartile Exc (1) =", AGGREGATE(19, 6, data, 1))         # 56
    """
    import pandas as pd, numpy as np
    if not isinstance(function_num, int) or not (1 <= function_num <= 19): raise ValueError("#VALUE! 🚫 function_num must be integer 1..19")
    if not isinstance(options, int) or not (0 <= options <= 7): raise ValueError("#VALUE! 🚫 options must be integer 0..7")
    s = pd.Series(array).reset_index(drop=True)
    n = len(s)
    def _validate_mask(name, mask):
        if mask is None: return pd.Series([False]*n)
        mask_s = pd.Series(mask).reset_index(drop=True)
        if len(mask_s) != n: raise ValueError(f"#VALUE! 🚫 {name} mask length must equal array length")
        return mask_s.astype(bool)
    hidden_mask = _validate_mask("hidden", hidden)
    subtotal_mask = _validate_mask("is_subtotal", is_subtotal)
    ignore_hidden = options in (1,3,5,7)
    ignore_errors = options in (2,3,6,7)
    ignore_subtotals = options in (4,5,6,7)
    coerced = pd.to_numeric(s, errors='coerce')  
    s = coerced
    keep_mask = pd.Series([True]*n)
    if ignore_hidden: keep_mask &= ~hidden_mask
    if ignore_subtotals: keep_mask &= ~subtotal_mask
    if ignore_errors: keep_mask &= ~s.isna()
    data = s[keep_mask].to_numpy(dtype=float)
    eff_mask = pd.Series([True]*n)
    if ignore_hidden: eff_mask &= ~hidden_mask
    if ignore_subtotals: eff_mask &= ~subtotal_mask
    eff_data = s[eff_mask]
    if (not ignore_errors) and eff_data.isna().any(): raise ValueError("#VALUE! 🚫 error value in array (set options to ignore errors)")
    if data.size == 0:
        if function_num in (2,3,9): return 0
        if function_num == 6: return 1 
        raise ValueError("#DIV/0! 🚫 no data to aggregate after applying options")
    def _nanmean(x): return np.nanmean(x)
    def _count(x): return int(np.count_nonzero(~np.isnan(x)))
    def _counta(x): return int(np.count_nonzero(~pd.isnull(x)))
    def _nanmax(x): return np.nanmax(x)
    def _nanmin(x): return np.nanmin(x)
    def _nanprod(x): return float(np.prod(x)) 
    def _stdev_s(x): return float(np.nanstd(x, ddof=1))
    def _stdev_p(x): return float(np.nanstd(x, ddof=0))
    def _var_s(x): return float(np.nanvar(x, ddof=1))
    def _var_p(x): return float(np.nanvar(x, ddof=0))
    def _median(x): return float(np.nanmedian(x))
    def _mode(x):
        sr = pd.Series(x).dropna()
        if sr.empty: raise ValueError("#N/A! 🚫 no mode")
        modes = sr.mode()
        if modes.empty: raise ValueError("#N/A! 🚫 no mode")
        return float(modes.iloc[0])
    def _large(x, kk):
        arr = np.sort(x)[~np.isnan(np.sort(x))]
        if kk is None: raise ValueError("#VALUE! 🚫 k is required for LARGE")
        kk = int(kk)
        if kk < 1 or kk > arr.size: raise ValueError("#NUM! 🚫 k out of range for LARGE")
        return float(arr[::-1][kk-1])
    def _small(x, kk):
        arr = np.sort(x)[~np.isnan(np.sort(x))]
        if kk is None: raise ValueError("#VALUE! 🚫 k is required for SMALL")
        kk = int(kk)
        if kk < 1 or kk > arr.size: raise ValueError("#NUM! 🚫 k out of range for SMALL")
        return float(arr[kk-1])
    def _percentile_inc(x, p):
        if p is None: raise ValueError("#VALUE! 🚫 k is required for PERCENTILE.INC (fraction 0..1)")
        p = float(p)
        if not (0 <= p <= 1): raise ValueError("#NUM! 🚫 p must be between 0 and 1 for PERCENTILE.INC")
        arr = np.sort(x[~np.isnan(x)])
        if arr.size == 0: raise ValueError("#DIV/0! 🚫 no data")
        n = arr.size
        if p == 0: return float(arr[0])
        if p == 1: return float(arr[-1])
        rank = 1 + (n - 1) * p
        lower = int(np.floor(rank)) - 1
        upper = int(np.ceil(rank)) - 1
        if lower == upper: return float(arr[lower])
        frac = rank - (lower + 1)
        return float(arr[lower] + frac * (arr[upper] - arr[lower]))
    def _quartile_inc(x, q):
        if q is None: raise ValueError("#VALUE! 🚫 k is required for QUARTILE.INC (0..4)")
        q = int(q)
        if not (0 <= q <= 4): raise ValueError("#NUM! 🚫 quart must be 0..4")
        if q == 0: return _percentile_inc(x, 0.0)
        if q == 4: return _percentile_inc(x, 1.0)
        return _percentile_inc(x, q/4.0)
    def _percentile_exc(x, p):
        if p is None: raise ValueError("#VALUE! 🚫 k is required for PERCENTILE.EXC (fraction 0..1 exclusive)")
        p = float(p)
        n = np.count_nonzero(~np.isnan(x))
        if n < 3: raise ValueError("#NUM! 🚫 PERCENTILE.EXC requires at least 3 data points")
        if not (0 < p < 1): raise ValueError("#NUM! 🚫 p must be between 0 and 1 (exclusive) for PERCENTILE.EXC")
        rank = p * (n + 1)
        if rank <= 1 or rank >= n: raise ValueError("#NUM! 🚫 p out of range for PERCENTILE.EXC")
        arr = np.sort(x[~np.isnan(x)])
        lower = int(np.floor(rank)) - 1
        upper = int(np.ceil(rank)) - 1
        if lower == upper: return float(arr[lower])
        frac = rank - (lower + 1)
        return float(arr[lower] + frac * (arr[upper] - arr[lower]))
    def _quartile_exc(x, q):
        if q is None: raise ValueError("#VALUE! 🚫 k is required for QUARTILE.EXC (0..4)")
        q = int(q)
        if not (0 <= q <= 4): raise ValueError("#NUM! 🚫 quart must be 0..4")
        if q == 0: return _percentile_exc(x, 0.0)  # Excel may error here; keep consistent
        if q == 4: return _percentile_exc(x, 1.0)
        return _percentile_exc(x, q/4.0)
    fm = {
        1: lambda x, kk=None: _nanmean(x),
        2: lambda x, kk=None: _count(x),
        3: lambda x, kk=None: _counta(x),
        4: lambda x, kk=None: _nanmax(x),
        5: lambda x, kk=None: _nanmin(x),
        6: lambda x, kk=None: _nanprod(x),
        7: lambda x, kk=None: _stdev_s(x),
        8: lambda x, kk=None: _stdev_p(x),
        9: lambda x, kk=None: float(np.nansum(x)),
        10: lambda x, kk=None: _var_s(x),
        11: lambda x, kk=None: _var_p(x),
        12: lambda x, kk=None: _median(x),
        13: lambda x, kk=None: _mode(x),
        14: lambda x, kk=None: _large(x, kk),
        15: lambda x, kk=None: _small(x, kk),
        16: lambda x, kk=None: _percentile_inc(x, kk),
        17: lambda x, kk=None: _quartile_inc(x, kk),
        18: lambda x, kk=None: _percentile_exc(x, kk),
        19: lambda x, kk=None: _quartile_exc(x, kk), }
    func = fm.get(function_num)
    if func is None: raise ValueError("#VALUE! 🚫 unsupported function_num")
    result = func(data, k)
    if isinstance(result, (np.floating, np.float64, np.float32)): return float(result)
    if isinstance(result, (np.integer, np.int64, np.int32)): return int(result)
    return result

def AMORLINC(cost: float, date_purchased: str, first_period: str, salvage: float, period: int, rate: float, basis: int = 1) -> float:
    """
    `=AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])` Returns the **linear depreciation** for each accounting period using the French accounting system.
    This method calculates depreciation linearly, prorated if needed for the first period.

    Parameters:
        cost (float): Initial cost of the asset
        date_purchased (str): Date of purchase (format: 'DD-MM-YYYY')
        first_period (str): End of first accounting period (format: 'DD-MM-YYYY')
        salvage (float): Salvage value at the end of asset life
        period (int): The depreciation period number (0 for first, 1 for second...)
        rate (float): Depreciation rate (e.g., 0.15 for 15%)
        basis (int, optional): Day count basis (default = 1):
            0 = 30/360 US
            1 = Actual/actual
            2 = Actual/360
            3 = Actual/365
            4 = 30/360 European

    Example Inputs:

     print(AMORLINC(10000,"01-01-2020","31-12-2020",1000,2,0.2)) # -> 2000.0
    `Returns: float: Depreciation amount for the specified period`
    """
    import pandas as pd
    try:
        purchase_date = pd.to_datetime(date_purchased, dayfirst=True)
        period_end = pd.to_datetime(first_period, dayfirst=True)
    except: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY'.")

    if salvage >= cost: raise ValueError("🚫 Salvage value must be less than cost.")
    if rate <= 0: raise ValueError("🚫 Rate must be a positive value.")

    if basis == 0 or basis == 4:
        d1 = purchase_date; d2 = period_end
        days_in_first = ((d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (d2.day - d1.day))
        year_basis = 360
    elif basis == 1: days_in_first = (period_end - purchase_date).days; year_basis = 366 if purchase_date.is_leap_year or period_end.is_leap_year else 365
    elif basis == 2: days_in_first = (period_end - purchase_date).days; year_basis = 360
    elif basis == 3: days_in_first = (period_end - purchase_date).days; year_basis = 365
    else: raise ValueError("🚫 Basis must be an integer between 0 and 4.")

    depreciation = 0.0
    for i in range(period + 1):
        if i == 0: depreciation = cost * rate * (days_in_first / year_basis)
        else:
            remaining = cost - (cost * rate * (days_in_first / year_basis)) - ((i - 1) * cost * rate)
            if remaining <= salvage: return 0.0
            depreciation = cost * rate
    return depreciation

def AND(*args) -> bool:
    """ 
    `=AND(logical1, [logical2], [logicaln])` Check whether all the arguments are **TRUE**, and returns **TRUE**, if all the arguments are **TRUE**, else return **FALSE**.

    *Example Inputs*:

     print(AND(True, 1, " "))          # -> True
     print(AND(2>=7))                  # -> False
     print(AND(True, 1, 3.14))         # -> True
     print(AND(True, 0, ""))           # -> False
     print(AND(True, None))            # -> True
     print(AND())                      # -> False
    """
    found_valid = False
    for val in args:
        if val is None or (isinstance(val, str) and val.strip() == ''): continue
        found_valid = True
        if isinstance(val, bool):
            if not val: return False
        elif isinstance(val, (int, float)):
            if val == 0: return False
        elif isinstance(val, str): 
            return False  
        else: return False  
    return True if found_valid else False

def ARABIC(input: str) -> int:
    """
    `=ARABIC(text)` Converts a **Roman** Numeral to **Arabic**.

    *Example Inputs*:

     print(ARABIC("XIV"))     # -> 14
     print(ARABIC("MMXXIV"))  # -> 2024
     print(ARABIC("IIII"))    # -> "#VALUE!"
     print(ARABIC("XYZ"))     # -> "#VALUE!"
    """
    roman_numerals = { 'M': 1000, 'CM': 900, 'D': 500, 'CD': 400,
                       'C': 100, 'XC': 90, 'L': 50, 'XL': 40,
                       'X': 10, 'IX': 9, 'V': 5, 'IV': 4, 'I': 1 }
    if not isinstance(input, str): raise ValueError("🚫 #VALUE!") 
    input = input.upper()
    i = 0; result = 0
    valid_roman = ""
    def int_to_roman(n):
        val_map = list(roman_numerals.items()); res = ""
        for roman, val in val_map:
            while n >= val:
                res += roman
                n -= val
        return res
    while i < len(input):
        if i + 1 < len(input) and input[i:i+2] in roman_numerals:
            result += roman_numerals[input[i:i+2]]
            valid_roman += input[i:i+2]
            i += 2
        elif input[i] in roman_numerals:
            result += roman_numerals[input[i]]
            valid_roman += input[i]
            i += 1
        else: raise ValueError("🚫 #VALUE!")
    if int_to_roman(result) != valid_roman: raise ValueError("#VALUE!") 
    return result

def AREAS(*args) -> int:
    """
    `=AREAS(reference)` Returns the number of areas in a reference. An Area is a **range of contiguous cells** or a single cell
    
    *Example Inputs*:

        print(AREAS('A1:B2'))                                   # 1 
        print(AREAS('A1:B2, C3:D4'))                            # 2
        print(AREAS(pd.DataFrame({"A": [1, 2], "B": [3, 4]})))  # 2 
        print(AREAS(pd.Series([1, 2, 3])))                      # 1 
    """
    import pandas as pd, numpy as np
    total = 0
    for arg in args:
        if isinstance(arg, pd.DataFrame): total += len(arg.columns)
        elif isinstance(arg, pd.Series): total += 1
        elif isinstance(arg, np.array): total += 1
        elif isinstance(arg, str):
            cleaned = arg.replace(" ", "")
            refs = cleaned.split(",")
            total += len(refs)
        else: raise ValueError("🚫 #VALUE!") 
    return total

def ARRAYTOTEXT(array, format=0) -> str:
    """
    `=ARRAYTOTEXT(array, [format])` Returns a **text** representation of an **array**.
    
    Parameters:
        array: Enter a list or an array
        format: Decides Consice or Strict
        0 -> Enclosed with **Square**, seperated with **Commas**
        1 -> Enclosed with **Curly Brace**, seperated with **semi-colon**
    
    *Example Inputs*:

     ARRAYTOTEXT(["apple", "banana", "cherry"])             # "apple","banana","cherry"
     ARRAYTOTEXT(["apple", "banana", "cherry"], format=1)   # {"apple";"banana";"cherry"}
     ARRAYTOTEXT([[1, 2], True, False], format=0)           # "1,2","1","0"
     ARRAYTOTEXT([[1, 2], [3, 4], None], format=1)          # {"1,2";"3,4";""}
     ARRAYTOTEXT(pd.Series(["X", "Y", "Z"]), format=True)  # {"X";"Y";"Z"}
     ARRAYTOTEXT(np.array(["A", "B", "C"]), format=False)    # "A","B","C" 
    """
    rows = []
    def to_excel_val(val):
        if val is None: return "" if format == 0 else '""'
        if isinstance(val, bool): return "1" if format == 0 and val else ("0" if format == 0 else ('"1"' if val else '"0"'))
        if isinstance(val, (int, float)): return str(val) if format == 0 else f'"{val}"'
        return str(val) if format == 0 else f'"{val}"'
    if not isinstance(array, (list, tuple)): array = [[array]]
    elif all(not isinstance(row, (list, tuple)) for row in array): array = [array]
    for row in array:
        if not isinstance(row, (list, tuple)): row = [row]
        if format == 0: rows.append(f'"{",".join(to_excel_val(x) for x in row)}"')
        else: rows.append(";".join(to_excel_val(x) for x in row)) 
    if format == 0: return "[" + ",".join(rows) + "]"
    elif format==1: return "{" + ";".join(rows) + "}"
    else: raise ValueError("#VALUE!")

def ASIN(*args: int | float | str) -> float:
    """
    **=ASIN(number)** Returns the arcsine (inverse sine) of a number, in radians, in the range of -pi/2 to pi/2.

    `Parameters: Only one -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ASIN(0)            # -> 0.0
     ASIN("1")          # -> 1.5707963268
     ASIN(-1)           # -> -1.5707963268
     ASIN("0.4+0.1")    # -> 0.5235987756
     ASIN(0.5+0.3)      # -> 0.927295218
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ASIN() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 ASIN() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 ASIN expects a numeric value or a math-like string. Got `{arg}`")
            if number < -1 or number > 1: raise ValueError("Range Error: 🚫 Input must be between -1 to 1 (inclusive).")
            return np.arcsin(number)
    except AssertionError as ae: raise ValueError(str(ae))

def ASINH(*args: int | float | str) -> float:
    """
    **=ASINH(number)** Returns the inverse hyperbolic sine of a number.

    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ASINH(0)           # -> 0.0
     ASINH("1")         # -> 0.881373587
     ASINH(-1)          # -> -0.881373587
     ASINH("2 + 3")     # -> 2.3124383413
     ASINH(-5)          # -> -2.3124383413
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ASINH() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 ASINH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 ASINH expects a numeric value or a math-like string. Got `{arg}`")
            return np.arcsinh(number)
    except AssertionError as ae: raise ValueError(str(ae))

def ATAN(*args: int | float | str) -> float:
    """
    **=ATAN(number)** Returns the arctangent (inverse tangent) of a number, in radians, in the range of -pi/2 to pi/2.

    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ATAN(0)            # -> 0.0
     ATAN("1")          # -> 0.7853981634
     ATAN(-1)           # -> -0.7853981634
     ATAN("1/3")        # -> 0.3217505544
     ATAN(5 - 2)        # -> 1.2490457724
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ATAN() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 ATAN() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 ATAN expects a numeric value or a math-like string. Got `{arg}`")
            return np.arctan(number)
    except AssertionError as ae: raise ValueError(str(ae))

def ATAN2(*args: int | float | str) -> float:
    """
    **=ATAN2(x, y)** Returns the arctangent of the two numbers `x` and `y` (the angle from the X-axis to a point).

    `Parameters: Two # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ATAN2(1, 1)          # -> 0.7853981634
     ATAN2(0, 1)          # -> 0.0
     ATAN2(1, 0)          # -> 1.5707963268
     ATAN2(-1, -1)        # -> -2.3561944902
     ATAN2("3", "4")      # -> 0.643501109
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ATAN2() requires two numeric inputs."
        if len(args) != 2: raise ValueError("Parameters Error: 🚫 ATAN2() requires exactly two inputs.")
        processed = []
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: processed.append(float(arg))
            except: raise TypeError(f"Type Error: 🚫 ATAN2 expects numeric values or math-like strings. Got `{arg}`")
        return np.arctan2(processed[0], processed[1])
    except AssertionError as ae: raise ValueError(str(ae))

def ATANH(*args: int | float | str) -> float:
    """
    **=ATANH(number)** Returns the inverse hyperbolic tangent of a number. Input must be between -1 and 1 (exclusive).

    `Parameters: Only one # -> Number (but accepts arithmetics too) int | float | str`

    *Example inputs*:

     ATANH(0)            # -> 0.0
     ATANH(0.5)          # -> 0.5493061443
     ATANH(-0.5)         # -> -0.5493061443
     ATANH("0.8")        # -> 1.0986122887
     ATANH("-0.9")       # -> -1.4722194896
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 ATANH() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 ATANH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 ATANH expects a numeric value or a math-like string. Got `{arg}`")
            if number <= -1 or number >= 1: raise ValueError("Range Error: 🚫 Input must be greater than -1 and less than 1.")
            return np.arctanh(number)
    except AssertionError as ae: raise ValueError(str(ae))

def AVEDEV(*args) -> float:
    """
    `=AVEDEV(num_1, [num_2] ...)`
    
    Returns the **average of the absolute deviations** of data points from their mean.
    Accepts numbers, booleans, arrays, and numeric strings.
    
    *`Parameters: Multiple -> Number (but accepts arithmetics too) int | float | bool | str`*
    
    *Example Inputs*:
    
     print(AVEDEV([1, 2, 3, 4, 5]))                # 1.2
     print(AVEDEV(True, False))                    # 0.5
     print(AVEDEV("2+3", "4"))                     # 0.5
     print(AVEDEV([2*5, "12"], 14, "TRUE"))        # 4.125
     print(AVEDEV("apple", "banana"))              # 🚫 #DIV/0!
    """
    import numpy as np
    values = []
    def extract_numbers(item):
        if isinstance(item, (list, tuple)):
            for sub in item:
                extract_numbers(sub)
        elif isinstance(item, bool): values.append(1 if item else 0)
        elif isinstance(item, (int, float)): values.append(item)
        elif isinstance(item, str):
            s = item.strip()
            if not s: return
            if s.lower() == "true": values.append(1)
            elif s.lower() == "false": values.append(0)
            else:
                try: values.append(float(s))
                except ValueError: pass
        elif item is None: pass
    for arg in args: extract_numbers(arg)
    if not values: raise ValueError("🚫 #DIV/0!")
    arr = np.array(values, dtype=float)
    mean_val = np.mean(arr)
    abs_dev = np.abs(arr - mean_val)
    return np.mean(abs_dev)

def AVERAGE(*args) -> float:
    """
    `=AVERAGE(num_1, [num_2] ...)`
    
    Returns the average **(arithmetic mean)** of its arguments, which can be a number or names or array, or references that contains number.
    
    *`Parameters: Multiple -> Number (but accepts arithmetics too) int | float | str`*

    *Example Inputs*:

     print(AVERAGE(1, 2, 3))                         # 2.0
     print(AVERAGE([1.5, True, False, "apple"]))     # 1.5
     print(AVERAGE("TRUE", "FALSE", "None"))         # 0.5
     print(AVERAGE(22/7, "", None, [3.14, 7]))       # 4.427619047619047
     print(AVERAGE("apple", "banana"))               # #DIV/0!
     print(AVERAGE([1, 2], (3, 4), "5"))             # 3
    """
    values = []
    def extract_numbers(item):
        if isinstance(item, (list, tuple)):
            for sub in item:
                extract_numbers(sub)
        elif isinstance(item, bool): values.append(1 if item else 0)
        elif isinstance(item, (int, float)): values.append(item)
        elif isinstance(item, str):
            s = item.strip()
            if not s:  return
            if s.lower() == "true": values.append(1)
            elif s.lower() == "false": values.append(0)
            else:
                try: values.append(float(s))
                except ValueError: pass 
        elif item is None: pass
    for arg in args: extract_numbers(arg)
    if not values: raise ValueError("🚫 #DIV/0!")
    return sum(values) / len(values)

def AVERAGEA(*values) -> float:
    """
    `=AVERAGEA(val_1, [val_2] ...)`

    Returns the average **(arithmetic mean)** of its arguments, evaluating texts and FALSE in arguments as 0; TRUE evaluates as 1, Arguments can be a number or names or array, or references that contains number.
    
    *`Parameters: Multiple -> Number (but accepts arithmetics too) int | float | str`*
    
    *Example Inputs*:

     print(AVERAGEA(10, 10*2, "text", True, False))       # (10 + 20 + 0 + 1 + 0) / 5 = 6.2
     print(AVERAGEA([1, 2, "abc", True], False, None))    # (1 + 2 + 0 + 1 + 0) / 5 = 0.8
     print(AVERAGEA("apple", "banana"))                   # (0 + 0) / 2 = 0.0
    """
    import numpy as np
    processed = []
    for v in values:
        if isinstance(v, (list, tuple, np.ndarray)): processed.extend(v)
        else: processed.append(v)
    converted = []
    for v in processed:
        if isinstance(v, bool): converted.append(1 if v else 0)
        elif isinstance(v, (int, float)) and not isinstance(v, bool): converted.append(v)     
        elif v is None: continue
        else: converted.append(0)
    if len(converted) == 0: return np.nan
    return np.mean(converted)

def AVERAGEIF(range_vals, criteria, average_range=None) -> float:
    """
    `=AVERAGEIF(range, criteria, [average_range])`
    Finds Average **(arithemetic mean)** for the cells specified by a given condition or criteria 

    Parameters:
        range :
            *The range of cells to evaluate with the criteria.*
        criteria :
            *The condition that determines which cells are averaged.*
            *Can be a number, expression (e.g., ">5"), or text (e.g., "Apple").*
        average_range (optional) :
            *The actual set of cells to average. If omitted, `range` is used.*

    *Example Inputs*:

     print(AVERAGEIF([10, 20, 30, 40], ">15"))                           # (20 + 30 + 40) / 3 = 30
     print(AVERAGEIF([5, 7, 9, 11], "<10", [50, 60, 70, 80]))            # (50 + 60 + 70) / 3 = 60
     print(AVERAGEIF(["Apple", "Orange", "Apple"], "Apple", [5, 6, 7]))  # (5 + 7) / 2 = 6
     print(AVERAGEIF([True, False, True], True, [10, 20, 30]))           # (10 + 30) / 2 = 20
     print(AVERAGEIF([1, 2, 3], ">5"))                                   # No match → NaN
    """
    import pandas as pd, numpy as np
    range_vals = np.array(range_vals)
    if average_range is None: average_range = range_vals
    else: average_range = np.array(average_range)
    s = pd.Series(range_vals)
    if criteria.startswith((">", "<", "=")): mask = s.astype(float).apply(lambda v: eval(f"v{criteria}"))
    else: mask = s == criteria
    filtered = average_range[mask]
    return np.mean(filtered) if len(filtered) > 0 else np.nan

def AVERAGEIFS(average_range, *criteria_pairs) -> float:
    """
    `=AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2] ...)`
    Finds Average **(arithemetic mean)** for the cells specified by a given set of condition or criteria.

    Parameters:
        average_range : 
            *The actual set of cells to average.*

        criteria_rangeN : array-like
            *The range of cells to evaluate against criteria 'N.*

        criteriaN : str | int | float
            *The condition that determines which cells are included. Can be a number, expression (e.g., ">5"), or text (e.g., "Apple").*

    *Example Inputs*:
     print(AVERAGEIFS([50, 60, 70, 80], [1, 2, 3, 4], ">1"))                       # (60 + 70 + 80) / 3 = 70
     print(AVERAGEIFS([50, 60, 70, 80], [1, 2, 3, 4], ">1", [5, 6, 7, 8], "<8"))   # (60 + 70) / 2 = 65
     print(AVERAGEIFS([5, 6, 7], ["A", "B", "A"], "A", [True, False, True], True)) # (5 + 7) / 2 = 6
     print(AVERAGEIFS([10, 20, 30, 40], [2, 4, 6, 8], ">3", [1, 2, 3, 4], "<4"))   # (20 + 30) / 2 = 25
     print(AVERAGEIFS([100, 200, 300], [True, True, False], True))                 # (100 + 200) / 2 = 150
    """
    import pandas as pd, numpy as np
    average_range = np.array(average_range)
    mask = np.ones(len(average_range), dtype=bool)
    
    for i in range(0, len(criteria_pairs), 2):
        crit_range = np.array(criteria_pairs[i])
        criteria = criteria_pairs[i+1]
        mask &= pd.Series(crit_range).apply(lambda x: eval(f"x{criteria}" if criteria[0] in "<>=" else f"x=={repr(criteria)}"))
    filtered = average_range[mask]
    return np.mean(filtered) if len(filtered) > 0 else np.nan

#B
def BAHTTEXT(number: int | float) -> str:
    """
    `=BAHTTEXT(number)` Converts a number to a text **(baht)**

    *Example Inputs*:
     
     print(BAHTTEXT(1234.56))  # หนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบหกสตางค์
     print(BAHTTEXT(5000))     # ห้าพันบาทถ้วน
    """
    number = round(float(number), 2)
    thai_numbers = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า']
    thai_positions = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน', 'ล้าน']
    def num_to_thai_words(n):
        n = int(n)
        if n == 0: return 'ศูนย์'
        words = ''
        position = 0
        while n > 0:
            digit = n % 10
            if digit != 0:
                if position == 0 and digit == 1 and words != '': words = 'เอ็ด' + words
                elif position == 1 and digit == 2: words = 'ยี่' + thai_positions[position] + words
                elif position == 1 and digit == 1: words = thai_positions[position] + words
                else: words = thai_numbers[digit] + thai_positions[position] + words
            position += 1
            n //= 10
        return words
    baht = int(number)
    satang = int(round((number - baht) * 100))
    baht_words = num_to_thai_words(baht) + 'บาท'
    if satang == 0: satang_words = 'ถ้วน'
    else: satang_words = num_to_thai_words(satang) + 'สตางค์'
    return baht_words + satang_words

def BASE(number: int, radix: int, min_length: int = 0) -> str:
    """
    `=BASE(number, radix, [min_length])` Converts a *number* into a text representation with the given **radix** (base)
    
    Parameters:
        number: → integer in base 10
        target_base: → (2 to 36) just means the numbering system you want to convert to:
            2 = binary (only 0,1)
            8 = octal (0–7)
            10 = decimal (normal numbers 0–9)
            16 = hexadecimal (0–9, A–F)
            Up to 36 = can use digits 0–9 and letters A–Z as symbols.
        min_length: → optional zero-padding length for output.

    *Example Input*:

     print(BASE(255, 16))          # FF
     print(BASE(255, 2, 12))       # 000011111111
     print(BASE(-255, 16, 4))      # -00FF
     print(BASE(123456, 36))       # 2N9C

    `So higher radix more symbols available for writing numbers.`
    """
    if not (2 <= radix <= 36): raise ValueError("#VALUE! 🚫 Radix must be between 2 and 36")
    if number == 0: return "0".zfill(min_length)
    digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    is_negative = number < 0
    number = abs(number)
    result = ""
    while number > 0:
        result = digits[number % radix] + result
        number //= radix
    if is_negative: result = "-" + result
    if len(result.lstrip("-")) < min_length: result = ("-" if is_negative else "") + result.lstrip("-").zfill(min_length)
    return result

def BESSELI(x, n) -> float:
    """
    `=BESSELI(x, n)` Returns the modified Bessel function **In(x)**

    Parameters:
        x (float): The value at which to evaluate the function.
        n (int): The order of the Bessel function (must be >= 0).

    *Example Input*:

        BESSELI(5, 2)   # 17.505614966624236
        BESSELI(10, 0)  # 2815.7166284662544
        BESSELI(3, 1)   # 3.95337021740261
        BESSELI(1, 0)   # 1.2660658777520084
        BESSELI(0.5, 0) # 1.0634833707413236
        BESSELI(7, 3)   # 85.1754868428438
        BESSELI(20, 5)  # 23018392.213413667
        BESSELI(50, 2)  # 2.8164306402451954e+20

    `Returns (float) The calculated Bessel I_n(x) value.`
    """
    from scipy.special import iv
    if n<0: raise ValueError("#VALUE! 🚫 'n' should be above or equal to zero !")
    else: return iv(n, x)

def BESSELJ(x, n) -> float:
    """
    `=BESSELJ(x, n)` Returns the Bessel function of first kind **jn(x)**.

    Parameters:
        x (float): The value at which to evaluate the function.
        n (int): The order of the Bessel function (must be >= 0).

    *Example Input*:

        BESSELJ(5, 2)   # 0.04656511627775229
        BESSELJ(10, 0)  # -0.24593576445134832
        BESSELJ(3, 1)   # 0.33905895852593626
        BESSELJ(1, 0)   # 0.7651976865579666
        BESSELJ(0.5, 0) # 0.938469807240813
        BESSELJ(7, 3)   # -0.16755558799533432
        BESSELJ(20, 5)  # 0.15116976798239493
        BESSELJ(50, 2)  # -0.05971280079425883
    """
    from scipy.special import jv
    if n<0: raise ValueError("#VALUE! 🚫 'n' should be above or equal to zero !")
    else: return jv(n, x)

def BESSELK(x, n) -> float:
    """
    `=BESSELK(x, n)` Returns the Bessel function of first kind **kn(x)**.

    Parameters:
        x (float): The value at which to evaluate the function.
        n (int): The order of the Bessel function (must be >= 0).

    *Example Input*:

       BESSELK(5, 2)   # 0.00530894371222346
       BESSELK(10, 0)  # 1.778006231616765e-05
       BESSELK(3, 1)   # 0.040156431128194184
       BESSELK(1, 0)   # 0.42102443824070834
       BESSELK(0.5, 0) # 0.9244190712276656
       BESSELK(7, 3)   # 0.0007710751535668902
       BESSELK(20, 5)  # 1.0538660139974233e-09
       BESSELK(50, 2)  # 3.547931838858198e-23
    """    
    from scipy.special import kv
    if n<0: raise ValueError("#VALUE! 🚫 'n' should be above or equal to zero !")
    else: return kv(n, x)

def BESSELY(x, n) -> float:
    """
    `=BESSELY(x, n)` Returns the Bessel function of first kind **Yn(x)**.

    Parameters:
        x (float): The value at which to evaluate the function.
        n (int): The order of the Bessel function (must be >= 0).

    *Example Input*:

     BESSELY(5, 2)    # 0.3676628826055246
     BESSELY(10, 0)   # 0.05567116728359934
     BESSELY(3, 1)    # 0.32467442479180014
     BESSELY(1, 0)    # 0.088256964215677
     BESSELY(0.5, 0)  # -0.44451873350670656
     BESSELY(7, 3)    # 0.26808060304231507
     BESSELY(20, 5)   # -0.10003576788953246
     BESSELY(50, 2)   # 0.09579316872759651
    """
    from scipy.special import yv
    if n<0: raise ValueError("#VALUE! 🚫 'n' should be above or equal to zero !")
    else: return yv(n, x)

def BIN2DEC(num: int | float | str) -> int:
    """
    `=BIN2DEC(num)` Converts a binary number to decimal.
    
    *Example Inputs*:

     print(BIN2DEC("1010"))    # 10
     print(BIN2DEC(1010))      # 10
     print(BIN2DEC("   110 ")) # 6
     print(BIN2DEC("1021"))    # ValueError
     print(BIN2DEC(12.5))      # TypeError

    `Any other String or Numbers except 0 or 1 return Error !`
    """
    if not isinstance(num, (str, int)): raise TypeError(" #VALUE! 🚫 BIN2DEC accepts only int or str inputs.")
    if isinstance(num, int): num_str = str(num)
    else: num_str = num.strip()
    if not num_str.isdigit() or any(ch not in "01" for ch in num_str): raise ValueError("#VALUE! 🚫 Input must contain only binary digits (0 or 1).")
    return int(num_str, 2)

def BIN2HEX(num) -> str:
    """
    `=BIN2HEX(num)` Converts a binary number to hexadecimal.
    
    *Example usage*:

     print(BIN2HEX("1010"))    # 'A'
     print(BIN2HEX(1011))      # 'B'
     print(BIN2HEX(" 1111 "))  # 'F'

    `Any other String or Numbers except 0 or 1 return Error !`
    """
    if not isinstance(num, (int, str)): raise TypeError("#VALUE! 🚫 BIN2HEX accepts only int or str.")
    bin_str = str(num).strip()
    if not bin_str: raise ValueError("#VALUE! 🚫 Empty input is not a valid binary number.")
    if not all(ch in '01' for ch in bin_str): raise ValueError("#VALUE! 🚫 Binary number must contain only 0s and 1s.")
    decimal_value = int(bin_str, 2)
    hex_value = hex(decimal_value)[2:].upper()
    return hex_value

def BIN2OCT(num: int | float | str) -> int:
    """
    `=BIN2OCT(num)` Convert a binary number to its octal representation.
    
    *Example usage*:

     print(BIN2OCT("1010"))    # 12
     print(BIN2OCT(1011))      # 15
     print(BIN2OCT(" 1111 "))  # 70
     print(BIN2OCT(111))       # 7

    `Any other String or Numbers except 0 or 1 return Error !`
    """
    bin_str = str(num).strip()
    if not bin_str or any(ch not in "01" for ch in bin_str): raise ValueError("#VALUE! 🚫 Input must be a binary number containing only 0 and 1.")
    decimal_value = int(bin_str, 2)
    octal_str = format(decimal_value, "o")
    return octal_str

def BETA_DIST(x: float, alpha: float, beta_param: float, cumulative: bool, A:float=0, B:float=1) -> float:
    """
    `=BETA.DIST(x, alpha, beta, cumulative, [A], [B])` Returns the beta probability distribution function
    
    Parameters:
        x: Value at which to evaluate the function
        alpha: Shape parameter α > 0
        beta_param: Shape parameter β > 0
        cumulative: True for CDF, False for PDF
        A (optional): Lower bound of interval (default 0)
        B (optional): Upper bound of interval (default 1)
    
    *Example Input*:

     print(BETA_DIST(0.5, 2, 3, True))                 # 0.6875
     print(BETA_DIST(0.5, 2, 3, False))                # 1.5000000000000004
     print(BETA_DIST(7, 2, 3, True, A=0, B=10))        # 0.9163
     print(BETA_DIST(7, 2, 3, False, A=0, B=10))       # 0.07559999999999999   
    """
    from scipy.stats import beta
    if B <= A: raise ValueError("#VALUE! 🚫 B must be greater than A")
    if not (A <= x <= B): raise ValueError(f"#VALUE! 🚫 x must be between {A} and {B}")
    z = (x - A) / (B - A)
    if cumulative: return beta.cdf(z, alpha, beta_param)
    else: return beta.pdf(z, alpha, beta_param) / (B - A)  

def BETA_INV(probability: float, alpha: float, beta: float, A: float = 0, B: float = 1) -> float:
    """
    `=BETA.INV(probability, alpha, beta, [A], [B])` Returns the inverse of the beta cumulative distribution function for a specified probability.

    Parameters:
        probability: The probability (0 ≤ p ≤ 1)
        alpha: Shape parameter α > 0
        beta_param: Shape parameter β > 0
        A (optional): Lower bound of interval (default 0)
        B (optional): Upper bound of interval (default 1)

    **Example Input**:

        print(BETA_INV(0.5, 2, 3))               # 0.385727...
        print(BETA_INV(0.95, 2, 3))              # 0.773...
        print(BETA_INV(0.95, 2, 3, 0, 10))       # 7.732...

    """
    from scipy.stats import beta
    if B <= A: raise ValueError("#VALUE! 🚫 B must be greater than A")
    if not (0 <= probability <= 1): raise ValueError("#VALUE! 🚫 probability must be between 0 and 1")
    result = beta.ppf(probability, alpha, beta)
    return A + result * (B - A)

def BINOM_DIST(number_s: int, trials: int, probability_s: float, cumulative=True) -> float:
    """
    `=BINOM.DIST(number_s, trials, probability_s, is_cummulative)` Returns the individual term of binomial distribution.

    Parameters:
        x : number of successes
        n : number of trials
        p : probability of success (0 <= p <= 1)
        cumulative : True for cumulative distribution, False for pmf

    *Example usage*:

     print(BINOM_DIST(2, 10, 0.5, False))  # 0.04394531250000004
     print(BINOM_DIST(2, 10, 0.5, True))   # 0.0546875
    """
    from scipy.stats import binom
    if cumulative: return binom.cdf(number_s, trials, probability_s)
    else: return binom.pmf(number_s, trials, probability_s)

def BINOM_DIST_RANGE(trials: int, probability_s: float, num_s: int, num_s2: int = None) -> float:
    """
    `=BINOM.DIST.RANGE(trials, probability_s, num_s, [num_s2])` Returns the probability of a trial result falling between two thresholds.

    Parameters:
        trials       : Number of independent trials (n >= 0)
        probability_s: Probability of success on each trial (0 <= p <= 1)
        num_s        : Number of successes for lower bound
        num_s2       : (Optional) Upper bound of successes. If omitted, only P(X = num_s) is returned.

    **Example Inputs**:

         print(BINOM_DIST_RANGE(60, 0.75, 45))         # 0.11822800461154298
         print(BINOM_DIST_RANGE(60, 0.75, 45, 50))     # 0.5236297934718878
    """
    from scipy.stats import binom
    if not (0 <= probability_s <= 1): raise ValueError("#VALUE! 🚫 probability_s must be between 0 and 1")
    if trials < 0 or num_s < 0 or (num_s2 is not None and num_s2 < 0): raise ValueError("#VALUE! 🚫 trials and successes must be non-negative integers")
    if num_s > trials or (num_s2 is not None and num_s2 > trials): raise ValueError("#VALUE! 🚫 successes cannot exceed number of trials")
    if num_s2 is not None and num_s2 < num_s: raise ValueError("#VALUE! 🚫 2nd number must be greater than or equal to First Number")
    if num_s2 is None: return binom.pmf(num_s, trials, probability_s)
    else: return binom.cdf(num_s2, trials, probability_s) - binom.cdf(num_s - 1, trials, probability_s)

def BINOM_INV(trials: int, probability_s: float, alpha: float) -> int:
    """
    `=BINOM.INV(trials, probability_s, alpha)`
    Returns the smallest number of successes for which the cumulative binomial
    distribution is greater than or equal to alpha.

    Parameters:
        trials       : Number of independent trials (n >= 0)
        probability_s: Probability of success on each trial (0 <= p <= 1)
        alpha        : Cumulative probability threshold (0 <= alpha <= 1)

    *Example Input*:

         print(BINOM_INV(6, 0.5, 0.75))   # 4
         print(BINOM_INV(10, 0.3, 0.9))   # 5
    """
    from scipy.stats import binom
    if not (0 <= probability_s <= 1): raise ValueError("#VALUE! 🚫 probability_s must be between 0 and 1")
    if not (0 <= alpha <= 1): raise ValueError("#VALUE! 🚫 alpha must be between 0 and 1")
    if trials < 0: raise ValueError("#VALUE! 🚫 trials must be non-negative integer")
    return int(binom.ppf(alpha, trials, probability_s))

def BITAND(number1: int, number2: int) -> int:
    """
    `=BITAND(number1, number2)` Returns a **bitwise** AND of two numbers.

    *Example Input*:

        print(BITAND(5, 3))    # 1  (0101 AND 0011 = 0001)
        print(BITAND(12, 25))  # 8  (1100 AND 11001 = 01000)
    """
    if number1 < 0 or number2 < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number1, int) and isinstance(number2, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return number1 & number2

def BITOR(number1: int, number2: int) -> int:
    """
    `=BITOR(number1, number2)` Returns a **bitwise** OR of two numbers.

    *Example Input*:

        print(BITOR(5, 3))    # 7  (0101 OR 0011 = 0111)
        print(BITOR(12, 25))  # 29 (1100 OR 11001 = 11101)
    """
    if number1 < 0 or number2 < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number1, int) and isinstance(number2, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return number1 | number2

def BITLSHIFT(number: int, shift_amount: int) -> int:
    """
    `=BITLSHIFT(number, shift_amount)` Returns a number **shifted left** by a given number of bits.

    *Example Input*:

        print(BITLSHIFT(5, 2))   # 20  (0101 << 2 = 10100)
        print(BITLSHIFT(12, 3))  # 96  (1100 << 3 = 1100000)
    """
    if number < 0 or shift_amount < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number, int) and isinstance(shift_amount, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return number << shift_amount

def BITRSHIFT(number: int, shift_amount: int) -> int:
    """
    `=BITRSHIFT(number, shift_amount)` Returns a number **shifted right** by a given number of bits.

    *Example Input*:

        print(BITRSHIFT(20, 2))  # 5   (10100 >> 2 = 0101)
        print(BITRSHIFT(96, 3))  # 12  (1100000 >> 3 = 1100)
    """
    if number < 0 or shift_amount < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number, int) and isinstance(shift_amount, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return number >> shift_amount

def BITNAND(number1: int, number2: int) -> int:
    """
    `=BITXAND(number1, number2)` Returns a **bitwise NOT AND** of two numbers.

    *Example Input*:

        print(BITXAND(5, 3))    # 14 (NOT(0101 AND 0011) = NOT(0001) = 1110)
        print(BITXAND(12, 25))  # 23 (NOT(1100 AND 11001) = NOT(01000) = 10111)
    """
    if number1 < 0 or number2 < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number1, int) and isinstance(number2, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return ~(number1 & number2) & ((1 << max(number1.bit_length(), number2.bit_length())) - 1)

def BITNOR(number1: int, number2: int) -> int:
    """
    `=BITNOR(number1, number2)` Returns a **bitwise NOT OR** of two numbers.

    *Example Input*:

        print(BITNOR(5, 3))    # 8  (NOT(0101 OR 0011) = NOT(0111) = 1000)
        print(BITNOR(12, 25))  # 2  (NOT(1100 OR 11001) = NOT(11101) = 00010)
    """
    if number1 < 0 or number2 < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number1, int) and isinstance(number2, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return ~(number1 | number2) & ((1 << max(number1.bit_length(), number2.bit_length())) - 1)

def BITXAND(number1: int, number2: int) -> int:
    """
    `=BITXAND(number1, number2)` Returns a **bitwise** XAND (equivalence) of two numbers.

    *Example Input*:

        print(BITXAND(5, 3))    # 9  (0101 XAND 0011 = 1001)
        print(BITXAND(12, 25))  # 10 (01100 XAND 11001 = 01010)
    """
    if number1 < 0 or number2 < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number1, int) and isinstance(number2, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return ~(number1 ^ number2) & ((1 << max(number1.bit_length(), number2.bit_length())) - 1)

def BITXOR(number1: int, number2: int) -> int:
    """
    `=BITXOR(number1, number2)` Returns a **bitwise** XOR of two numbers.

    *Example Input*:

        print(BITXOR(5, 3))    # 6  (0101 XOR 0011 = 0110)
        print(BITXOR(12, 25))  # 21 (1100 XOR 11001 = 10101)
    """
    if number1 < 0 or number2 < 0: raise ValueError("#NUM! 🚫 numbers must be non-negative integers")
    if not (isinstance(number1, int) and isinstance(number2, int)): raise ValueError("#VALUE! 🚫 numbers must be integers")
    return number1 ^ number2

#C
def CEILING_MATH(number: int | float, significant: int | float = 1, mode: int = 0) -> float:
    """
    `=CEILING.MATH(number, [significant], [mode])`
    Rounds a number up, to the nearest integer or the nearest multiple of significance.

    Parameters:
        number : The number to round
        significant : Multiple to round to (default 1, must be > 0)
        mode : int — For negative numbers: 
                     0 = round away from zero (default)
                     nonzero = round toward zero

    *Example Input*:

        print(CEILING_MATH(4.3))             # 5
        print(CEILING_MATH(-4.3))            # -4
        print(CEILING_MATH(-4.3, 2))         # -4
        print(CEILING_MATH(-4.3, 2, 1))      # -2
        print(CEILING_MATH(4.3, 2))          # 6
        print(CEILING_MATH(4.3, 0.5))        # 4.5
    """
    import numpy as np
    if significant < 0: raise ValueError("#NUM! 🚫 significant must be positive")
    if significant == 0: return 0.0
    sign = np.sign(number) 
    return np.floor(number / significant) * significant if sign < 0 and mode != 0 else np.ceil(np.abs(number) / significant) * significant * sign

def CELL(info_type: str, reference: any) -> any:
    """
    `=CELL(info_type, reference)` Returns information about the formatting, location, or contents of the first cell, according to the sheet's reading order, in a reference .

    **Parameters**:
        **info_type**
            The type of cell information you want to retrieve **(case-insensitive)** :
                 "address"       → Returns the cell address as an absolute reference (e.g., "$B$2").
                 "column"        → Returns the column number of the reference (A=1, B=2, ...).
                 "contents"      → Returns the value/content of the reference cell.
                 "format"        → Returns a code representing the cell's number format ("G" for General, "A" for Alphanumeric).
                 "parentheses"   → Returns TRUE if the cell value is formatted with parentheses (negative values), FALSE otherwise.
                 "prefix"        → Returns the label prefix of the cell ("'", "\"", "^", or "") used for text alignment in Excel.
                 "protect"       → Returns TRUE if the cell is locked/protected (simulated TRUE).
                 "row"           → Returns the row number of the reference (1-based).
                 "type"          → Returns "b" if blank, "l" if label (text), "v" if value (number).
                 "width"         → Returns the approximate column width (simulated as length of string representation).

    `reference : scalar value, NP array, or PD cell (The cell to retrieve information from)`

    **Example Input**:

        df = pd.DataFrame({"A":[10, 20, np.nan], "B":["x", "y", "z"]})
        print(CELL("address", df.iloc[1,1]))        # "$B$2"
        print(CELL("column", df.iloc[1,1]))         # 2
        print(CELL("contents", df.iloc[1,1]))       # "y"
        print(CELL("format", df.iloc[0,0]))         # "G"
        print(CELL("parentheses", -50))             # True
        print(CELL("prefix", "Hello"))              # "'"
        print(CELL("protect", df.iloc[0,0]))        # True
        print(CELL("row", df.iloc[1,1]))            # 2
        print(CELL("type", df.iloc[0,0]))           # "v"
        print(CELL("width", df.iloc[0,0]))          # 2

    """
    import pandas as pd, numpy as np
    info_type = info_type.lower()
    valid_info = {"address", "col", "contents", "format", "parentheses", "prefix", "protect", "row", "type", "width" }
    if info_type not in valid_info: raise ValueError("#VALUE! 🚫 Invalid info_type")
    if isinstance(reference, pd.Series):
        if len(reference) != 1: raise ValueError("#VALUE! 🚫 reference must be 1 cell")
        val = reference.iloc[0]; col_label = reference.name; row_label = reference.index[0]
    elif isinstance(reference, pd.DataFrame):
        if reference.shape != (1, 1): raise ValueError("#VALUE! 🚫 reference must be 1×1 DataFrame")
        val = reference.iat[0, 0]; col_label = reference.columns[0]; row_label = reference.index[0]
    else: val = reference; col_label = "A"; row_label = 1
    def col_letter(col):
        if isinstance(col, int): n = col
        elif isinstance(col, str) and col.isalpha(): return col.upper()
        else:
            try: n = int(col) + 1
            except: n = 1
        letters = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return letters
    if info_type == "address": return f"${col_letter(col_label)}${row_label}"
    elif info_type == "col": return (ord(col_letter(col_label)[0]) - 64)
    elif info_type == "contents": return val
    elif info_type == "format":
        if isinstance(val, (int, float, np.number)): return "G"  # General
        elif isinstance(val, str): return "@"
        else: return "G"
    elif info_type == "parentheses": return 1 if isinstance(val, (int, float, np.number)) and val < 0 else 0
    elif info_type == "prefix":
        if isinstance(val, str):
            if val.startswith("'"): return "'"
            elif val.startswith('"'): return '"'
            else: return ""
        return ""
    elif info_type == "protect": return True
    elif info_type == "row": return row_label if isinstance(row_label, int) else 1
    elif info_type == "type":
        if val is None or (isinstance(val, float) and np.isnan(val)): return "b"
        elif isinstance(val, str): return "l"
        else: return "v"
    elif info_type == "width": return 8.43  
    raise ValueError("#VALUE! 🚫 Unsupported info_type")

def CHAR(number: int) -> str:
    """
    `=CHAR(number)` Returns the character specified by a number according to the ANSI/Unicode character set.

    **Example Input**:

        print(CHAR(65))    # "A"   → ANSI code 65 is uppercase A
        print(CHAR(97))    # "a"   → ANSI code 97 is lowercase a
        print(CHAR(48))    # "0"   → ANSI code 48 is digit zero
        print(CHAR(36))    # "$"   → ANSI code 36 is dollar sign
        print(CHAR(10))    # "\n"  → ANSI code 10 is line feed (newline)

    `**Parameters**: An integer between 1 and 255 representing the ANSI code of the desired character.`

    **Excel Notes**:
        - On Windows, `CHAR()` uses the ANSI character set (code page 1252 by default).
        - If you want to handle Unicode values above 255 in Excel, you must use `excelfred.UNICHAR()`.
    """
    import numpy as np
    if not isinstance(number, (int, np.integer)): raise ValueError("#VALUE! 🚫 number must be an integer")
    if number < 1 or number > 255: raise ValueError("#VALUE! 🚫 number must be between 1 and 255")
    return chr(number)

def CHISQ_DIST(x: float, deg_freedom: int, cumulative: bool) -> float:
    """
    `=CHISQ.DIST(x, deg_freedom, cumulative)` Returns the **left-tailed** probability of the **chi-squared** distribution.

    Parameters:
        x : float → Value at which to evaluate the distribution (must be ≥ 0).
        deg_freedom : int → Degrees of freedom (must be ≥ 1).
        cumulative : bool → TRUE for cumulative distribution (CDF), FALSE for probability density (PDF).

    *Example Input*:

         print(CHISQ_DIST(0.5, 1, True))    # 0.5204998778
         print(CHISQ_DIST(0.5, 1, False))   # 0.4393912895
         print(CHISQ_DIST(2, 2, True))      # 0.6321205588
         print(CHISQ_DIST(2, 2, False))     # 0.1839397206
         print(CHISQ_DIST(5, 10, True))     # 0.0954659664
    """
    from scipy import stats
    if x < 0: raise ValueError("#NUM! 🚫 x must be non-negative")
    if deg_freedom < 1: raise ValueError("#NUM! 🚫 degrees of freedom must be ≥ 1")
    if cumulative: return stats.chi2.cdf(x, deg_freedom)
    else: return stats.chi2.pdf(x, deg_freedom)

def CHISQ_DIST_RT(x: float, deg_freedom: int) -> float:
    """
    `=CHISQ.DIST.RT(x, deg_freedom)` Returns the **right-tailed** probability of the **chi-squared** distribution.

    Parameters:
        x : float → Value at which to evaluate (must be ≥ 0).
        deg_freedom : int → Degrees of freedom (must be ≥ 1).

    *Example Input*:

         print(CHISQ_DIST_RT(0.5, 1))   # 0.4795001222
         print(CHISQ_DIST_RT(2, 2))     # 0.3678794412
         print(CHISQ_DIST_RT(5, 10))    # 0.9045340337
         print(CHISQ_DIST_RT(15, 20))   # 0.8282028557
         print(CHISQ_DIST_RT(30, 25))   # 0.2424253566
    """
    from scipy import stats
    if x < 0: raise ValueError("#NUM! 🚫 x must be non-negative")
    if deg_freedom < 1: raise ValueError("#NUM! 🚫 degrees of freedom must be ≥ 1")
    return stats.chi2.sf(x, deg_freedom)

def CHISQ_INV(prob: float, deg_freedom: int) -> float:
    """
    `=CHISQ.INV(prob, deg_freedom)` Returns the **inverse** of the **left-tailed** probability of the chi-squared distribution.

    Parameters:
        prob : float → Probability (0 < prob < 1).
        deg_freedom : int → Degrees of freedom (must be ≥ 1).

    *Example Input*:

        print(CHISQ_INV(0.5204998778, 1))   # 0.49999999997030736
        print(CHISQ_INV(0.6321205588, 2))   # 1.9999999998447435
        print(CHISQ_INV(0.0954659664, 10))  # 4.793572111719336
        print(CHISQ_INV(0.5, 5))            # 4.351460191
        print(CHISQ_INV(0.9, 3))            # 6.251389
    """
    from scipy import stats
    if not (0 < prob < 1): raise ValueError("#NUM! 🚫 prob must be between 0 and 1")
    if deg_freedom < 1: raise ValueError("#NUM! 🚫 degrees of freedom must be ≥ 1")
    return stats.chi2.ppf(prob, deg_freedom)

def CHISQ_INV_RT(prob: float, deg_freedom: int) -> float:
    """
    `=CHISQ.INV.RT(prob, deg_freedom)` Returns the **inverse** of the **right-tailed** probability of the chi-squared distribution.

    Parameters:
        prob : float → Probability (0 < prob < 1).
        deg_freedom : int → Degrees of freedom (must be ≥ 1).

    *Example Input*:

        print(CHISQ_INV_RT(0.4795001222, 1))  # 0.49999999997030786
        print(CHISQ_INV_RT(0.3678794412, 2))  # 1.9999999998447444
        print(CHISQ_INV_RT(0.9045340337, 10)) # 4.793572110121122
        print(CHISQ_INV_RT(0.5, 5))           # 4.351460191
        print(CHISQ_INV_RT(0.1, 3))           # 6.251389
    """
    from scipy import stats
    if not (0 < prob < 1): raise ValueError("#NUM! 🚫 prob must be between 0 and 1")
    if deg_freedom < 1: raise ValueError("#NUM! 🚫 degrees of freedom must be ≥ 1")
    return stats.chi2.isf(prob, deg_freedom)

def CHISQ_TEST(test_range, expected_range) -> float:
    """
    `=CHISQ.TEST(test_range, expected_range)` Returns the test for independence from the chi-squared distribution.

    Parameters:
        test_range : array-like → Observed data.
        expected_range : array-like → Expected values.

    *Example Input*:

        observed = np.array([[10, 20, 30], [6,  9,  17]])
        expected = np.array([[8,  18, 34], [8, 11, 13]])
        print(CHISQ_TEST(observed, expected))  # 0.606
    """
    import numpy as np
    from scipy import stats
    observed = np.array(test_range, dtype=float)
    expected = np.array(expected_range, dtype=float)
    if observed.shape != expected.shape: raise ValueError("#N/A 🚫 observed and expected ranges must have the same dimensions")
    chi_sq = ((observed - expected) ** 2 / expected).sum()
    df = (observed.shape[0] - 1) * (observed.shape[1] - 1)
    return stats.chi2.sf(chi_sq, df)

def CHOOSE(index_num: int, *values: any) -> any:
    """
    `=CHOOSE(index_num, val1, [val2], ..., [valN])` Returns the value from a list of values based on the given index.

    Parameters:
        index_num :
            The position of the value to return. (1-based index, same as Excel)
        *values :
            One or more values among which to choose.

    Raises:
        ValueError: If index_num is not between 1 and len(values), or if index_num is not an integer.

    *Example Input*:

        print(CHOOSE(1, "Apple", "Banana", "Cherry"))   # Apple
        print(CHOOSE(3, 10, 20, 30, 40))                # 30
        print(CHOOSE(2, True, False))                   # False
        print(CHOOSE(4, "A", "B", "C", "D", "E"))       # D
        print(CHOOSE(1, 5.5, 6.6, 7.7))                 # 5.5
    """
    import numpy as np
    if not isinstance(index_num, (int, np.integer)): raise ValueError("#VALUE! 🚫 index_num must be an integer")
    if index_num == 0: raise TypeError("#NUM! 🚫 Index in CHOOSE starts from 1, Unlike usual array format")
    if index_num < 1 or index_num > len(values): raise ValueError("#VALUE! 🚫 index_num is out of range")
    return values[index_num - 1]

def CODE(text: str) -> int:
    """
    `=CODE(text)` Returns a numeric code for the first character of a text string.
    In Excel, this is based on **ANSI/Unicode** codes depending on version.

    *Example Input*:
    
        print(CODE("A"))       # 65   (ASCII for 'A')
        print(CODE("a"))       # 97   (ASCII for 'a')
        print(CODE("😀"))      # 128512 (Unicode for 😀)
        print(CODE("Excel"))   # 69   (ASCII for 'E')
        print(CODE("1"))       # 49   (ASCII for '1')

    `Parameters: The text string from which to return the code of the first character.`
    """
    if not isinstance(text, str): raise ValueError("#VALUE! 🚫 text must be a string")
    if text == "": raise ValueError("#VALUE! 🚫 text cannot be empty")
    return ord(text[0])

def COLUMN(reference: any) -> int | list:
    """
    `=COLUMN(reference)` Returns the **column number** of a reference

    *Example Inputs*:

     print("COLUMN('A') →", COLUMN('A'))                # 1
     print("COLUMN('AA') →", COLUMN('AA'))              # 27
     arr = np.array([[1, 2, 3], [4, 5, 6]])
     print("COLUMN(arr) →", COLUMN(arr))                # [1, 2, 3] 

     df = pd.DataFrame({ 'Name': ['Alice', 'Bob'],
     'Age': [25, 30], 'City': ['NY', 'LA'] })
     print("COLUMN(df) →", COLUMN(df))                  # [1, 2, 3]
     print("COLUMN(df['City']) →", COLUMN(df['City']))  # 3

    `Parameter - Accepts reference in dataframe, series, array, list formats `
    """
    import pandas as pd, numpy as np
    if isinstance(reference, str):
        reference = reference.strip().upper()
        col_num = 0
        for char in reference:
            if not 'A' <= char <= 'Z': raise ValueError("#NUM! 🚫 Invalid column letter.")
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        return col_num    
    if isinstance(reference, pd.Series):
        if reference.name is None: raise ValueError("#VALUE! 🚫 Series has no column name.")
        col_names = reference.to_frame().columns.tolist()
        return col_names.index(reference.name) + 1    
    if isinstance(reference, pd.DataFrame): return list(range(1, len(reference.columns) + 1))    
    if isinstance(reference, np.ndarray):
        if reference.ndim == 1: return 1 
        return list(range(1, reference.shape[1] + 1))
    raise TypeError("#REF! 🚫 Unsupported reference type.")

def COLUMNS(array: list | dict) -> int:
    """
    `=COLUMNS(array)` Returns the number of columns in an array or a reference

    *Example Inputs*:

     print("COLUMNS(df) =>", COLUMNS(pd.DataFrame({'A': [1,2,3],'B': [4,5,6], 'C': [7,8,9]})))  # 3
     print("COLUMNS(lst) =>", COLUMNS([[1,2],[3,4]]))                                           # 2
    
    `Parameter - Accepts reference in dataframe, series, array, list formats `
    """    
    import pandas as pd, numpy as np
    if isinstance(array, pd.DataFrame): return array.shape[1]
    elif isinstance(array, pd.Series): return 1    
    elif isinstance(array, np.ndarray): return 1 if array.ndim == 1 else array.shape[1] 
    elif isinstance(array, list):
        if len(array) == 0: return 0
        elif isinstance(array[0], list): return len(array[0])
        else: return 1
    else: raise TypeError("#REF! 🚫 Unsupported type for COLUMNS function.")

def COMBIN(number: int, number_chosen: int) -> int:
    """
    `=COMBIN(number, number_chosen)` Returns the **number of combinations** for a given number of items.

    Parameters:
        number         → Total number of items (n)
        number_chosen  → Number of items to choose (k)

    *Example Inputs*:

         print(COMBIN(10, 2))    # 45
         print(COMBIN(5, 3))     # 10
         print(COMBIN(8, 0))     # 1
         print(COMBIN(6, 6))     # 1
         print(COMBIN(6, 1))     # 6
    """
    from scipy.special import comb
    if not (isinstance(number, int) and isinstance(number_chosen, int)): raise ValueError("#VALUE! 🚫 Parameters must be integers.")
    if number < 0 or number_chosen < 0: raise ValueError("#NUM! 🚫 Parameters must be non-negative.")
    if number_chosen > number: raise ValueError("#NUM! 🚫 number_chosen cannot be greater than number.")
    return int(comb(number, number_chosen, exact=True))

def COMBINA(number: int, number_chosen: int) -> int:
    """
    `=COMBINA(number, number_chosen)` Returns the **number of combinations with repetitions** allowed.

    `Parameters`:
        number         → Total number of items (n)
        number_chosen  → Number of items to choose (k)

    *Example Inputs*:

         print(COMBINA(10, 2))    # 55
         print(COMBINA(5, 3))     # 35
         print(COMBINA(8, 0))     # 1
         print(COMBINA(6, 6))     # 462
         print(COMBINA(6, 1))     # 6
    """
    from scipy.special import comb
    if not (isinstance(number, int) and isinstance(number_chosen, int)): raise ValueError("#VALUE! 🚫 Parameters must be integers.")
    if number <= 0 or number_chosen < 0: raise ValueError("#NUM! 🚫 number must be > 0 and number_chosen must be non-negative.")
    return int(comb(number + number_chosen - 1, number_chosen, exact=True))

def COMPLEX(real_num: float, img_num: float, suffix: str="i") -> str:
    """
    `=COMPLEX(real_num, i_num, [suffix])` Converts real and imaginary **co-efficients** into complex numbers

    Parameters:
        real_num: 
            All Natural numbers
        imaginary_num: 
            ends with i (denoted as root of -1)
        [suffix]: (optional)
            Imaginary number unit ends with either 'i' or 'j'

    *Example Inputs*:
    
     print(COMPLEX(3, 4))          # 3+4i
     print(COMPLEX(0, -2))         # 0-2i
     print(COMPLEX(-5, 0))         # -5+0i
     print(COMPLEX(0, 0))          # 0+0i
     print(COMPLEX(2.5, 3.7, "j")) # 2.5+3.7j
    """
    import numpy as np
    if suffix not in ("i", "j"): raise ValueError("#VALUE! 🚫 Suffix must be 'i' or 'j'")
    if np.isnan(real_num) or np.isnan(img_num): return np.nan
    if real_num == 0 and img_num == 0: return f"0{suffix}"
    real_str = str(int(real_num)) if real_num == int(real_num) else str(real_num)
    imag_str = str(abs(int(img_num)) if img_num == int(img_num) else abs(img_num))    
    sign = "+" if img_num >= 0 else "-"
    return f"{real_str}{sign}{imag_str}{suffix}"

def CONCAT(*args: any) -> str:
    """
    `=CONCAT(text1, ...)` **Concatenates** a list or a range of text strings.

    *Example Inputs*:

     print(CONCAT("excel","fred"))                         # excelfred
     print(CONCAT(1, 2, 3))                                # 123
     print(CONCAT(True, False))                            # TRUEFALSE
     print(CONCAT(np.array([1, None, "X"])))               # 1X
     print(CONCAT(pd.Series([1, np.nan, 3])))              # 13
     print(CONCAT(range(3)," ", "Done"))                   # 012 Done

    `parameters - accepts any type of list arrays series ranges that in integer/float/string `
    """
    import numpy as np, pandas as pd
    result_parts = []
    for arg in args:
        if isinstance(arg, (pd.Series, pd.DataFrame)): values = arg.values.flatten()
        elif isinstance(arg, (np.ndarray, list, tuple, set, range)): values = np.array(arg, dtype=object).flatten()
        else: values = [arg]
        for v in values:
            if v is None or (isinstance(v, float) and pd.isna(v)): continue
            if isinstance(v, bool): result_parts.append("TRUE" if v else "FALSE")
            else: result_parts.append(str(v))
    return "".join(result_parts)

def CONFIDENCE_NORM(alpha, std_dev, size) -> float:
    """
    `=CONFIDENCE.NORM(alpha, std_dev, size)`

    **Parameter**:
        **Alpha(α)**: 
            Significance level (probability of error, e.g., 0.05 for 95% confidence)
        **Standard Deviation**:
            Standard deviation of the Population (sigma)
        **Size(n)**: 
            Number of observations (AKA sample size)

    Example Inputs:

    print(CONFIDENCE.NORM(0.05,2.5,50))   # 0.6929519121748389
    """
    import numpy as np; from scipy import stats
    if isinstance(alpha, str) or isinstance(std_dev, str) or isinstance(size, str): raise ValueError(f"🚫 String Error: Invalid Datatype, Enter float or int instead.")
    if size <= 0 or std_dev < 0: raise ValueError("size must be > 0 and std_dev >= 0")
    z = stats.norm.ppf(1 - alpha / 2)  
    return z * (std_dev / np.sqrt(size))

def CONFIDENCE_T(alpha, std_dev, size):
    """
    `=CONFIDENCE.T(alpha, std_dev, size)`

    **Parameter**:
        **Alpha(α)**: 
            Significance level (probability of error, e.g., 0.05 for 95% confidence)
        **Standard Deviation**:
            Standard deviation of the sample.
        **Size(n)**: 
            Number of observations (AKA sample size)

    Example Inputs:

     print(CONFIDENCE.T(0.05,2.5,50))   # 0.7104921387393247
    """
    import numpy as np; from scipy import stats
    if isinstance(alpha, str) or isinstance(std_dev, str) or isinstance(size, str): raise ValueError(f"🚫 String Error: Invalid Datatype, Enter float or int instead.")
    if size <= 1 or std_dev < 0: raise ValueError("size must be > 1 and std_dev >= 0")
    t = stats.t.ppf(1 - alpha / 2, df=size - 1)  
    return t * (std_dev / np.sqrt(size))    

def CONVERT(number: float, from_unit: str, to_unit: str) -> float:
    """
    `=CONVERT(num, from_unit, to_unit)` Converts a number from one **measurement system** to another.

    **Supported Units**:
    `Length: m, km, cm, mm, um, nm, pm, fm, in, ft, yd, mi, nmi, ang, ly, pc, fath, ch, rd`
    `Mass: g, kg, mg, ug, st, lbm, oz, t, ton, cwt`
    `Volume: l, ml, m3, ft3, in3, gal, qt, pt, cup, floz`
    `Time: s, min, hr, day, yr, mo`
    `Pressure: pa, atm, bar, torr, psi`
    `Energy: j, kj, cal, kcal, wh, kwh, ev, btu`
    `Power: w, kw, mw, hp`
    `Area: m2, km2, cm2, mm2, ft2, yd2, in2, ac, ha`
    `Angle: rad, deg, grad, gon`
    `Temperature: c, f, k, r`
         
    **Parameter**:
     Number:
        The numeric value you want to convert from one measurement system to another.
        (Example: 10 for 10 meters)
     From_Unit:
        The text abbreviation of the unit you are converting from.
        (Example: "m" for meters, "lbm" for pounds mass)
     To_Unit:
        The text abbreviation of the unit you are converting to.
        (Example: "ft" for feet, "kg" for kilograms)

    *Example Inputs*:
    
        print(CONVERT(20, 'cwt', 'kg'))                   # 907.18474
        print(CONVERT(20, 'cwt', 'g'))                    # 907184.74
        print(CONVERT(2204.622621848776, 'g', 'lbm'))     # ~4.8573561563
        print(CONVERT(100, 'km', 'm'))                    # 100000
        print(CONVERT(5, 'mi', 'km'))                     # 8.04672
        print(CONVERT(1, 'ft', 'm'))                      # 0.3048
        print(CONVERT(12, 'in', 'cm'))                    # 30.48
        print(CONVERT(2, 'yd', 'm'))                      # 1.8288
        print(CONVERT(1, 'ly', 'm'))                      # 9.4607e+15
        print(CONVERT(3, 'pc', 'ly'))                     # 9.7821
        print(CONVERT(60, 'min', 's'))                    # 3600
        print(CONVERT(1, 'hr', 'min'))                    # 60
        print(CONVERT(1, 'day', 'hr'))                    # 24
        print(CONVERT(101325, 'pa', 'atm'))               # 1
        print(CONVERT(14.7, 'psi', 'pa'))                 # ~101352.83
        print(CONVERT(500, 'cal', 'j'))                   # 2092
        print(CONVERT(1, 'hp', 'w'))                      # 745.6998715822702
        print(CONVERT(300, 'kg', 'g'))                    # 300000

        # Area units
        print(CONVERT(10000, 'm2', 'ac'))                 # ~2.47105381
        print(CONVERT(2, 'km2', 'm2'))                    # 2,000,000
        print(CONVERT(500, 'cm2', 'm2'))                  # 0.05
        print(CONVERT(100, 'ft2', 'm2'))                  # 9.290304
        print(CONVERT(1, 'ha', 'm2'))                     # 10000

        # Angle units
        print(CONVERT(180, 'deg', 'rad'))                 # 3.141592653589793 (pi)
        print(CONVERT(200, 'grad', 'deg'))                # 180
        print(CONVERT(100, 'gon', 'rad'))                 # 1.5707963267948966 (pi/2)
        print(CONVERT(3.141592653589793, 'rad', 'deg'))   # 180

        # Temperature
        print(CONVERT(0, 'c', 'f'))                       # 32
        print(CONVERT(32, 'f', 'c'))                      # 0
        print(CONVERT(0, 'k', 'c'))                       # -273.15
        print(CONVERT(491.67, 'r', 'f'))                  # 32 
    """
    length_units = {"m":1,"km":1000,"cm":0.01,"mm":0.001,"um":1e-6,"nm":1e-9,"pm":1e-12,"fm":1e-15, "in":0.0254,"ft":0.3048,"yd":0.9144,"mi":1609.344,"nmi":1852,"ang":1e-10, "ly":9.4607e15,"pc":3.0857e16,"fath":1.8288,"ch":20.1168,"rd":5.0292 }
    mass_units = {"g":1,"kg":1000,"mg":0.001,"ug":1e-6,"st":6350.29318,"lbm":453.59237,"oz":28.349523125, "t":1e6,"ton":907184.74,"cwt":45359.237  }
    volume_units = {"l":1,"ml":0.001,"m3":1000,"ft3":28.316846592,"in3":0.016387064, "gal":3.785411784,"qt":0.946352946,"pt":0.473176473,"cup":0.24,"floz":0.0295735295625 }
    time_units = {"s":1,"min":60,"hr":3600,"day":86400,"yr":31557600,"mo":2629800}
    pressure_units = {"pa":1,"atm":101325,"bar":100000,"torr":133.322368,"psi":6894.757293168}
    energy_units = {"j":1,"kj":1000,"cal":4.184,"kcal":4184,"wh":3600,"kwh":3.6e6,"ev":1.60218e-19,"btu":1055.06}
    power_units = {"w":1,"kw":1000,"mw":1e6,"hp":745.69987158227022}
    area_units = {"m2":1,"km2":1e6,"cm2":1e-4,"mm2":1e-6,"ft2":0.09290304,"yd2":0.83612736, "in2":0.00064516,"ac":4046.8564224,"ha":10000 }
    angle_units = {"rad":1,"deg":0.0174532925199433,"grad":0.015707963267948967,"gon":0.015707963267948967}
    aliases = {"lb": "lbm", "lbs": "lbm", "liter": "l", "litre": "l", "meters": "m", "metre": "m", "metres": "m", "seconds": "s", "sec": "s", "hrs": "hr", "hour": "hr", "hours": "hr", "degrees": "deg", "radians": "rad", "gallons": "gal", "pounds": "lbm", "tons": "ton", "tons_us": "ton", "tons_uk": "t", "acres": "ac", "stone": "stone", "oz": "ozm", "grain": "grain", "u": "u", "cwt": "cwt", "shweight": "cwt", "uk_cwt": "uk_cwt", "lcwt": "uk_cwt", "hweight": "uk_cwt", "ton": "ton", "uk_ton": "uk_ton", "LTON": "uk_ton", "brton": "uk_ton", "mi": "mi", "nmi": "nmi", "in": "in", "ft": "ft", "yd": "yd", "ang": "ang", "ell": "ell", "ly": "ly", "parsec": "pc", "pc": "pc", "Picapt": "pica", "Pica": "pica", "pica": "pica", "survey_mi": "survey_mi", "yr": "yr", "day": "day", "d": "day", "hr": "hr", "mn": "min", "min": "min", "sec": "s", "atm": "atm", "at": "atm", "mmHg": "mmHg", "psi": "psi", "kg": "kg", "g": "g", "mg": "mg", "ug": "ug", "st": "st", "lbm": "lbm", "ozm": "ozm", "grain": "grain", "u": "u", "cwt": "cwt", "shweight": "cwt", "uk_cwt": "uk_cwt", "lcwt": "uk_cwt", "hweight": "uk_cwt", "ton": "ton", "uk_ton": "uk_ton", "LTON": "uk_ton", "brton": "uk_ton", "mi": "mi", "nmi": "nmi", "in": "in", "ft": "ft", "yd": "yd", "ang": "ang", "ell": "ell", "ly": "ly", "parsec": "pc", "pc": "pc", "Picapt": "pica", "Pica": "pica", "pica": "pica", "survey_mi": "survey_mi", "yr": "yr", "day": "day", "d": "day", "hr": "hr", "mn": "min", "min": "min", "sec": "s", "atm": "atm", "at": "atm", "mmHg": "mmHg", "psi": "psi", "kg": "kg", "g": "g", "mg": "mg", "ug": "ug", "st": "st", "lbm": "lbm", "ozm": "ozm", "grain": "grain", "u": "u", "cwt": "cwt", "shweight": "cwt", "uk_cwt": "uk_cwt", "lcwt": "uk_cwt", "hweight": "uk_cwt", "ton": "ton", "uk_ton": "uk_ton", "LTON": "uk_ton", "brton": "uk_ton", "mi": "mi", "nmi": "nmi", "in": "in", "ft": "ft", "yd": "yd", "ang": "ang", "ell": "ell", "ly": "ly", "parsec": "pc", "pc": "pc", "Picapt": "pica", "Pica": "pica", "pica": "pica", "survey_mi": "survey_mi", "yr": "yr", "day": "day", "d": "day", "hr": "hr", "mn": "min", "min": "min", "sec": "s", "atm": "atm", "at": "atm", "mmHg": "mmHg", "psi": "psi", "kg": "kg", "g": "g", "mg": "mg", "ug": "ug", "st": "st", "lbm": "lbm", "ozm": "ozm", "grain": "grain", "u": "u", "cwt": "cwt", "shweight": "cwt", "uk_cwt": "uk_cwt", "lcwt": "uk_cwt", "hweight": "uk_cwt", "ton": "ton", "uk_ton": "uk_ton", "LTON": "uk_ton", "brton": "uk_ton", "mi": "mi", "nmi": "nmi", "in": "in", "ft": "ft", "yd": "yd", "ang": "ang", "ell": "ell", "ly": "ly", "parsec": "pc", "pc": "pc", "Picapt": "pica", "Pica": "pica", "pica": "pica", "survey_mi": "survey_mi", "yr": "yr", "day": "day", "d": "day", "hr": "hr", "mn": "min", "min": "min", "sec": "s", "atm": "atm", "at": "atm", "mmHg": "mmHg", "psi": "psi", "kg": "kg", "g": "g", "mg": "mg", "ug": "ug", "st": "st", "lbm": "lbm", "ozm": "ozm", "grain": "grain", "u": "u", "cwt": "cwt", "shweight": "cwt", "uk_cwt": "uk_cwt", "lcwt": "uk_cwt", "hweight": "uk_cwt", "ton": "ton", "uk_ton": "uk_ton", "LTON": "uk_ton", "brton": "uk_ton", "mi": "mi", "nmi": "nmi", "in": "in", "ft": "ft", "yd": "yd", "ang": "ang", "ell": "ell", "ly": "ly", "parsec": "pc", "pc": "pc", "Picapt": "pica", "Pica": "pica", "pica": "pica", "survey_mi": "survey_mi", "yr": "yr", "day": "day", "d": "day", "hr": "hr", "mn": "min", "min": "min", "sec": "s", "atm": "atm", "at": "atm", "mmHg": "mmHg", "psi": "psi", "kg": "kg", "g": "g", "mg": "mg", "ug": "ug", "st": "st", "lbm": "lbm", "ozm": "ozm", "grain": "grain", "u": "u", "cwt": "cwt", "shweight": "cwt", "uk_cwt": "uk_cwt", "lcwt": "uk_cwt", "hweight": "uk_cwt", "ton": "ton", "uk_ton": "uk_ton", "LTON": "uk_ton", "brton": "uk_ton", "mi": "mi", "nmi": "nmi", "in": "in", "ft": "ft", "yd": "yd", "ang": "ang", "ell": "ell", "ly": "ly", "parsec": "pc", "pc": "pc", "Picapt": "pica", "Pica": "pica", "pica": "pica", "survey_mi": "survey_mi", "yr": "yr", "day": "day", "d": "day", "hr": "hr", "mn": "min", "min": "min", "sec": "s", "atm": "atm", "at": "atm", "mmHg": "mmHg", "psi": "psi"}
    temperature_units = {"c","f","k","r"}
    def convert_temperature(value, from_u, to_u):
        if from_u == to_u: return value
        if from_u == "c": c = value
        elif from_u == "f": c = (value - 32) * 5/9
        elif from_u == "k": c = value - 273.15
        elif from_u == "r": c = (value - 491.67) * 5/9
        else: raise ValueError(f"#N/A 🚫 Unknown temperature unit '{from_u}'")
        if to_u == "c": return c
        elif to_u == "f": return c * 9/5 + 32
        elif to_u == "k": return c + 273.15
        elif to_u == "r": return (c + 273.15) * 9/5
        else: raise ValueError(f"#N/A 🚫 Unknown temperature unit '{to_u}'")
    from_unit = from_unit.strip().lower(); to_unit = to_unit.strip().lower()
    if from_unit == to_unit: return float(number)    
    if from_unit in aliases: from_unit = aliases[from_unit]
    if to_unit in aliases: to_unit = aliases[to_unit]
    if from_unit == to_unit: return float(number)
    if from_unit in temperature_units and to_unit in temperature_units: return convert_temperature(number, from_unit, to_unit)
    all_categories = [  length_units, mass_units, volume_units, time_units,
                        pressure_units, energy_units, power_units, area_units, angle_units ]
    def find_category(unit):
        for cat in all_categories:
            if unit in cat: return cat
        return None
    from_cat = find_category(from_unit); to_cat = find_category(to_unit)
    if from_cat is None or to_cat is None: raise ValueError(f"#N/A 🚫 Unknown unit '{from_unit}' or '{to_unit}'")
    if from_cat != to_cat: raise ValueError(f"#N/A 🚫 Incompatible units '{from_unit}' and '{to_unit}'")
    base_value = number * from_cat[from_unit]
    result = base_value / to_cat[to_unit]
    return result

def CORREL(array1, array2) -> float:
    """
    `=CORREL(array1, array2)` Returns the correlation coefficient between two datasets 
    Parameters:
        array1 (list | pandas.Series | numpy.ndarray): First dataset of numbers
        array2 (list | pandas.Series | numpy.ndarray): Second dataset of numbers

    *Example inputs*:

        CORREL([43, 21, 25, 42, 57, 59], [99, 65, 79, 75, 87, 81]) # ➜ 0.529809
        CORREL([1, 2, 3, None, 5], [2, 4, None, 8, 10])            # ➜ 1.0
    """
    import numpy as np, pandas as pd
    s1 = pd.Series(array1, dtype='float64')
    s2 = pd.Series(array2, dtype='float64')
    valid_mask = ~(s1.isna() | s2.isna())
    s1 = s1[valid_mask]
    s2 = s2[valid_mask]
    if len(s1) < 2: raise ValueError(" #DIV/0! 🚫 less than 2 valid pairs.",np.nan)
    return s1.corr(s2)

def COS(*args: int | float | str) -> float:
    """
    **=COS(number)** Returns the cosine of an angle given in radians.
    
    `Parameters: Only one -> Number (accepts int | float | str)`
    
    *Example inputs*:

     import excelfred as xl
     xl.COS(0)                  # -> 1.0
     xl.COS("3.1415926536")     # -> -1.0
     xl.COS(3.1415926536/2)     # -> 0.0
     xl.COS("2+1")              # -> -0.9899924966
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 COS() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 COS() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 COS expects a numeric value or a math-like string. Got `{arg}`")
            return np.cos(number)
    except AssertionError as ae: raise ValueError(str(ae))

def COSH(*args: int | float | str) -> float:
    """
    **=COSH(number)** Returns the hyperbolic cosine of a number.
    
    `Parameters: Only one -> Number (accepts int | float | str)`
    
    *Example inputs*:
    
     import excelfred as xl
     xl.COSH(0)         # -> 1.0
     xl.COSH("2")       # -> 3.7621956911
     xl.COSH(-2)        # -> 3.7621956911
     xl.COSH("3+2")     # -> 74.2099485248
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 COSH() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 COSH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 COSH expects a numeric value or a math-like string. Got `{arg}`")
            return np.cosh(number)
    except AssertionError as ae: raise ValueError(str(ae))

def COT(*args: int | float | str) -> float:
    """
    **=COT(number)** Returns the cotangent of an angle given in radians (cos(x) / sin(x)).
    
    Parameters: Only one -> Number (accepts int | float | str)
    
    *Example inputs*:
    
     COT(1)                 # -> 0.6420926159
     COT("3.1415926536/4")  # -> 1.0
     COT(3.1415926536/2)    # -> 0.0  (near zero; beware division by zero at multiples of pi)
     COT("2+1")             # -> -7.0152525514
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 COT() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 COT() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 COT expects a numeric value or a math-like string. Got `{arg}`")
            if np.sin(number) == 0: raise ZeroDivisionError("Math Error: 🚫 COT undefined for multiples of π.")
            return np.cos(number) / np.sin(number)
    except AssertionError as ae: raise ValueError(str(ae))

def COTH(*args: int | float | str) -> float:
    """
    **=COTH(number)** Returns the hyperbolic cotangent of a number (cosh(x) / sinh(x)).
    Undefined for 0.
    
    Parameters: Only one -> Number (accepts int | float | str)
    
    *Example inputs*:
    
     COTH(1)        # -> 1.3130352855
     COTH(-1)       # -> -1.3130352855
     COTH("2")      # -> 1.0373147207
     COTH("3+2")    # -> 1.00009080398
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 COTH() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 COTH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 COTH expects a numeric value or a math-like string. Got `{arg}`")
            if number == 0: raise ZeroDivisionError("Math Error: 🚫 COTH undefined for 0.")
            return np.cosh(number) / np.sinh(number)
    except AssertionError as ae: raise ValueError(str(ae))

def COUNT(*args) -> int:
    """
    `=COUNT(val_1, [val_2] ...)` Counts the number of cells in a range that contains numbers.

    *Example Inputs*:
    
     print(COUNT(1, 2, 3))                         # 3
     print(COUNT([1, 2, "apple", True, None]))     # 2
     print(COUNT("2+3", "abc", 4.5))               # 2
     print(COUNT("", None, "banana"))              # 0
     print(COUNT([np.nan, " ", "10"]))             # 1   
     print(COUNT(pd.Series([1, 2, "3"])))          # 3 
    
    *`Parameters: Multiple -> Any type`*
    """
    import pandas as pd, numpy as np
    flat_items = []
    for arg in args:
        if isinstance(arg, pd.DataFrame): flat_items.extend(arg.to_numpy().flatten())
        elif isinstance(arg, (pd.Series, pd.Index, np.ndarray, list, tuple)): flat_items.extend(np.array(arg, dtype=object).flatten())
        else: flat_items.append(arg)
    count = 0
    for item in flat_items:
        if isinstance(item, bool): count += 1
        elif isinstance(item, (int, float)): count += 1
        elif isinstance(item, str):
            s = item.strip()
            if not s: continue
            try: float(eval(s)); count += 1
            except: pass
    return count

def COUNTA(*args) -> int:
    """
    `=COUNTA(val_1, [val_2] ...)` Counts the number of cells that are not empty.
    
    *Example Inputs*:
    
     print(COUNTA(1, 2, "excelfred", None, ""))    # 3
     print(COUNTA([True, False, "", "text"]))      # 3
     print(COUNTA("2+3", "abc", 4.5))              # 3
     print(COUNTA(None, ""))                       # 0
     print(COUNTA([np.nan, ""]))                   # 1   (NaN counts since excel accpet its #N/A)
     print(COUNTA(pd.Series([None, "Hi", 7])))     # 2
    
    *`Parameters: Multiple -> Any type`*
    """
    import pandas as pd, numpy as np
    flat_items = []
    for arg in args:
        if isinstance(arg, pd.DataFrame): flat_items.extend(arg.to_numpy().flatten())
        elif isinstance(arg, (pd.Series, pd.Index, np.ndarray, list, tuple)): flat_items.extend(np.array(arg, dtype=object).flatten())
        else: flat_items.append(arg)

    count = 0
    for item in flat_items:
        if item is None: continue
        if isinstance(item, str) and item.strip() == "": continue
        count += 1
    return count

def COUNTBLANK(*args) -> int:
    """
    `=COUNTBLANK(val_1, [val_2] ...)` Counts the number of empty cells that in a specified range.
        
    *Example Inputs*:
    
     print(COUNTBLANK(1, "", None))                      # 2
     print(COUNTBLANK([None, "apple", ""]))              # 2
     print(COUNTBLANK("", "", ""))                       # 3
     print(COUNTBLANK([1, 2, 3]))                        # 0
     print(COUNTBLANK(np.nan, None, ""))                 # 3
     print(COUNTBLANK(pd.Series([np.nan, "x", None])))   # 2

    *`Parameters: Multiple -> Any type`*
    """
    import pandas as pd, numpy as np
    flat_items = []
    for arg in args:
        if isinstance(arg, pd.DataFrame): flat_items.extend(arg.to_numpy().flatten())
        elif isinstance(arg, (pd.Series, pd.Index, np.ndarray, list, tuple)): flat_items.extend(np.array(arg, dtype=object).flatten())
        else: flat_items.append(arg)
    count = 0
    for item in flat_items:
        if item is None: count += 1
        elif isinstance(item, str) and item.strip() == "": count += 1
        elif isinstance(item, float) and np.isnan(item): count += 1
    return count

def COUNTIF(range_vals, criteria) -> int:
    """
    `=COUNTIF(range, criteria)` Counts the number of cells in a range that meet the given **condition**.
    
    Parameters:
        range_vals (list | tuple | pandas.Series | numpy.ndarray): The range of cells to evaluate.
        criteria (int | float | str | bool): The condition to test against each cell.
    
    *Example Inputs*:
    
     print(COUNTIF([1, 2, 3, 4, 5], ">3"))                  # 2
     print(COUNTIF(["Apple", "Banana", "Apple"], "Apple"))  # 2
     print(COUNTIF([True, False, True], True))              # 2
     print(COUNTIF([10, 20, 30], "<=20"))                   # 2
     print(COUNTIF([None, "", 5], 5))                       # 1
     print(COUNTIF(pd.Series([1, 2, 3, 4]), "<4"))          # 3
    """
    import pandas as pd, numpy as np
    if isinstance(range_vals, pd.DataFrame): arr = range_vals.to_numpy().flatten()
    elif isinstance(range_vals, (pd.Series, pd.Index, np.ndarray, list, tuple)): arr = np.array(range_vals, dtype=object).flatten()
    else: arr = [range_vals]
    if isinstance(criteria, str) and criteria.strip() and criteria[0] in "<>=":
        try: func = lambda v: eval(f"v{criteria}")
        except: raise ValueError(f"#VALUE!: 🚫 Invalid criteria: {criteria}")
    else: func = lambda v: v == criteria
    count = 0
    for v in arr:
        if func(v): count += 1
    return count

def COUNTIFS(*criteria_pairs) -> int:
    """
    `=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2] ...)` Counts the number of cells that meet multiple criteria across one or more ranges.
    
    Parameters:
        criteria_rangeN (list | tuple | pandas.Series | numpy.ndarray): Range of cells for criteria N.
        criteriaN (int | float | str | bool): Condition to apply to criteria_rangeN.
    
    *Example Inputs*:
    
     print(COUNTIFS([1, 2, 3, 4], ">1"))                                         # 3
     print(COUNTIFS([1, 2, 3, 4], ">1", [10, 20, 30, 40], "<30"))                # 2
     print(COUNTIFS(["A", "B", "A"], "A", [True, False, True], True))            # 2
     print(COUNTIFS([5, 6, 7], ">5", [1, 2, 3], "<3"))                           # 1
     print(COUNTIFS(pd.Series([1, 2, 3]), "<3", pd.Series([10, 20, 30]), ">10")) # 1
     print(COUNTIFS([None, 1, 2], ">0", [True, False, True], True))              # 1
    """
    import pandas as pd, numpy as np
    if len(criteria_pairs) % 2 != 0: raise ValueError("#VALUE!: 🚫 COUNTIFS requires pairs of (range, criteria).")
    mask = None
    for i in range(0, len(criteria_pairs), 2):
        range_vals = criteria_pairs[i]
        if isinstance(range_vals, pd.DataFrame): arr = range_vals.to_numpy().flatten()
        elif isinstance(range_vals, (pd.Series, pd.Index, np.ndarray, list, tuple)): arr = np.array(range_vals, dtype=object).flatten()
        else: arr = [range_vals]
        criteria = criteria_pairs[i + 1]
        if isinstance(criteria, str) and criteria.strip() and criteria[0] in "<>=": func = lambda v: eval(f"v{criteria}")
        else: func = lambda v: v == criteria
        current_mask = np.array([False if v is None else func(v) for v in arr])
        if mask is None: mask = current_mask
        else: mask = mask & current_mask
    return int(mask.sum()) if mask is not None else 0

def COUPDAYSBF(settlement: str, maturity: str, frequency: int, basis: int = 0) -> int:
    """
    `=COUPDAYSBF(settlement, maturity, frequency, [basis])` Returns the **number of days from the beginning of the coupon period to the settlement date**.

    Parameters:
        settlement (str | pandas.Timestamp): Settlement date ('DD-MM-YYYY')
        maturity   (str | pandas.Timestamp): Maturity date ('DD-MM-YYYY')
        frequency  (int): 1 = annual, 2 = semi-annual, 4 = quarterly
        basis      (int, optional): Day-count basis
            0 = US (NASD) 30/360
            1 = Actual/Actual
            2 = Actual/360
            3 = Actual/365
            4 = European 30/360

    *Example Inputs*:

     print(COUPDAYSBF("15-01-2011", "15-11-2011", 1))         # 61
     print(COUPDAYSBF("15-01-2011", "15-11-2011", 2, 1))      # 61
     print(COUPDAYSBF("15-01-2011", "15-11-2011", 4, 0))      # 60  (30/360 US count)
     print(COUPDAYSBF("01-09-2022", "15-11-2025", 2))         # 47
     print(COUPDAYSBF("01-09-2022", "15-11-2025", 2, 4))      # 45  (30/360 EU count)
    """
    import pandas as pd
    try:
        sd = pd.to_datetime(settlement, dayfirst=True) if not isinstance(settlement, pd.Timestamp) else settlement
        md = pd.to_datetime(maturity,   dayfirst=True) if not isinstance(maturity,   pd.Timestamp) else maturity
    except Exception: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY' or pandas.Timestamp.")
    if sd >= md: raise ValueError("🚫 Settlement must be before maturity.")
    if frequency not in (1, 2, 4): raise ValueError("🚫 Frequency must be 1, 2, or 4.")
    if basis not in (0, 1, 2, 3, 4): raise ValueError("🚫 Basis must be an integer in {0,1,2,3,4}.")
    months = 12 // frequency; coupon = md
    while coupon > sd: coupon -= pd.DateOffset(months=months)
    start = coupon
    if basis in (1, 2, 3): days = (sd - start).days
    elif basis in (0, 4):   
        def _eom(d: pd.Timestamp) -> bool: return d.day == (d + pd.offsets.MonthEnd(0)).day
        y1, m1, d1 = start.year, start.month, start.day
        y2, m2, d2 = sd.year, sd.month, sd.day
        if basis == 0:
            if (m1 == 2 and _eom(start)): d1 = 30
            if (m2 == 2 and _eom(sd) and (m1 == 2 and _eom(start))): d2 = 30
            if d1 == 31: d1 = 30
            if d2 == 31 and d1 in (30, 31): d2 = 30
        else:
            if d1 == 31: d1 = 30
            if d2 == 31: d2 = 30
        days = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    else: days = (sd - start).days
    return int(days)

def COUPDAYS(settlement: str, maturity: str, frequency: int, basis: int = 0) -> int:
    """
    `=COUPDAYS(settlement, maturity, frequency, [basis])` Returns the **number of days in the coupon period** that contains the settlement date.

    Parameters:
        settlement (str | pandas.Timestamp): Settlement date ('DD-MM-YYYY')
        maturity   (str | pandas.Timestamp): Maturity date ('DD-MM-YYYY')
        frequency  (int): 1 = annual, 2 = semiannual, 4 = quarterly
        basis      (int, optional): Day-count basis
            0 = US (NASD) 30/360      → returns 360/frequency
            1 = Actual/Actual         → returns actual days in period
            2 = Actual/360            → returns 360/frequency
            3 = Actual/365            → returns 365/frequency
            4 = European 30/360       → returns 360/frequency

    *Example Inputs*:

     print(COUPDAYS("25-01-2011", "15-11-2011", 1))           # 304 (actual)
     print(COUPDAYS("25-01-2011", "15-11-2011", 2, 1))        # 181 (actual)
     print(COUPDAYS("25-01-2011", "15-11-2011", 2, 0))        # 180 (30/360)
     print(COUPDAYS("01-09-2022", "15-11-2025", 2))           # 76  (actual, Sep 1 → Nov 15)
     print(COUPDAYS("01-09-2022", "15-11-2025", 4, 3))        # 91  (365/4)
    """
    import pandas as pd
    try:
        sd = pd.to_datetime(settlement, dayfirst=True) if not isinstance(settlement, pd.Timestamp) else settlement
        md = pd.to_datetime(maturity,   dayfirst=True) if not isinstance(maturity,   pd.Timestamp) else maturity
    except Exception: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY' or pandas.Timestamp.")
    if sd >= md: raise ValueError("🚫 Settlement must be before maturity.")
    if frequency not in (1, 2, 4): raise ValueError("🚫 Frequency must be 1, 2, or 4.")
    if basis not in (0, 1, 2, 3, 4): raise ValueError("🚫 Basis must be an integer in {0,1,2,3,4}.")
    months = 12 // frequency; coupon = md
    while coupon > sd: coupon -= pd.DateOffset(months=months)
    start = coupon; end = coupon + pd.DateOffset(months=months)
    if basis == 1: days = (end - start).days
    elif basis in (0, 4, 2, 3):
        if basis in (0, 4, 2): days = int(360 // frequency)
        elif basis == 3: days = int(round(365 / frequency))
    else: days = (end - start).days
    return int(days)

def COUPDAYSNC(settlement: str, maturity: str, frequency: int, basis: int = 0) -> int:
    """
    `=COUPDAYSNC(settlement, maturity, frequency, [basis])` Returns the **number of days from the settlement date to the next coupon date**.

    Parameters:
        settlement (str | pandas.Timestamp): Settlement date ('DD-MM-YYYY')
        maturity   (str | pandas.Timestamp): Maturity date ('DD-MM-YYYY')
        frequency  (int): 1 = annual, 2 = semiannual, 4 = quarterly
        basis      (int, optional): Day-count basis
            0 = US (NASD) 30/360
            1 = Actual/Actual
            2 = Actual/360
            3 = Actual/365
            4 = European 30/360

    *Example Inputs*:

     print(COUPDAYSNC("15-01-2011", "15-11-2011", 1))         # 303
     print(COUPDAYSNC("15-01-2011", "15-11-2011", 2, 1))      # 120
     print(COUPDAYSNC("15-01-2011", "15-11-2011", 4, 0))      # 90   
     print(COUPDAYSNC("01-09-2022", "15-11-2025", 2))         # 75
     print(COUPDAYSNC("01-09-2022", "15-11-2025", 2, 4))      # 75   
    """
    import pandas as pd
    try:
        sd = pd.to_datetime(settlement, dayfirst=True) if not isinstance(settlement, pd.Timestamp) else settlement
        md = pd.to_datetime(maturity,   dayfirst=True) if not isinstance(maturity,   pd.Timestamp) else maturity
    except Exception: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY' or pandas.Timestamp.")
    if sd >= md: raise ValueError("🚫 Settlement must be before maturity.")
    if frequency not in (1, 2, 4): raise ValueError("🚫 Frequency must be 1, 2, or 4.")
    if basis not in (0, 1, 2, 3, 4): raise ValueError("🚫 Basis must be an integer in {0,1,2,3,4}.")
    months = 12 // frequency; coupon = md
    while coupon > sd: coupon -= pd.DateOffset(months=months)
    next_coupon = coupon + pd.DateOffset(months=months)
    if basis in (1, 2, 3): days = (next_coupon - sd).days
    elif basis in (0, 4):
        def _eom(d: pd.Timestamp) -> bool: return d.day == (d + pd.offsets.MonthEnd(0)).day
        y1, m1, d1 = sd.year, sd.month, sd.day
        y2, m2, d2 = next_coupon.year, next_coupon.month, next_coupon.day
        if basis == 0:
            if (m1 == 2 and _eom(sd)): d1 = 30
            if (m2 == 2 and _eom(next_coupon) and (m1 == 2 and _eom(sd))): d2 = 30
            if d1 == 31: d1 = 30
            if d2 == 31 and d1 in (30, 31): d2 = 30
        else:
            if d1 == 31: d1 = 30
            if d2 == 31: d2 = 30
        days = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    else: days = (next_coupon - sd).days
    return int(days)

def COUPNCD(settlement: str, maturity: str, frequency: int, basis: int = 0) -> str:
    """
    `=COUPNCD(settlement, maturity, frequency, [basis])` Returns the **next coupon date after the settlement date** in `'DD-MM-YYYY'` format.

    Parameters:
        settlement (str | pandas.Timestamp): Settlement date ('DD-MM-YYYY')
        maturity   (str | pandas.Timestamp): Maturity date ('DD-MM-YYYY')
        frequency  (int): 1 = annual, 2 = semiannual, 4 = quarterly
        basis      (int, optional): Day-count basis (not used for the date itself; included for Excel signature parity)

    *Example Inputs*:

     print(COUPNCD("15-01-2011", "15-11-2011", 1))            # '15-11-2011'
     print(COUPNCD("15-01-2011", "15-11-2011", 2, 1))         # '15-05-2011'
     print(COUPNCD("15-01-2011", "15-11-2011", 4, 0))         # '15-02-2011'
     print(COUPNCD("01-09-2022", "15-11-2025", 2))            # '15-11-2022'
     print(COUPNCD("14-11-2025", "15-11-2025", 2))            # '15-11-2025'
    """
    import pandas as pd
    try:
        sd = pd.to_datetime(settlement, dayfirst=True) if not isinstance(settlement, pd.Timestamp) else settlement
        md = pd.to_datetime(maturity,   dayfirst=True) if not isinstance(maturity,   pd.Timestamp) else maturity
    except Exception: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY' or pandas.Timestamp.")
    if sd >= md: raise ValueError("🚫 Settlement must be before maturity.")
    if frequency not in (1, 2, 4): raise ValueError("🚫 Frequency must be 1, 2, or 4.")
    if basis not in (0, 1, 2, 3, 4): raise ValueError("🚫 Basis must be an integer in {0,1,2,3,4}.")
    months = 12 // frequency; coupon = md
    while coupon > sd: coupon -= pd.DateOffset(months=months)
    next_coupon = coupon + pd.DateOffset(months=months)
    return next_coupon.strftime("%d-%m-%Y")

def COUPNUM(settlement: str, maturity: str, frequency: int, basis: int = 0) -> int:
    """
    `=COUPNUM(settlement, maturity, frequency, [basis])` Returns the **number of coupons payable between the settlement and maturity**.

    Parameters:
        settlement (str | pandas.Timestamp): Settlement date ('DD-MM-YYYY')
        maturity   (str | pandas.Timestamp): Maturity date ('DD-MM-YYYY')
        frequency  (int): 1 = annual, 2 = semiannual, 4 = quarterly
        basis      (int, optional): Day-count basis (not used in the count; included for Excel signature parity)

    *Example Inputs*:

     print(COUPNUM("15-01-2011", "15-11-2011", 1))            # 1
     print(COUPNUM("15-01-2011", "15-11-2011", 2, 1))         # 2
     print(COUPNUM("15-01-2011", "15-11-2011", 4, 0))         # 4
     print(COUPNUM("01-09-2022", "15-11-2025", 2))            # 7
     print(COUPNUM("15-11-2024", "15-11-2025", 2))            # 2
    """
    import pandas as pd
    try:
        sd = pd.to_datetime(settlement, dayfirst=True) if not isinstance(settlement, pd.Timestamp) else settlement
        md = pd.to_datetime(maturity,   dayfirst=True) if not isinstance(maturity,   pd.Timestamp) else maturity
    except Exception: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY' or pandas.Timestamp.")
    if sd >= md: raise ValueError("🚫 Settlement must be before maturity.")
    if frequency not in (1, 2, 4): raise ValueError("🚫 Frequency must be 1, 2, or 4.")
    if basis not in (0, 1, 2, 3, 4): raise ValueError("🚫 Basis must be an integer in {0,1,2,3,4}.")
    months = 12 // frequency; coupon = md; count = 0
    while coupon > sd:
        coupon -= pd.DateOffset(months=months)
        count += 1
    return int(count)

def COUPPCD(settlement: str, maturity: str, frequency: int, basis: int = 0) -> str:
    """
    `=COUPPCD(settlement, maturity, frequency, [basis])` Returns the **previous coupon date before the settlement date** in `'DD-MM-YYYY'` format.

    Parameters:
        settlement (str | pandas.Timestamp): Settlement date ('DD-MM-YYYY')
        maturity   (str | pandas.Timestamp): Maturity date ('DD-MM-YYYY')
        frequency  (int): 1 = annual, 2 = semiannual, 4 = quarterly
        basis      (int, optional): Day-count basis (not used for the date itself; included for Excel signature parity)

    *Example Inputs*:

     print(COUPPCD("15-01-2011", "15-11-2011", 1))            # '15-11-2010'
     print(COUPPCD("15-01-2011", "15-11-2011", 2, 1))         # '15-11-2010'
     print(COUPPCD("15-01-2011", "15-11-2011", 4, 4))         # '15-12-2010'
     print(COUPPCD("01-09-2022", "15-11-2025", 2))            # '15-05-2022'
     print(COUPPCD("14-11-2025", "15-11-2025", 2))            # '15-05-2025'
    """
    import pandas as pd
    try:
        sd = pd.to_datetime(settlement, dayfirst=True) if not isinstance(settlement, pd.Timestamp) else settlement
        md = pd.to_datetime(maturity,   dayfirst=True) if not isinstance(maturity,   pd.Timestamp) else maturity
    except Exception: raise ValueError("🚫 Invalid date format. Use 'DD-MM-YYYY' or pandas.Timestamp.")
    if sd >= md: raise ValueError("🚫 Settlement must be before maturity.")
    if frequency not in (1, 2, 4): raise ValueError("🚫 Frequency must be 1, 2, or 4.")
    if basis not in (0, 1, 2, 3, 4): raise ValueError("🚫 Basis must be an integer in {0,1,2,3,4}.")
    months = 12 // frequency; coupon = md
    while coupon > sd: coupon -= pd.DateOffset(months=months)
    return coupon.strftime("%d-%m-%Y")

def COVARIANCE_P(array1, array2) -> float:
    """
    `=COVARIANCE.P(array1, array2)` Returns the **population covariance** of two datasets.

    Parameters:
        array1 (list | pandas.Series | numpy.ndarray): First dataset (numbers, numeric strings; text/None ignored)
        array2 (list | pandas.Series | numpy.ndarray): Second dataset (numbers, numeric strings; text/None ignored)

    Notes:
        • Pairs are formed by position; pairs with a non-numeric in either array are ignored.  
        • Requires at least 1 valid numeric pair; otherwise → `#DIV/0!`.  
        • If array lengths differ, extra elements are ignored (Excel requires same-size ranges, but this function
          follows the CORREL-style behavior you provided: only aligned/valid pairs are used).

    *Example inputs*:

        from excelfred import COVARIANCE_S; import pandas as pd, numpy as np
        print(COVARIANCE_P([2, 4, 6], [1, 3, 5]))                            # ➜ 2.6666666667
        print(COVARIANCE_P([1, 2, 3, None, 5], [2, 4, None, 8, 10]))         # ➜ 5.7777777778
        print(COVARIANCE_P(["1", "x", "3"], [2, 5, "7"]))                    # ➜ 3.0
        print(COVARIANCE_P(pd.Series([1, np.nan, 3]), pd.Series([4, 5, 6]))) # ➜ 0.5
    """
    import numpy as np, pandas as pd
    try:
        s1 = pd.Series(array1, dtype="object")
        s2 = pd.Series(array2, dtype="object")
    except Exception: raise TypeError("Type Error: 🚫 array1/array2 must be array-like (list/Series/ndarray).")
    s1 = s1.map(lambda v: np.nan if isinstance(v, (bool, np.bool_)) else v)
    s2 = s2.map(lambda v: np.nan if isinstance(v, (bool, np.bool_)) else v)
    s1 = pd.to_numeric(s1, errors="coerce")
    s2 = pd.to_numeric(s2, errors="coerce")
    valid = ~(s1.isna() | s2.isna())
    x = s1[valid].to_numpy(dtype=float); y = s2[valid].to_numpy(dtype=float)
    n = x.size
    if n < 1: raise ValueError("#DIV/0! 🚫 No valid numeric pairs to compute COVARIANCE.P.")
    mx = float(np.mean(x)); my = float(np.mean(y))
    cov = float(np.sum((x - mx) * (y - my)) / n)
    return cov

def COVARIANCE_S(array1, array2) -> float:
    """
    `=COVARIANCE.S(array1, array2)` Returns the **sample covariance** of two datasets.

    Parameters:
        array1 (list | pandas.Series | numpy.ndarray): First dataset (numbers, numeric strings; text/None ignored)
        array2 (list | pandas.Series | numpy.ndarray): Second dataset (numbers, numeric strings; text/None ignored)

    Notes:
        • Pairs are formed by position; pairs with a non-numeric in either array are ignored.  
        • Uses the sample formula (divides by N−1); requires at least 2 valid numeric pairs; otherwise → `#DIV/0!`.  
        • If array lengths differ, extra elements are ignored (same alignment behavior as above).

    *Example inputs*:

        from excelfred import COVARIANCE_S; import pandas as pd, numpy as np
        print(COVARIANCE_S([2, 4, 6], [1, 3, 5]))                                   # ➜ 4.0
        print(COVARIANCE_S([1, 2, 3, None, 5], [2, 4, None, 8, 10]))                # ➜ 8.6666666667
        print(COVARIANCE_S(["1", "x", "3"], [2, 5, "7"]))                           # ➜ 4.5
        print(COVARIANCE_S(np.array([1, 2, np.nan, 4]), pd.Series([2, 1, 3, 0])))   # ➜ -1.0
    """
    import numpy as np, pandas as pd
    try:
        s1 = pd.Series(array1, dtype="object")
        s2 = pd.Series(array2, dtype="object")
    except Exception: raise TypeError("Type Error: 🚫 array1/array2 must be array-like (list/Series/ndarray).")
    s1 = s1.map(lambda v: np.nan if isinstance(v, (bool, np.bool_)) else v)
    s2 = s2.map(lambda v: np.nan if isinstance(v, (bool, np.bool_)) else v)
    s1 = pd.to_numeric(s1, errors="coerce")
    s2 = pd.to_numeric(s2, errors="coerce")
    valid = ~(s1.isna() | s2.isna())
    x = s1[valid].to_numpy(dtype=float); y = s2[valid].to_numpy(dtype=float)
    n = x.size
    if n < 2: raise ValueError("#DIV/0! 🚫 Need at least 2 valid numeric pairs for COVARIANCE.S.")
    mx = float(np.mean(x)); my = float(np.mean(y))
    cov = float(np.sum((x - mx) * (y - my)) / (n - 1))
    return cov

def CSC(*args: int | float | str) -> float:
    """
    **=CSC(number)** Returns the cosecant of an angle given in radians.

    Formula: `CSC(x) = 1 / SIN(x)`

    `Parameters: Only one -> Number (accepts int | float | str)`

    *Example inputs*:

     import excelfred as xl
     xl.CSC(1)                  # -> 1.1883951058
     xl.CSC("3.1415926536/2")   # -> 1.0
     xl.CSC("2+1")              # -> 7.0861673957
     xl.CSC(0.5235987756)       # -> 2.0
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 CSC() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 CSC() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 CSC expects a numeric value or a math-like string. Got `{arg}`")
            sin_val = np.sin(number)
            if sin_val == 0: raise ZeroDivisionError("Math Error: 🚫 CSC undefined for angles where SIN(x) = 0.")
            return 1 / sin_val
    except AssertionError as ae: raise ValueError(str(ae))

def CSCH(*args: int | float | str) -> float:
    """
    **=CSCH(number)** Returns the hyperbolic cosecant of a number.

    Formula: `CSCH(x) = 1 / SINH(x)`

    `Parameters: Only one -> Number (accepts int | float | str)`

    *Example inputs*:

     import excelfred as xl
     xl.CSCH(1)        # -> 0.8509181282
     xl.CSCH("-2")     # -> -0.2757205648
     xl.CSCH("3+2")    # -> 0.0134752822
     xl.CSCH(0.5)      # -> 1.9190347513
    """
    import numpy as np
    try:
        assert args, "Value Error: 🚫 CSCH() requires one numeric input."
        if len(args) != 1: raise ValueError("Parameters Error: 🚫 CSCH() only takes one input.")
        for arg in args:
            if isinstance(arg, str) and not arg.replace('.', '', 1).replace('-', '', 1).isdigit():
                try: arg = eval(arg)
                except: raise ValueError(f"Value Error: 🚫 Cannot evaluate the expression `{arg}`.")
            try: number = float(arg)
            except: raise TypeError(f"Type Error: 🚫 CSCH expects a numeric value or a math-like string. Got `{arg}`")
            sinh_val = np.sinh(number)
            if sinh_val == 0: raise ZeroDivisionError("Math Error: 🚫 CSCH undefined for x = 0.")
            return 1 / sinh_val
    except AssertionError as ae: raise ValueError(str(ae))

class Cube:
    """
    `A`**`OLAP`**`engine (or cube) for in-memory analytics.`
     
    **Library module**:

      from excelfred import CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, 
      CUBESET, CUBESETCOUNT, CUBERANKEDMEMBER

    Attributes:
        name (str): Name of the cube.
        dimensions (dict): Dimensions of the cube.
        measures_meta (dict): Measures and their metadata.
        kpis (dict): Key Performance Indicators.
        data (DataFrame): Underlying fact table.

    **Example Inputs**:

        # Create Table
        import pandas; from excelfred import cube # import numpy or scipy if required !
        df = pd.DataFrame({ "Product": ["Bike","Bike","Helmet","Helmet","Bike","Helmet"],
                            "Region":  ["US","US","US","EU","EU","EU"],
                            "Year":    [2023, 2024, 2023, 2023, 2023, 2024],
                            "Sales":   [1000, 1100, 300,   200,  700,  500],
                            "Qty":     [10,   12,   4,     3,    8,    5],
                            "InventoryEnd":[50,55,20,18,40,35] })

        # Create Database
        cube = Cube("SalesCube") 
        cube.add_data(df) 

        # Insert Dimensions
        cube.add_dimension("Product",
          { "Bike": {"caption":"Bike","unique_name":"[Product].[Bike]","unary":1},
            "Helmet": {"caption":"Helmet","unique_name":"[Product].[Helmet]","unary":1} })
        cube.add_dimension("Region", 
          { "US": {"caption":"United States","unique_name":"[Region].[US]","unary":1},
            "EU": {"caption":"Europe","unique_name":"[Region].[EU]","unary":1} })
        cube.add_dimension("Year",
          {  2023: {"caption":"2023","unique_name":"[Date].[Year].[2023]","unary":1},
             2024: {"caption":"2024","unique_name":"[Date].[Year].[2024]","unary":1} })
        cube.add_dimension("Date",
          { "2023": {},
            "2024": {} })

        # Measures
        cube.add_measure("Sales", "Sales", agg="sum")
        cube.add_measure("Qty", "Qty", agg="sum")
        cube.add_measure("InventoryEnd", "InventoryEnd", agg="last_non_empty", time_dim="Year")

        # KPI
        cube.add_kpi("Revenue KPI", {
                    "value": "Sales",
                    "goal":  5000,
                    "status": lambda c, ctx: c.evaluate_measure_vectorized("Sales", ctx) / 5000 })

    `Now use CUBE series to display output by attached snippet in description of each functions`
    """
    def __init__(self, name): self.name = name; self.dimensions = {}; self.measures_meta = {}; self.kpis = {}; self.data = None
    def add_data(self, df): self.data = df
    def add_dimension(self, name, members): self.dimensions[name] = members
    def add_measure(self, name, column, agg="sum", time_dim=None): self.measures_meta[name] = { "column": column, "agg": agg, "time_dim": time_dim }
    def add_kpi(self, name, parts): self.kpis[name] = parts
    def evaluate_measure_vectorized(self, measure, ctx):
        meta = self.measures_meta[measure]
        df = self.data
        mask = None
        for dim, key in ctx.items():
            m = (df[dim] == key)
            mask = m if mask is None else (mask & m)
        sub = df if mask is None else df[mask]
        if sub.empty: return 0.0
        return float(sub[meta["column"]].sum())

def CUBEKPIMEMBER(cube: object, kpi_name: str, kpi_property: int | str, caption: str | None = None) -> dict:
    """
    `=CUBEKPIMEMBER(connection, kpi_name, kpi_property, [caption])` Returns a **Key Performance Indicator (KPI)** property from the cube and display the KPI name in the cell.

    Parameters:
        cube (object): The cube connection object. Must not be None.
        kpi_name (str): The name of the KPI in the cube.
        kpi_property (int | str): The property of the KPI to return.
            Accepts 1 = value, 2 = goal, 3 = status, 4 = trend, 5 = weight,
            or the property name as a string.
        [caption] (str, optional): A caption for the cell. If omitted, the default caption is used.
        
    **SAMPLE CODE**:

     kpi_value_handle = CUBEKPIMEMBER(cube, "Revenue KPI","value)
     print(kpi_value_handle) 

     # OUTPUT:
     # >>> {'type': 'kpi', 'kpi_name': 'Revenue KPI', 'kpi_part': 'value', 'caption': 'Revenue KPI - value'}

     kpi_goal_handle = CUBEKPIMEMBER(cube, "Revenue KPI", "goal", "My KPI Goal")
     print(kpi_goal_handle)
     # OUTPUT:
     # >>> {'type': 'kpi', 'kpi_name': 'Revenue KPI', 'kpi_part': 'goal', 'caption': 'My KPI Goal'}
    
    `Refer "Cube" Class in excelfred to understand database.`
    """
    if cube is None: raise ValueError("#NAME? 🚫 Invalid or missing connection.")
    if not hasattr(cube, "kpis") or kpi_name not in cube.kpis: raise ValueError(f"#N/A 🚫 KPI not found: {kpi_name}")
    prop_map = {1:"value", 2:"goal", 3:"status", 4:"trend", 5:"weight"}
    if isinstance(kpi_property, int): prop = prop_map.get(kpi_property)
    elif isinstance(kpi_property, str): prop = kpi_property.strip().lower()
    else: raise ValueError("#VALUE! 🚫 invalid KPI property type")
    if prop is None: raise ValueError("#VALUE! 🚫 invalid KPI property")
    kpi_parts = cube.kpis[kpi_name]
    part = kpi_parts.get(prop)
    if prop in ("status", "trend") and callable(part):
        def wrapper(c, filters):
            try:  val = part(c, filters)
            except TypeError:  val = part(filters)
            goal_val = kpi_parts.get("goal")
            if isinstance(goal_val, str):  goal_val = c.evaluate_measure_vectorized(goal_val, filters)
            elif callable(goal_val):
                try: goal_val = goal_val(c, filters)
                except TypeError: goal_val = goal_val(filters)
            if goal_val is None or goal_val == 0: raise ValueError("#DIV/0! 🚫 goal is zero or missing")
            return float(val) / float(goal_val)
        part = wrapper
    return { "type": "kpi", "kpi_name": kpi_name, "kpi_part": prop, "caption": caption or f"{kpi_name} - {prop}" }

def CUBEMEMBER(cube: object, member_expression: str, caption: str | None = None) -> dict:
    """
    `=CUBEMEMBER(connection, member_expression, [caption])` Returns a **member** or tuple from a cube.

    Parameters:
        cube (object): The cube connection object. Must not be None.
        member_expression (str): The unique name of a member in the cube.
        [caption] (str, optional): A caption for the cell. If omitted, the member’s default caption is used.
    
    **SAMPLE CODE**:

     print(CUBEMEMBER(cube, "[Product].[Bike]")) 
     # OUTPUT >>> {'type': 'member', 'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'Bike'}
    
     print(CUBEMEMBER(cube, "[Product].[Bike]", "My Bike"))
     # OUTPUT >>> {'type': 'member', 'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'My Bike'}

    `Refer "Cube" Class in excelfred to understand database.`
    """
    if cube is None: raise ValueError("#NAME? 🚫 Invalid or missing connection.")
    if not isinstance(member_expression, str): raise ValueError("#VALUE! 🚫 member_expression must be a string")
    s = member_expression.strip()
    try:
        token = s
        if token.startswith("[") and token.endswith("]"): token = token[1:-1]
        parts = token.split("].[")
        dim = parts[0]; mem = parts[-1]
    except Exception: raise ValueError(f"#N/A 🚫 Invalid member expression: {member_expression}")
    if not hasattr(cube, "dimensions") or dim not in cube.dimensions: raise ValueError(f"#N/A 🚫 Dimension not found: {dim}")
    members = cube.dimensions.get(dim, {})
    if mem in members:
        info = members[mem]
        return {"type": "member", "unique_name": member_expression, "dimension": dim, "key": mem, "caption": caption or info.get("caption", mem)}
    for key, info in members.items():
        if str(info.get("caption", "")).lower() == mem.lower(): return {"type": "member", "unique_name": member_expression, "dimension": dim, "key": key, "caption": caption or info.get("caption", key)}
    for key, info in members.items():
        if info.get("unique_name") == member_expression: return {"type": "member", "unique_name": member_expression, "dimension": dim, "key": key, "caption": caption or info.get("caption", key)}
    raise ValueError(f"#N/A 🚫 Member not found: {member_expression}")

def CUBEMEMBERPROPERTY(cube: object, member_expression: str | dict, property_name: str) -> str:
    """
    `=CUBEMEMBERPROPERTY(connection, member, property)` Returns the **property value** of a cube member.

    Parameters:
        cube (object): The cube connection object. Must not be None.
        member_expression (str | dict): The unique name of a member or a member handle.
        property_name (str): The name of the property to return.
            Common values: "MEMBER_CAPTION", "MEMBER_UNIQUE_NAME", "KEY".

    **SAMPLE CODE**:

     print(CUBEMEMBERPROPERTY(cube, "[Product].[Bike]", "caption"))
     # OUTPUT >>> Bike
    
    `Refer "Cube" Class in excelfred to understand database.`
    """
    if cube is None: raise ValueError("#NAME? 🚫 Invalid or missing connection.")
    if not property_name: raise ValueError("#VALUE! 🚫 property_name is required")
    if isinstance(member_expression, dict) and member_expression.get("type") == "member":
        handle = member_expression; prop = property_name.strip().lower()
        if prop in ("member_caption", "caption"): return handle.get("caption")
        if prop in ("member_unique_name", "unique_name"): return handle.get("unique_name")
        if prop in ("key",): return handle.get("key")
        raise ValueError(f"#N/A 🚫 Property not found: {property_name}")
    if isinstance(member_expression, str):
        s = member_expression.strip()
        try:
            token = s
            if token.startswith("[") and token.endswith("]"): token = token[1:-1]
            parts = token.split("].[")
            dim = parts[0]; mem = parts[-1]
        except Exception: raise ValueError(f"#N/A 🚫 Invalid member expression: {member_expression}")
        if not hasattr(cube, "dimensions") or dim not in cube.dimensions: raise ValueError(f"#N/A 🚫 Dimension not found: {dim}")
        members = cube.dimensions.get(dim, {}); key = None; info = None
        if mem in members: key = mem; info = members[mem]
        else:
            for k, v in members.items(): 
                if str(v.get("caption", "")).lower() == mem.lower(): key = k; info = v; break
            if not key:
                for k, v in members.items():
                    if v.get("unique_name") == member_expression: key = k; info = v; break
        if not key: raise ValueError(f"#N/A 🚫 Member not found: {member_expression}")
        prop = property_name.strip().lower()
        if prop in ("member_caption", "caption"): return info.get("caption", key)
        if prop in ("member_unique_name", "unique_name"): return info.get("unique_name", f"[{dim}].[{key}]")
        if prop in ("key",): return key
        raise ValueError(f"#N/A 🚫 Property not found: {property_name}")
    raise ValueError("#VALUE! 🚫 invalid member handle or expression")

def CUBESET(cube: object, set_expression: str | list[str | dict] | tuple, caption: str | None = None, sort_order: int | None = None, sort_by: str | None = None) -> dict:  
    """
    `=CUBESET(connection, set_expression, [caption], [sort_order], [sort_by])` Returns a **set** of members from a cube.

    Parameters:
        cube (object): The cube connection object. Must not be None.
        set_expression (str | list | tuple): A set expression, a list/tuple of members, or a dimension name.
        [caption] (str, optional): A caption for the cell. If omitted, the set_expression is used.
        [sort_order] (int, optional): Sort order for the set. 1 = descending, 2 = ascending, 5 or 6 = sort by measure (requires sort_by).
        [sort_by] (str, optional): The measure name to sort by (required if sort_order is 5 or 6).
    
    **SAMPLE CODE**:

     product_set = CUBESET(cube, "Product")
     print(product_set)  #refer "Cube Class" in excelfred to understand database
     # >>> {'type': 'set', 'members': [{'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'Bike'}, {'unique_name': '[Product].[Helmet]', 'dimension': 'Product', 'key': 'Helmet', 'caption': 'Helmet'}], 'caption': 'Product'}
    
     sorted_product_set = CUBESET(cube, "Product", "Products Sorted", sort_order=2)
     print(sorted_product_set)
     # >>> {'type': 'set', 'members': [{'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'Bike'}, {'unique_name': '[Product].[Helmet]', 'dimension': 'Product', 'key': 'Helmet', 'caption': 'Helmet'}], 'caption': 'Products Sorted'}
     #  {'type': 'set', 'members': [{'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'Bike'}, {'unique_name': '[Product].[Helmet]', 'dimension': 'Product', 'key': 'Helmet', 'caption': 'Helmet'}], 'caption': 'Products Sorted'}

     sorted_by_sales_asc = CUBESET(cube,"Product","Products Sorted by Sales Asc", 6, "Sales")
     print(sorted_by_sales_asc)
     # >>> {'type': 'set', 'members': [
     #     {'unique_name': '[Product].[Helmet]', 'dimension': 'Product', 'key': 'Helmet', 'caption': 'Helmet'},
     #     {'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'Bike'} ], 
     #     'caption': 'Products Sorted by Sales Asc'}

    `Refer "Cube" Class in excelfred to understand database.`
    """    
    if cube is None: raise ValueError("#NAME? 🚫 Invalid or missing connection.")
    members_list = []
    if isinstance(set_expression, (list, tuple)):
        for m in set_expression:
            if isinstance(m, dict) and m.get("type") == "member": members_list.append({"unique_name": m.get("unique_name"), "dimension": m.get("dimension"), "key": m.get("key"), "caption": m.get("caption")}); continue
            if not isinstance(m, str): raise ValueError("#VALUE! 🚫 set member must be string or member-handle")
            s = m.strip()
            try:
                token = s
                if token.startswith("[") and token.endswith("]"): token = token[1:-1]
                parts = token.split("].[")
                dim = parts[0]; mem = parts[-1]
            except Exception: raise ValueError(f"#N/A 🚫 Invalid member in set: {m}")
            if dim not in cube.dimensions: raise ValueError(f"#N/A 🚫 Dimension not found: {dim}")
            members = cube.dimensions.get(dim, {})
            if mem in members:
                info = members[mem]
                members_list.append({"unique_name": m, "dimension": dim, "key": mem, "caption": info.get("caption", mem)})
            else:
                found = False
                for key, info in members.items():
                    if str(info.get("caption","")).lower() == mem.lower() or info.get("unique_name")==m:
                        members_list.append({"unique_name": m, "dimension": dim, "key": key, "caption": info.get("caption", key)})
                        found = True; break
                if not found: raise ValueError(f"#N/A 🚫 Member not found in set: {m}")
    elif isinstance(set_expression, str):
        s = set_expression.strip()
        if s in cube.dimensions:
            members_map = cube.dimensions.get(s, {})
            for key, info in members_map.items(): members_list.append({"unique_name": info.get("unique_name", f"[{s}].[{key}]"), "dimension": s, "key": key, "caption": info.get("caption", key)})
        else:
            try:
                token = s
                if token.startswith("[") and token.endswith("]"): token = token[1:-1]
                dim = token.split("].[")[0]
            except Exception: dim = None
            if dim and dim in cube.dimensions:
                members_map = cube.dimensions.get(dim, {})
                for key, info in members_map.items(): members_list.append({"unique_name": info.get("unique_name", f"[{dim}].[{key}]"), "dimension": dim, "key": key, "caption": info.get("caption", key)})
            else: raise ValueError(f"#N/A 🚫 Invalid set expression: {set_expression}")
    else: raise ValueError("#VALUE! 🚫 set_expression must be string or list/tuple")
    if len(members_list) == 0: raise ValueError("#N/A 🚫 The set is empty.")
    if sort_order in (1, 2): members_list = sorted(members_list, key=lambda x: str(x.get("caption","")), reverse=(sort_order==1))
    elif sort_order in (5, 6):
        if not sort_by: raise ValueError("#VALUE! 🚫 sort_by measure is required for this sort_order")
        scores = []
        for m in members_list:
            try:
                meta = cube.measures_meta.get(sort_by)
                if meta is None: score = 0.0
                else:
                    col = meta.get("column"); agg = meta.get("agg","sum"); time_dim = meta.get("time_dim")
                    df = cube.data if hasattr(cube, "data") else None
                    if df is None: score = 0.0
                    else:
                        if m["dimension"] not in df.columns: score = 0.0
                        else:
                            sub = df[df[m["dimension"]] == m["key"]]
                            if sub.shape[0] == 0: score = 0.0
                            else:
                                series_vals = sub[col].astype(float)
                                if agg == "sum": score = float(series_vals.sum())
                                elif agg == "avg":
                                    vals = series_vals.dropna()
                                    if vals.shape[0] == 0: score = 0.0
                                    else: score = float(vals.mean())
                                elif agg == "last_non_empty":
                                    if time_dim and time_dim in sub.columns:
                                        orded = sub.loc[series_vals.notna()].sort_values(by=time_dim)
                                        score = float(orded.iloc[-1][col]) if orded.shape[0]>0 else 0.0
                                    else: score = float(series_vals.sum())
                                else: score = float(series_vals.sum())
                scores.append(score)
            except Exception: scores.append(0.0)
        zipped = list(zip(members_list, scores))
        zipped.sort(key=lambda x: x[1], reverse=(sort_order==6))
        members_list = [z[0] for z in zipped]
    else: pass
    return {"type": "set", "members": members_list, "caption": caption or (set_expression if isinstance(set_expression, str) else "set")}

def CUBESETCOUNT(set_handle: dict | list | tuple) -> int:
    """
    `=CUBESETCOUNT(set)` Returns the **number of items** in a set.

    **SAMPLE CODE**:

     print(CUBESETCOUNT(CUBESET(cube, "Product"))) #refer "Cube Class" in excelfred to understand database
     # OUTPUT >>> 2

    `Parameters: set_handle (dict | list | tuple) A set handle returned by CUBESET or a list/tuple of members.`
    """
    if set_handle is None: raise ValueError("#N/A 🚫 Set is empty or not available.")
    if isinstance(set_handle, dict) and set_handle.get("type") == "set":
        members = set_handle.get("members", [])
        return int(len(members))
    if isinstance(set_handle, (list, tuple)): return int(len(set_handle))
    raise ValueError("#VALUE! 🚫 Argument must be a set-handle or sequence")

def CUBERANKEDMEMBER(cube: object, set_handle: dict, rank: int, caption: str | None = None) -> dict:
    """
    `=CUBERANKEDMEMBER(connection, set, rank, [caption])` Returns the **nth-ranked member** of a set from a cube.

    Parameters:
        cube (object): The cube connection object. Must not be None.
        set_handle (dict): A set handle returned by CUBESET.
        rank (int): The position of the member to return (1-based index).
        [caption] (str, optional): A caption for the cell. If omitted, the member’s default caption is used.
    
    **SAMPLE CODE**:

     rank1 = CUBERANKEDMEMBER(cube, product_set, 1)
     print(rank1) 
     # OUTPUT >>> {'type': 'member', 'unique_name': '[Product].[Bike]', 'dimension': 'Product', 'key': 'Bike', 'caption': 'Bike'}
     #            {'type': 'kpi', 'kpi_name': 'Revenue KPI', 'kpi_part': 'value', 'caption': 'Revenue KPI - value'}
    
    `Refer "Cube Class" in excelfred to understand database`
    """
    if cube is None: raise ValueError("#NAME? 🚫 Invalid or missing connection.")
    if not isinstance(set_handle, dict) or set_handle.get("type") != "set": raise ValueError("#VALUE! 🚫 set_handle must be a set-handle")
    if not isinstance(rank, int) or rank < 1: raise ValueError("#VALUE! 🚫 rank must be an integer >= 1")
    members = set_handle.get("members", [])
    if rank > len(members): raise ValueError("#N/A 🚫 rank exceeds set size.")
    m = members[rank-1]
    return {"type":"member","unique_name": m.get("unique_name"), "dimension": m.get("dimension"), "key": m.get("key"), "caption": caption or m.get("caption")}

def CUBEVALUE(cube: object, *member_expressions: object) -> float:
    """
    `=CUBEVALUE(connection, member_expression1, [member_expression2], ...)` Returns the **aggregated value** from the cube for the given members, set, or KPI.

    Parameters:
        cube (object): The cube connection object. Must not be None.
        member_expressions (object): One or more member expressions, sets, KPI handles, or measure names. Each argument defines a filter context for the value.
    
    **SAMPLE OUTPUT**:

     print(CUBEVALUE(cube, rank1, "Sales"))             # 2800.0
     print(CUBEVALUE(cube, rank1, kpi_value_handle))    # 2800.0
     print(CUBEVALUE(cube, product_set, "Sales"))       # 4300.0

    `Refer "Cube Class" in excelfred to understand database`
    """
    import pandas as pd
    if cube is None: raise ValueError("#NAME? 🚫 Invalid or missing connection.")
    if not hasattr(cube, "measures_meta") or not hasattr(cube, "data"): raise ValueError("#N/A 🚫 cube missing measures or data")
    def _resolve_member_inline(expr):
        if not isinstance(expr, str): raise ValueError("#VALUE! 🚫 member expression must be string")
        s = expr.strip()
        try:
            t = s
            if t.startswith("[") and t.endswith("]"): t = t[1:-1]
            parts = t.split("].[")
            dim = parts[0]; mem = parts[-1]
        except Exception: raise ValueError(f"#N/A 🚫 Invalid member expression: {expr}")
        if dim not in cube.dimensions: raise ValueError(f"#N/A 🚫 Dimension not found: {dim}")
        members = cube.dimensions.get(dim, {})
        if mem in members: return dim, mem, members[mem]
        for key, info in members.items():
            if str(info.get("caption","")).lower() == mem.lower() or info.get("unique_name") == expr: return dim, key, info
        raise ValueError(f"#N/A 🚫 Member not found: {expr}")
    measure_name = None; kpi_handle = None
    contexts = [{}]  
    for arg in member_expressions:
        if arg is None: continue
        if isinstance(arg, dict):
            t = arg.get("type")
            if t == "member":
                dim = arg.get("dimension"); key = arg.get("key")
                if dim is None or key is None:
                    try: dim, key, _ = _resolve_member_inline(arg.get("unique_name"))
                    except Exception as e: raise ValueError(str(e))
                for c in contexts: c[dim] = key
                continue
            if t == "set":
                members = arg.get("members", [])
                if len(members) == 0: raise ValueError("#N/A 🚫 empty set")
                new_contexts = []
                for mem in members:
                    for c in contexts:
                        copied = dict(c)
                        copied[mem["dimension"]] = mem["key"]
                        new_contexts.append(copied)
                contexts = new_contexts
                continue
            if t == "kpi": kpi_handle = arg; continue
            raise ValueError("#VALUE! 🚫 unsupported handle passed to CUBEVALUE")
        if isinstance(arg, str):
            s = arg.strip()
            if s.startswith("[") and "measures" in s.lower():
                name = s.strip("[]").split("].[")[-1]
                measure_name = name
                continue
            if s in cube.measures_meta:
                measure_name = s
                continue
            if s in cube.dimensions or (s.startswith("[") and s.strip("[]").split("].[")[0] in cube.dimensions):
                try:
                    if s in cube.dimensions: dim = s
                    else:
                        token = s
                        if token.startswith("[") and token.endswith("]"): token = token[1:-1]
                        dim = token.split("].[")[0]
                    members_map = cube.dimensions.get(dim, {})
                    members = [{"dimension": dim, "key": k, "caption": v.get("caption", k)} for k, v in members_map.items()]
                    if len(members) == 0: raise ValueError("#N/A 🚫 The set is empty.")
                    new_contexts = []
                    for mem in members:
                        for c in contexts:
                            copied = dict(c)
                            copied[mem["dimension"]] = mem["key"]
                            new_contexts.append(copied)
                    contexts = new_contexts
                    continue
                except Exception as e: raise ValueError(str(e))
            try:
                dim, key, info = _resolve_member_inline(s)
            except Exception as e: raise ValueError(str(e))
            for c in contexts: c[dim] = key
            continue
        raise ValueError("#VALUE! 🚫 unsupported argument type for CUBEVALUE",info)
    if kpi_handle is not None:
        kpi_name = kpi_handle.get("kpi_name"); kpi_part = kpi_handle.get("kpi_part")
        if kpi_name is None or kpi_part is None: raise ValueError("#VALUE! 🚫 invalid KPI handle")
        total = 0.0
        for ctx in contexts:
            try:
                if hasattr(cube, "evaluate_kpi_part"): val = cube.evaluate_kpi_part(kpi_name, kpi_part, ctx)
                else:
                    parts = cube.kpis.get(kpi_name)
                    if parts is None: raise ValueError(f"#N/A 🚫 KPI not found: {kpi_name}")
                    part = parts.get(kpi_part)
                    if isinstance(part, str):
                        meta = cube.measures_meta.get(part)
                        if meta is None: val = 0.0
                        else:
                            col = meta.get("column"); agg = meta.get("agg","sum"); time_dim = meta.get("time_dim")
                            df = cube.data
                            if df is None: val = 0.0
                            else:
                                mask = None
                                for d,k in ctx.items():
                                    if d not in df.columns: raise ValueError(f"#N/A 🚫 Dimension column not found in data: {d}")
                                    m = (df[d]==k)
                                    mask = m if mask is None else (mask & m)
                                sub = df if mask is None else df[mask]
                                if sub.shape[0]==0: val = 0.0
                                else:
                                    series_vals = sub[col].astype(float)
                                    if agg=="sum": val = float(series_vals.sum())
                                    elif agg=="avg":
                                        vals = series_vals.dropna()
                                        if vals.shape[0]==0: raise ValueError("#DIV/0! 🚫 no numeric data for average")
                                        val = float(vals.mean())
                                    elif agg=="last_non_empty":
                                        if time_dim and time_dim in sub.columns:
                                            orded = sub.loc[series_vals.notna()].sort_values(by=time_dim)
                                            val = float(orded.iloc[-1][col]) if orded.shape[0]>0 else 0.0
                                        else: val = float(series_vals.sum())
                                    else: val = float(series_vals.sum())
                    elif callable(part):
                        try: val = part(cube, ctx)
                        except TypeError: val = part(ctx)
                    else: val = float(part)
                total += float(val)
            except ZeroDivisionError: raise ValueError("#DIV/0! 🚫 division by zero in KPI calculation")
            except Exception as e:
                msg = str(e)
                if msg.startswith("#"): raise
                raise ValueError(f"#VALUE! 🚫 {msg}")
        return float(total)
    if measure_name is None:
        if "Value" in cube.measures_meta: measure_name = "Value"
        else:
            if len(cube.measures_meta) == 1: measure_name = list(cube.measures_meta.keys())[0]
            else: raise ValueError("#N/A 🚫 No measure specified and no default available")
    def _eval_one_context(ctx):
        try:
            meta = cube.measures_meta.get(measure_name)
            if meta is None: raise ValueError(f"#N/A 🚫 Measure not found: {measure_name}")
            col = meta.get("column"); agg = meta.get("agg","sum"); time_dim = meta.get("time_dim")
            df = cube.data
            if df is None: raise ValueError("#N/A 🚫 Data not loaded in cube")
            mask = None
            for d,k in ctx.items():
                if d not in df.columns: raise ValueError(f"#N/A 🚫 Dimension column not found in data: {d}")
                m = (df[d] == k)
                mask = m if mask is None else (mask & m)
            sub = df if mask is None else df[mask]
            if sub.shape[0] == 0: return 0.0
            mul = None
            unary_present = False
            for d,k in ctx.items():
                info = cube.dimensions.get(d, {}).get(k, {})
                if info and info.get("unary", 1) != 1:
                    unary_present = True
                    break
            if unary_present:
                mul = (pd.Series(1, index=sub.index))
                for d,k in ctx.items():
                    info = cube.dimensions.get(d, {}).get(k, {})
                    u = info.get("unary", 1)
                    if u == -1: mul = mul * -1
                    elif u == 0: mul = mul * 0
                    else: mul = mul * 1
                series_vals = sub[col].astype(float) * mul
            else: series_vals = sub[col].astype(float)
            if agg == "sum": return float(series_vals.sum())
            if agg == "avg":
                vals = series_vals.dropna()
                if vals.shape[0] == 0: raise ValueError("#DIV/0! 🚫 no numeric data for average")
                return float(vals.mean())
            if agg == "last_non_empty":
                if time_dim is None or time_dim not in sub.columns: return float(series_vals.sum())
                ordered = sub.loc[series_vals.notna()].sort_values(by=time_dim)
                if ordered.shape[0] == 0: return 0.0
                return float(ordered.iloc[-1][col])
            return float(series_vals.sum())
        except ZeroDivisionError: raise
        except Exception as e:
            msg = str(e)
            if msg.startswith("#"): raise
            raise ValueError(f"#VALUE! 🚫 {msg}")
    result = 0.0
    for ctx in contexts:
        val = _eval_one_context(ctx)
        result += float(val)
    return float(result)

def CUMIPMT(rate, n_per, pv, start_period, end_period, payment_type=0) -> float:
    """
    =CUMIPMT(rate, n_per, pv, start_period, end_period, type)
    Returns the cumulative interest paid on a loan between start_period and end_period.

    Parameters:
        rate (float): Interest rate per period (e.g., monthly rate if payments are monthly).
        n_per (int): Total number of payment periods.
        pv (float): Present value (principal).
        start_period (int): First period in the range (1-based).
        end_period (int): Last period in the range (1-based).
        type (int, optional): When payments are due:
            0 = payment at end of period (default)
            1 = payment at beginning of period

    *Sample Code*:

     montly rate = 0.05 / 12 
     print(CUMIPMT(monthly_rate, 60, 10000, 1, 12, 0))  #-494.54
     print(CUMIPMT(monthly_rate, 60, 10000, 1, 12, 1))  #-7684.88 
    """
    import numpy as np
    try: rate = float(rate)
    except: raise ValueError("#VALUE! 🚫 rate must be numeric")
    try: n_per = int(n_per)
    except: raise ValueError("#VALUE! 🚫 n_per must be an integer")
    try: pv = float(pv)
    except: raise ValueError("#VALUE! 🚫 pv must be numeric")
    try: start_period = int(start_period); end_period = int(end_period)
    except: raise ValueError("#VALUE! 🚫 start_period and end_period must be integers")
    try: payment_type = int(payment_type)
    except: raise ValueError("#VALUE! 🚫 type must be 0 or 1")
    if n_per <= 0: raise ValueError("#NUM! 🚫 n_per must be > 0")
    if start_period < 1 or end_period < 1 or start_period > n_per or end_period > n_per: raise ValueError("#NUM! 🚫 start_period and end_period must be between 1 and n_per (inclusive)")
    if start_period > end_period: raise ValueError("#NUM! 🚫 start_period cannot be greater than end_period")
    if payment_type not in (0, 1): raise ValueError("#VALUE! 🚫 type must be 0 (end) or 1 (beginning)")
    if np.isclose(rate, 0.0): return -0.0
    pv_adj = pv if payment_type == 0 else (pv / (1.0 + rate))
    try: pmt = - (pv_adj * rate) / (1.0 - np.power(1.0 + rate, -n_per))
    except FloatingPointError: raise ValueError("#DIV/0! 🚫 numerical issue computing payment (check rate/n_per)")
    ipmt_for_period = lambda per: pv_adj * rate * np.power(1.0 + rate, per - 1) + \
                                  pmt * (np.power(1.0 + rate, per - 1) - 1.0)
    total_interest = np.sum([ipmt_for_period(per) for per in range(start_period, end_period + 1)])
    return -float(total_interest)

def CUMPRINC(rate, n_per, pv, start_period, end_period, payment_type=0) -> float:
    """
    =CUMPRINC(rate, n_per, pv, start_period, end_period, type)
    Returns the cumulative principal paid on a loan between start_period and end_period.

    Parameters:
        rate (float): Interest rate per period.
        n_per (int): Total number of payment periods.
        pv (float): Present value (principal).
        start_period (int): First period in the range (1-based).
        end_period (int): Last period in the range (1-based).
        type (int, optional): When payments are due:
            0 = payment at end of period (default)
            1 = payment at beginning of period

    *Sample Code*:

     montly rate = 0.05 / 12 
     print(CUMPRINC(monthly_rate, 60, 10000, 1, 12, 0))  #-516.43 
     print(CUMPRINC(monthly_rate, 60, 10000, 1, 12, 1))  #-7663.00 
    """
    import numpy as np
    try: rate = float(rate)
    except: raise ValueError("#VALUE! 🚫 rate must be numeric")
    try: n_per = int(n_per)
    except: raise ValueError("#VALUE! 🚫 n_per must be an integer")
    try: pv = float(pv)
    except: raise ValueError("#VALUE! 🚫 pv must be numeric")
    try: start_period = int(start_period); end_period = int(end_period)
    except: raise ValueError("#VALUE! 🚫 start_period and end_period must be integers")
    try: payment_type = int(payment_type)
    except: raise ValueError("#VALUE! 🚫 type must be 0 or 1")
    if n_per <= 0: raise ValueError("#NUM! 🚫 n_per must be > 0")
    if start_period < 1 or end_period < 1 or start_period > n_per or end_period > n_per: raise ValueError("#NUM! 🚫 start_period and end_period must be between 1 and n_per (inclusive)")
    if start_period > end_period: raise ValueError("#NUM! 🚫 start_period cannot be greater than end_period")
    if payment_type not in (0, 1): raise ValueError("#VALUE! 🚫 type must be 0 (end) or 1 (beginning)")
    if np.isclose(rate, 0.0):
        pmt = -pv / n_per
        return float(pmt * (end_period - start_period + 1))
    pv_adj = pv if payment_type == 0 else (pv / (1.0 + rate))
    try: pmt = - (pv_adj * rate) / (1.0 - np.power(1.0 + rate, -n_per))
    except FloatingPointError: raise ValueError("#DIV/0! 🚫 numerical issue computing payment (check rate/n_per)")
    ipmt_for_period = lambda per: pv_adj * rate * np.power(1.0 + rate, per - 1) + \
                                  pmt * (np.power(1.0 + rate, per - 1) - 1.0)
    total_interest = np.sum([ipmt_for_period(per) for per in range(start_period, end_period + 1)])
    total_payments = pmt * (end_period - start_period + 1)
    return float(total_payments - total_interest)

