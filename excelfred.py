"""
excelfred - A beginner-friendly open-source Python package recreating Excel 514 functions
`Author: Samuel Raj P (FRED)`
https://www.linkedin.com/in/samuel-raj23
"""

import pandas as pd, numpy as np

def ABS(*args: int | float | str) -> float:
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

def ACCRINT(issue: str, first_interest: str, settlement: str, rate: float, par: float = 15000.00, frequency: int = 1, basis: int = 0, calc_method: bool = True) -> float:
    """
    `=ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])`

    Calculates accrued interest for a security that pays periodic interest.

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

def ACCRINTM(issue: str, maturity: str, rate: float, par: float = 15000.00, basis: int = 0) -> float:
    """
    `=ACCRINTM(issue, settlement maturity, rate, par, [basis])`

    Returns the accrued interest for a security that pays interest at maturity (no periodic coupons).

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
    total = 0
    for arg in args:
        if isinstance(arg, pd.DataFrame): total += len(arg.columns)
        elif isinstance(arg, pd.Series): total += 1
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
    average_range = np.array(average_range)
    mask = np.ones(len(average_range), dtype=bool)
    
    for i in range(0, len(criteria_pairs), 2):
        crit_range = np.array(criteria_pairs[i])
        criteria = criteria_pairs[i+1]
        mask &= pd.Series(crit_range).apply(lambda x: eval(f"x{criteria}" if criteria[0] in "<>=" else f"x=={repr(criteria)}"))
    filtered = average_range[mask]
    return np.mean(filtered) if len(filtered) > 0 else np.nan

import numpy as np
import pandas as pd

def bahttext(number: int | float) -> str:
    """
    `=BAHTTEXT(number)` Converts a number to a text **(baht)**

    *Example Inputs*:
     
     print(bahttext(1234.56))  # หนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบหกสตางค์
     print(bahttext(5000))     # ห้าพันบาทถ้วน
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

