from string import ascii_uppercase
import pylightxl as xl
import sys

def get_qty_from_refdes(str):
    str = "".join(c for c in str if c not in ascii_uppercase)
    str_arr = str.split(',')
    sum_qty = 0
    for i in range(0, len(str_arr)):
        if '-' in str_arr[i]:
            sum_qty += abs(eval(str_arr[i])) + 1
        else:
            sum_qty += 1
    return sum_qty

if len(sys.argv) == 4:
    filename = sys.argv[1]
    refdes_col = sys.argv[2]
    qty_col = sys.argv[3]
else:
    sys.exit("Wrong args given")

bom = xl.readxl(filename)

refdes_list = bom.ws(ws=filename[:-5]).col(col=int(refdes_col))[3:]
qty_list = bom.ws(ws=filename[:-5]).col(col=int(qty_col))[3:]

mismatch_flag = False
for i in range(0, len(refdes_list)):
    #print(f'{refdes_list[i]:15}|{qty_list[i]:10}')
    mismatch = get_qty_from_refdes(refdes_list[i]) != int(qty_list[i])
    if mismatch:
        print(f'Found mismatch in excel row {i+3}, refdes {refdes_list[i]}, calculated: {get_qty_from_refdes(refdes_list[i])}, actual: {int(qty_list[i])}')
        mismatch_flag = True

if mismatch_flag:
    print("Found mismatch in refdes and qty in excel, view messages above")
else:
    print("Check finished successfully, mismatches are not found")
