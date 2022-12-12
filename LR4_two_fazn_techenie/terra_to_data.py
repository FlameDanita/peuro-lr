import pandas as pd
from scipy.interpolate import interp2d, interp1d
from numpy import pi
import os

_df = None

def open_terra_df(file_path):
    global _df
    _df = pd.read_excel(os.path.dirname(__file__) + '\\' + file_path, header=[1])
    _df.columns = _df.columns.str.lstrip()

def to_data(key,file_path = None):
    if key not in _df:
        raise KeyError(f'Key {key} not found. Use `get_keys()` for more information')
    ps = list(_df['p'][_df['T'] == 500])
    Ts = list(_df['T'][_df['p'] == ps[0]])
    if file_path:
        data = {}
        for p in ps:
            data[p] = list(_df[key][_df['p'] == p])
        
        resdf = pd.DataFrame(data,index = Ts)
        with pd.ExcelWriter(os.path.dirname(__file__) + '\\' + file_path) as writer:
            resdf.to_excel(writer)
    else:
        matrix = []
        for p in ps:
            matrix.append(list(_df[key][_df['p'] == p]))
        
        resint = interp2d(Ts,ps,matrix)
        return resint

def get_keys():
    return list(_df.keys())

def interp_from_data(file_path):
    dfinput = pd.read_excel(os.path.dirname(__file__) + '\\' + file_path, index_col=0)
    Ts = list(dfinput.index)
    ps = list(dfinput.columns)

    matrix = []
    for p in ps:
        matrix.append(list(dfinput[p]))

    resint = interp2d(Ts,ps,matrix)
    return resint

_dfs = None

def open_soplo(file_path):
    global _dfs
    _dfs = pd.read_excel(os.path.dirname(__file__) + '\\' + file_path, header=None)

def soplo():
    xs = []
    Rs = []
    for x,R in zip(_dfs[0],_dfs[1]):
        xs.append(x / 1e3)
        Rs.append(R / 1e3)
    
    Ri = interp1d(xs,Rs)

    return xs, Rs, Ri

def F_soplo():
    xs, Rs, Ri = soplo()
    Fs = [pi * R**2 for R in Rs]

    Fi = interp1d(xs, Fs)

    return Fi

def get_a(file_path):
    df = pd.read_excel(os.path.dirname(__file__) + '\\' + file_path, sheet_name='as')
    pls = list(df['pl'])
    a1s = list(df['a1'])
    a2s = list(df['a2'])
    a3s = list(df['a3'])

    a1 = interp1d(pls,a1s)
    a2 = interp1d(pls,a2s)
    a3 = interp1d(pls,a3s)

    return a1, a2, a3

def get_b(file_path):
    df = pd.read_excel(os.path.dirname(__file__) + '\\' + file_path, sheet_name='bs')
    pls = list(df['pl'])
    b1s = list(df['b1'])
    b2s = list(df['b2'])

    b1 = interp1d(pls,b1s)
    b2 = interp1d(pls,b2s)

    return b1, b2

def get_c(file_path):
    df = pd.read_excel(os.path.dirname(__file__) + '\\' + file_path, sheet_name='cs')
    pls = list(df['pl'])
    c1s = list(df['c1'])
    c2s = list(df['c2'])
    c3s = list(df['c3'])
    c4s = list(df['c4'])

    c1 = interp1d(pls,c1s)
    c2 = interp1d(pls,c2s)
    c3 = interp1d(pls,c3s)
    c4 = interp1d(pls,c4s)

    return c1, c2, c3, c4