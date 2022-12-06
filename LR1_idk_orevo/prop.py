import os

__prop = None
__prop_path = os.path.join(os.path.dirname(__file__), 'propellants.csv')

def __init_prop():
    global __prop
    __prop = {}
    headers = ['name', 'rho', 'f', 'R', 'T0', 'k', 'I_ud', 'B', 'u_1', 'nu']
    with open(__prop_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            print(values)
            pd = {headers[0]: values[0]}
            for head, val in zip(headers[1:], values[1:]):
                pd[head] = float(val)
            __prop[values[0]] = pd

class __Prop:
    def __init__(self, dic) -> None:
        self.rho = dic['rho']
        self.f = dic['f']
        self.R = dic['R']
        self.T0 = dic['T0']
        self.k = dic['k']
        self.I_ud = dic['I_ud']
        self.B = dic['B']
        self.F1 = _get_u(dic['name'])
        self.u_1 = dic['u_1']
        self.nu = dic['nu']
        self.sigma_t = 300
        self.lambda_t = 1.2
        self.nu_t = 0.7
        self.c_t = 1250
        self.T_s = 550
    
    def F3(self,T_n):
        return self.B / (self.B - (T_n - 291.15))

    def u(self,p,T_n):
        return self.F1(p/1e6) * self.F3(T_n)

def _get_u(name):

    def P_1(p):
        if 5 < p < 45:
            return 0.003 * p**0.7
        elif p < 60:
            return 0.00046*(p - 42)**1.17 + 0.0381
        else:
            return 0.00086 * p

    def P_2(p):
        # if 20 <= p <= 150:
            return 0.0000315 * (9.81 * p) ** 1.17
        # else:
            # raise ValueError(f'Давление {p:.3f} МПа выходит за рамки диапазона [20;150]')
    
    def P_3(p):
        # if 39 <= p <= 200:
            return 0.000306 * (9.81 * p) ** 0.78
        # else:
            # raise ValueError(f'Давление {p:.3f} МПа выходит за рамки диапазона [39;200]')

    def B_1(p):
        # if 30 <= p <= 80:
            return 0.00294 * p ** 0.65
        # else:
        #     raise ValueError(f'Давление {p:.3f} МПа выходит за рамки диапазона [30;80]')
    
    def B_2(p):
        # if 34 <= p <= 150:
            return 0.000198 * (9.81 * p) ** 0.89
        # else:
        #     raise ValueError(f'Давление {p:.3f} МПа выходит за рамки диапазона [34;150]')

    def B_3(p):
        # if 16 <= p <= 150:
            return 0.00085 * (9.81 * p) ** 0.69
        # else:
        #     raise ValueError(f'Давление {p:.3f} МПа выходит за рамки диапазона [16;150]')
    
    def T_1(p):
            return 8.58e-3 * p**-0.151
    
    def T_2(p):
            return 7.58e-3 * p**0.0768
        
    def T_3(p):
            return 6.45e-3 * p**-0.095
        
    foo = {
        'P-1': P_1,
        'P-2': P_2,
        'P-3': P_3,
        'B-1': B_1,
        'B-2': B_2,
        'B-3': B_3,
        'T-1': T_1,
        'T-2': T_2,
        'T-3': T_3
    }.get(name)
    
    return foo

def get_propellant(prop_name):
    if __prop is None:
        __init_prop()
    if prop_name not in __prop:
        raise ValueError(f'Такого топлива в таблице нет: {prop_name}.')
    return __Prop(__prop[prop_name])