import os

_materials = None
_materials_path = os.path.join(os.path.dirname(__file__), 'materials\Materials for RDTT.csv')

def _init_material(mat_path):
    global _materials 
    _materials = {}
    headers = ['material', 'rho', 'sigma_02', 'sigma_v', 'delta', 'sigma_v_rho', 'use']
    with open(mat_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            pd = {headers[0]: values[0]}
            for h, val in zip(headers[1:-1], values[1:-1]):
                pd[h] = float(val)
            pd['use'] = values[-1][:-1]
            _materials[values[0]] = pd

def get_material(material):
    """Возвращает словарь с параметрами материала с именем `material` из таблицы 1
    лекции 3 раздаточного материала.

    Аргументы
    ---------
        material : str
            Название стали, титанового или аллюминиевого сплава. Список материалов
            можно узнать, вызвав функцию `get_material_names()`.

    Выходные параметры
    ------------------
        dict: словарь следующего вида
            {
                'material' : str
                    Название материала.
                'rho' : float
                    Плотность материала, в кг/м^3.
                'sigma_02' : float
                    Предел текучести при растяжении, в Па.
                'sigma_v' : float
                    Временное сопротивление при растяжении, в Па.
                'delta' : float
                    Относительное удлинение, в -.
                'sigma_v_rho' : float
                    Удельная прочность, в м.
                'use' : str
                    Рекомендуемое применение.
            }
    """
    if _materials is None:
        _init_material(_materials_path)
    if material not in _materials:
        raise ValueError(f'Такого материала в таблице нет: {material}. Список доступных имен можно получить из функции get_material_names()')
    return _materials[material]

def get_material_names():
    """Возвращает список материалов из таблицы 1 лекции 3 раздаточного материала.

    Выходные параметры
    ------------------
        list: список материалов
    """
    if _materials is None:
        _init_material(_materials_path)
    return list(_materials.keys())

_composites = None
_composites_path = os.path.join(os.path.dirname(__file__), 'materials\Composite.csv')

def _init_composite(comp_path):
    global _composites 
    _composites = {}
    headers = ['material', 'rho', 'sigma_r', 'E']
    with open(comp_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            pd = {headers[0]: values[0]}
            for h, val in zip(headers[1:], values[1:]):
                pd[h] = float(val)
            _composites[values[0]] = pd

def get_composite(composite):
    """Возвращает словарь с параметрами композитов с именем `composite` из таблицы 2
    лекции 3 раздаточного материала.

    Аргументы
    ---------
        composite : str
            Название композитного материала. Список материалов
            можно узнать, вызвав функцию `get_composite_names()`.

    Выходные параметры
    ------------------
        dict: словарь следующего вида
            {
                'material' : str
                    Название композита.
                'rho' : float
                    Плотность материала, в кг/м^3.
                'sigma_r' : float
                    Прочность при растяжении, в Па.
                'E' : float
                    Модуль упругости, в Па.
            }
    """
    if _composites is None:
        _init_composite(_composites_path)
    if composite not in _composites:
        raise ValueError(f'Такого композита в таблице нет: {composite}. Список доступных имен можно получить из функции get_composite_names()')
    return _composites[composite]

def get_composite_names():
    """Возвращает список композитных материалов из таблицы 2 лекции 3 раздаточного материала.

    Выходные параметры
    ------------------
        list: список композитов.
    """
    if _composites is None:
        _init_composite(_composites_path)
    return list(_composites.keys())

_tzp = None
_tzp_path = os.path.join(os.path.dirname(__file__), 'materials\TZP.csv')

def _init_tzp(tzp_path):
    global _tzp
    _tzp = {}
    headers = ['material', 'rho', 'delta_razr', 'sigma_razr', 'lambda_t', 'c_p']
    with open(tzp_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            pd = {headers[0]: values[0]}
            for h, val in zip(headers[1:], values[1:]):
                pd[h] = float(val)
            _tzp[values[0]] = pd

def get_tzp(tzp):
    """Возвращает словарь с параметрами ТЗП с именем `tzp` из таблицы 4
    лекции 3 раздаточного материала.

    Аргументы
    ---------
        tzp : str
            Название ТЗП. Список ТЗП можно узнать, вызвав функцию `get_tzp_names()`.

    Выходные параметры
    ------------------
        dict: словарь следующего вида
            {
                'material' : str
                    Название ТЗП.
                'rho' : float
                    Плотность материала, в кг/м^3.
                'delta_razr' : float
                    Относительное удлинение при разрыве, в -.
                'sigma_razr' : float
                    Предел прочности при разрыве, в Па.
                'lambda_t' : float
                    Коэффициент теплопроводности, в Вт/(м*К).
                'c_p' : float
                    Удельная теплоёмкость, в Дж/(кг*К).
            }
    """
    if _composites is None:
        _init_tzp(_tzp_path)
    if tzp not in _tzp:
        raise ValueError(f'Такого ТЗП в таблице нет: {tzp}. Список доступных имен можно получить из функции get_tzp_names()')
    return _tzp[tzp]

def get_tzp_names():
    """Возвращает список ТЗП из таблицы 4 лекции 3 раздаточного материала.

    Выходные параметры
    ------------------
        list: список ТЗП
    """
    if _tzp is None:
        _init_tzp(_tzp_path)
    return list(_tzp.keys())

_mater_for_flan_and_shpan = None
_mater_for_flan_and_shpan_path = os.path.join(os.path.dirname(__file__), 'materials\Materials for flanec and shpangout.csv')

def _init_flan_shpan_mater(mat_path):
    global _mater_for_flan_and_shpan 
    _mater_for_flan_and_shpan = {}
    headers = ['material', 'rho', 'sigma_v', 'sigma_v_rho', 'use']
    with open(mat_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            pd = {headers[0]: values[0]}
            for h, val in zip(headers[1:-1], values[1:-1]):
                pd[h] = float(val)
            pd['use'] = values[-1][:-1]
            _mater_for_flan_and_shpan[values[0]] = pd

def get_fln_shpn_mater(material):
    """Возвращает словарь с параметрами материала для фланцев и шпангоутов с именем `material`
    из таблицы 3 лекции 3 раздаточного материала.

    Аргументы
    ---------
        material : str
            Название материала. Список материалов
            можно узнать, вызвав функцию `get_fln_shpn_mater_names()`.

    Выходные параметры
    ------------------
        dict: словарь следующего вида
            {
                'material' : str
                    Название материала.
                'rho' : float
                    Плотность материала, в кг/м^3.
                'sigma_v' : float
                    Временное сопротивление при растяжении, в Па.
                'sigma_v_rho' : float
                    Удельная прочность, в м.
                'use' : str
                    Рекомендуемое применение.
            }
    """
    if _mater_for_flan_and_shpan is None:
        _init_flan_shpan_mater(_mater_for_flan_and_shpan_path)
    if material not in _mater_for_flan_and_shpan:
        raise ValueError(f'Такого материала для фланцев и шпангоутов в таблице нет: {material}. Список доступных имен можно получить из функции get_fln_shpn_mater_names()')
    return _mater_for_flan_and_shpan[material]

def get_fln_shpn_mater_names():
    """Возвращает список материалов для фланцев и шпангоутов из таблицы 3 лекции 3 раздаточного материала.

    Выходные параметры
    ------------------
        list: список материалов
    """
    if _mater_for_flan_and_shpan is None:
        _init_flan_shpan_mater(_mater_for_flan_and_shpan_path)
    return list(_mater_for_flan_and_shpan.keys())

# Propellant

# Mixed propellant
_mixed_propellant = None
_mixed_propellant_path = os.path.join(os.path.dirname(__file__), 'propellant\Mixed propellant.csv')

def _init_mixed_propellant(mat_path):
    global _mixed_propellant 
    _mixed_propellant = {}
    headers = ['Propellant', 'Number', 'I_ud', 'rho_т', 'R_г', 'k', 'T_0', 'nu', 'u_1', 'D_t', 'p_min']
    with open(mat_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            pd = {headers[0]: values[0]}
            pd = {headers[1]: values[1]}
            for h, val in zip(headers[2:], values[2:]):
                pd[h] = float(val)
            _mixed_propellant[values[0]] = pd
            

def get_mixed_propellant(mixed_propellant):
    """Возвращает словарь с параметрами топлива из таблицы О.С. Серпинского.

    Аргументы
    ---------
        mixed_propellant : str
            Название топлива, список топлив можно узнать, вызвав функцию `get_mixed_propellant_names()`.

    Выходные параметры
    ------------------
        dict: словарь следующего вида
            {
                'Number' : float
                    Порядковый номер топлива.
                'I_ud' : float
                    Удельный импульс топлива, в м/с.
                'rho_т' : float
                    Плотность топлива, в кг/м^3.
                'R_г' : float
                    Газовая постоянная продуктов сгорания, в Дж/кг*К.
                'k' : float
                    Показатель адиабаты продуктов сгорания.
                'T_0' : float
                    Температура торможения продуктов сгорания, в К.
                'nu' : float
                    Показатель степени в законе горения.
                'u_1' : float
                    Единичная скорость горения, в м/с * МПа.
                'D_t' : float
                    Коэффициент для температурной зависимости, в 1/К.
                'p_min' : float
                    Минимальное давление для устойчивого горения смесевого топлива, в Па.
            }
    """
    if _mixed_propellant is None:
        _init_mixed_propellant(_mixed_propellant_path)
    if mixed_propellant not in _mixed_propellant:
        raise ValueError(f'Такого топлива в таблице нет: {mixed_propellant}. Список доступных имен можно получить из функции get_mixed_propellant_names()')
    return _mixed_propellant[mixed_propellant]

def get_mixed_propellant_names():
    """Возвращает список смесевых топлив из таблицы О.С. Серпинского.

    Выходные параметры
    ------------------
        list: список топлив
    """
    if _mixed_propellant is None:
        _init_mixed_propellant(_mixed_propellant_path)
    return list(_mixed_propellant.keys())

# Ballistic propellant
_ball_propellant = None
_ball_propellant_path = os.path.join(os.path.dirname(__file__), 'propellant\Bal propellant.csv')

def _init_ball_propellant(mat_path):
    global _ball_propellant 
    _ball_propellant = {}
    headers = ['Number', 'I_ud', 'rho_т', 'R_г', 'k', 'T_0', 'nu', 'u_1', 'B_т', 'Q', 'p_min']
    with open(mat_path, encoding='utf-8')  as f:
        f.readline()
        for line in f.readlines():
            values = line.split(';')
            pd = {
                headers[0]: values[0],
                headers[1]: values[1]
                }
            for h, val in zip(headers[2:], values[2:]):
                pd[h] = float(val)
            _ball_propellant[values[0]] = pd
            

def get_ball_propellant(propellant_number):
    """Возвращает словарь с параметрами топлива из банка баллиститных топлив.

    Аргументы
    ---------
        propellant_number : str
            Порядковый номер топлива из таблицы, список топлив можно узнать, вызвав функцию `get_ball_propellant_names()`.

    Выходные параметры
    ------------------
        dict: словарь следующего вида
            {
                'Number' : float
                    Порядковый номер топлива.
                'I_ud' : float
                    Удельный импульс топлива, в м/с.
                'rho_т' : float
                    Плотность топлива, в кг/м^3.
                'R_г' : float
                    Газовая постоянная продуктов сгорания, в Дж/кг*К.
                'k' : float
                    Показатель адиабаты продуктов сгорания.
                'T_0' : float
                    Температура торможения продуктов сгорания, в К.
                'nu' : float
                    Показатель степени в законе горения.
                'u_1' : float
                    Единичная скорость горения, в м/с * МПа.
                'B_т': float
                    Коэффициент для температурной зависимости, в К.
                'Q': float
                    Калорийность топлива, в МДж/кг.
                'p_min' : float
                    Минимальное давление для устойчивого горения топлива, в Па.
            }
    """
    if _ball_propellant is None:
        _init_ball_propellant(_ball_propellant_path)
    if propellant_number not in _ball_propellant:
        raise ValueError(f'Такого номера топлива в таблице нет: {propellant_number}. Список доступных имен можно получить из функции get_ball_propellant_names()')
    return _ball_propellant[propellant_number]

def get_ball_propellant_names():
    """Возвращает список топлив из банка баллиститных топлив.

    Выходные параметры
    ------------------
        list: список топлив
    """
    if _ball_propellant is None:
        _init_ball_propellant(_ball_propellant_path)
    return list(_ball_propellant.keys())
