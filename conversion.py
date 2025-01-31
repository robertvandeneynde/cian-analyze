import openpyxl as xl
from glob import glob
import re

def get_minute_foot(x):
    try:
      return int(re.compile(r'(\d+) мин пешком').search(x).group(1))
    except AttributeError:
       raise ValueError(f"Cannot find {x!r}")
    except TypeError:
        raise ValueError(f"Cannot find {x!r}")


def get_metro_name(x):
    try:
      return str(re.compile(r'м[.] (.*) [(]\d+ мин пешком[)]').fullmatch(x).group(1))
    except (AttributeError, TypeError):
       raise ValueError(f"Cannot find {x!r}")

def get_surface_main(x):
    try:
      return float(x.split("/")[0])
    except ValueError:
       raise ValueError(f"Cannot find surface in {x!r}")

def get_price(x):
    try:
      return int(re.compile(r'(\d+)[.]\d+' + ' ' + re.escape('руб./ За месяц') + '.*').fullmatch(x).group(1))
    except (ValueError, AttributeError):
       raise ValueError(f"Cannot find price in {x!r}")

def get_metroline(name):
    import metrodata
    CIRCLE_LINE = 5
    try:
        return next(iter(((y, x) for y,x in metrodata.STATIONS.items() if x['name'] == name)))[1]['line']
    except StopIteration:
        return name


def get_metrocircle(name):
    import metrodata
    try:
        stationid = next(iter(((y, x) for y,x in metrodata.STATIONS.items() if x['name'] == name)))[0]
    except StopIteration:
        return '', ''

    CIRCLE = 5
    # dijkstra baby
    openset = dict()
    prev = dict()
    closedset = set()
    openset[stationid] = 0    
    while openset:
        current, currentval = min(openset.items(), key=lambda x:x[1])

        if metrodata.STATIONS[current]['line'] == CIRCLE:
            destination = current
            destinationval = currentval
            break

        del openset[current]
        closedset.add(current)

        neighs = (
            [(y, d) for x,y,d in metrodata.LINKS if x == current] +
            [(x, d) for x,y,d in metrodata.LINKS if y == current])

        for o, d in neighs:
            if o in closedset:
                pass
            elif o not in openset or currentval + d < openset[o]:
                openset[o] = currentval + d
                prev[o] = current
                continue
        
    if not openset:
        return '', ''
    else:
        return metrodata.STATIONS[destination]['name'], destinationval / 60


def get_target_station_time(name, targets:list[str]):
    import metrodata

    def get_stationid(name):
        try:
            return next(iter((y, x) for y,x in metrodata.STATIONS.items() if x['name'] == name))[0]
        except StopIteration:
            return None

    stationid = get_stationid(name)
    if stationid is None:
        return ('', ) * len(targets)

    targets_id_list = list(map(get_stationid, targets))
    targets_id_set = set(targets_id_list)
    if None in targets_id_set:
        return ('', ) * len(targets)

    # dijkstra baby
    openset = dict()
    prev = dict()
    closedset = dict()
    openset[stationid] = 0    
    while openset:
        current, currentval = min(openset.items(), key=lambda x:x[1])

        del openset[current]
        closedset[current] = currentval

        if targets_id_set <= closedset.keys():
            break

        neighs = (
            [(y, d) for x,y,d in metrodata.LINKS if x == current] +
            [(x, d) for x,y,d in metrodata.LINKS if y == current])

        for o, d in neighs:
            if o in closedset:
                pass
            elif o not in openset or currentval + d < openset[o]:
                openset[o] = currentval + d
                prev[o] = current
                continue
        
    def reconstruct_path(x_station_id):
        x = x_station_id
        L = [x]
        while x in prev:
            L.append(x := prev[x])
        return list(reversed(L))

    def map_station_name(L):
        return [metrodata.STATIONS[id]['name'] for id in L]

    if not targets_id_set <= closedset.keys():
        return ('', ) * len(targets)
    else:
        return list(map(lambda target_id: closedset[target_id] / 60, targets_id_list))
    

def onlyone(it):
    if len(L := list(it)) != 1:
        raise ValueError('No values' if len(L) == 0 else 'Too much values')
    return L[0]

wb = xl.load_workbook(input_filename := onlyone(glob('offers*xlsx')))
ws = wb.worksheets[0]

def icell(i, j):
    return ws.cell(row=i+1, column=j+1)

def iinsertcols(i, *args, **kwargs):
    ws.insert_cols(1+i, *args, **kwargs)

def one_to_N(input_header, outcols, function):
    headers = [x.value for x in list(ws.rows)[0]]
    col = headers.index(input_header)

    ninsert = len(outcols)
    for _ in range(ninsert):
        iinsertcols(col)
    for i in range(ninsert):
        icell(0, col+i).value = outcols[i]

    for i in range(1, len(list(ws.rows))):
        computation = function(icell(i, col+ninsert).value)
        for j in range(ninsert):
            icell(i, col+j).value = computation[j]

def one_to_one(input_header, outcol, function):
    from functools import partial
    one_to_N(input_header, outcols=[outcol], function=lambda x: [function(x)])

# metro (foot)
one_to_one(
    input_header="Метро",
    outcol="Metro (Foot)",
    function=get_minute_foot)

# metro (location)
one_to_one(
    input_header="Метро",
    outcol="Metro (Location)",
    function=get_metro_name)

# surface
one_to_one(
    input_header="Площадь, м2",
    outcol="Surface (m²)",
    function=get_surface_main)

# price
one_to_one(
    input_header="Цена",
    outcol="Price (rub)",
    function=get_price)

# metroline
one_to_one(
    input_header="Metro (Location)",
    outcol="Metro (Line)",
    function=get_metroline)

# time to circle
one_to_N(
    input_header="Metro (Location)",
    outcols=["Metro (Circle destination)", "Metro (Circle time)"],
    function=get_metrocircle) 

# targets
from functools import partial 
one_to_N(
    input_header="Metro (Location)",
    outcols=["Baumanska (Time)", "Aeroport (Time)"],
    function=partial(get_target_station_time, targets=["Бауманская", "Аэропорт"]))

wb.save(output_filename := 'rich_offers_v4_' + re.sub('^offers|[.]xlsx$', '', input_filename).strip() + '.xlsx')
print("Saved:", output_filename)
