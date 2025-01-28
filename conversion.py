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
        return onlyone((y, x) for y,x in metrodata.STATIONS.items() if x['name'] == name)[1]['line']
    except ValueError:
        return name


def get_metrocircle(name):
    import metrodata
    try:
        stationid = onlyone((y, x) for y,x in metrodata.STATIONS.items() if x['name'] == name)[0]
    except ValueError:
        return '', ''

    CIRCLE = 5
    # dijkstra baby
    openset = dict()
    prev = dict()
    closedset = set()
    openset[stationid] = 0    
    while openset:
        current = next(iter(openset))
        currentval = openset[current]

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
            return onlyone((y, x) for y,x in metrodata.STATIONS.items() if x['name'] == name)[0]
        except ValueError:
            return None

    stationid = get_stationid(name)
    if stationid is None:
        return ('', ) * len(targets)

    targets_id = set(map(get_stationid, targets))
    if any(map(lambda x:x is None, targets_id)):
        return ('', ) * len(targets)

    # dijkstra baby
    openset = dict()
    prev = dict()
    closedset = dict()
    openset[stationid] = 0    
    while openset:
        current = next(iter(openset))
        currentval = openset[current]

        del openset[current]
        closedset[current] = currentval

        if targets_id <= closedset.keys():
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
        
    if not openset:
        return ('', ) * len(targets)
    else:
        return list(map(lambda target_id: closedset[target_id] / 60, targets_id))
    

def onlyone(it):
    if len(L := list(it)) != 1:
        raise ValueError('No values' if len(L) == 0 else 'Too much values')
    return L[0]

wb = xl.load_workbook(input_filename := onlyone(glob('offers*xlsx')))
ws = wb.worksheets[0]
headers = [x.value for x in list(ws.rows)[0]]
METRO = 'Метро'
iMETRO = headers.index(METRO)

def icell(i, j):
    return ws.cell(row=i+1, column=j+1)

def iinsertcols(i, *args, **kwargs):
    ws.insert_cols(1+i, *args, **kwargs)

col = iMETRO
iinsertcols(col)
icell(0, col).value = "Metro (Foot)"

for i in range(1, len(list(ws.rows))):
    icell(i, col).value = get_minute_foot(icell(i, col+1).value)

col += 1
iinsertcols(col)
icell(0, col).value = "Metro (Location)"

for i in range(1, len(list(ws.rows))):
    icell(i, col).value = get_metro_name(icell(i, col+1).value)

headers = [x.value for x in list(ws.rows)[0]]
iSURFACE = headers.index("Площадь, м2")

col = iSURFACE
iinsertcols(col)
icell(0, col).value = "Surface (m²)"

for i in range(1, len(list(ws.rows))):
    icell(i, col).value = get_surface_main(icell(i, col+1).value)

headers = [x.value for x in list(ws.rows)[0]]
col = headers.index("Цена")

iinsertcols(col)
icell(0, col).value = "Price (rub)"

for i in range(1, len(list(ws.rows))):
    icell(i, col).value = get_price(icell(i, col+1).value)

headers = [x.value for x in list(ws.rows)[0]]
col = headers.index("Metro (Location)")

iinsertcols(col)
icell(0, col).value = "Metro (Line)"

for i in range(1, len(list(ws.rows))):
    icell(i, col).value = get_metroline(icell(i, col+1).value)

# time to circle
headers = [x.value for x in list(ws.rows)[0]]
col = headers.index("Metro (Location)")

outcols = ["Metro (Circle destination)", "Metro (Circle time)"]
ninsert = len(outcols)
for _ in range(ninsert):
    iinsertcols(col)
for i in range(ninsert):
    icell(0, col+i).value = outcols[i]

for i in range(1, len(list(ws.rows))):
    computation = get_metrocircle(icell(i, col+ninsert).value)
    for j in range(ninsert):
        icell(i, col+j).value = computation[j]

# targets
from functools import partial
input_header = "Metro (Location)"
outcols = ["Baumanska (Time)", "Aeroport (Time)"]
targets = ["Бауманская", "Аэропорт"]
function = partial(get_target_station_time, targets=targets)

# function
headers = [x.value for x in list(ws.rows)[0]]
col = headers.index(input_header)

outcols = outcols
ninsert = len(outcols)
for _ in range(ninsert):
    iinsertcols(col)
for i in range(ninsert):
    icell(0, col+i).value = outcols[i]

for i in range(1, len(list(ws.rows))):
    computation = function(icell(i, col+ninsert).value)
    for j in range(ninsert):
        icell(i, col+j).value = computation[j]

wb.save(output_filename := 'rich_offers_v2_' + re.sub('^offers|[.]xlsx$', '', input_filename).strip() + '.xlsx')
print("Saved:", output_filename)
