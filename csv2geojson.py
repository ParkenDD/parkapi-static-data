import sys
import csv
import json
from pathlib import Path

argv = sys.argv[1:]
if len(argv) == 0:
    print("csv2geojson: missing csv file")
    sys.exit(2)
csvfile = Path(argv[0])

features = []
with csvfile.open('r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        if not row.get('uid'):
            print(f'Attribute \'uid\' missing in row {row}')
            continue
        if not row.get('lat') or not row.get('lon'):
            print(f'Attributes \'lat\'/\'lon\' missing in row {row}')
            continue
        properties = { 'uid': row['uid'] }
        for column in ['address', 'type']:
            if row.get(column):
                properties[column] = row[column]
        for column in ['max_height', 'max_width', 'max_depth']:
            if row.get(column) and row[column].isdigit():
                properties[column] = int(row[column])
        if row.get('park_and_ride_type'):
            properties['park_and_ride_type'] = [row['park_and_ride_type']]
        if row.get('DHID'):
            properties['external_identifiers'] = [{ 'type': 'DHID', 'value': row['DHID'] }]
        features.append({ 'type': 'Feature', 'properties': properties, 'geometry': { 'type': 'Point', 'coordinates': [float(row['lon']), float(row['lat'])] } })

jsondata = { 'type': 'FeatureCollection', 'features': features }
jsonfile = csvfile.with_suffix('.geojson')
with jsonfile.open('w') as file:
    json.dump(jsondata, file, ensure_ascii=False, indent=4)
