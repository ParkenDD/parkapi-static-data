import sys
import csv
import json

argv = sys.argv[1:]
if len(argv) == 0:
    print("no csv file specified")
    sys.exit(2)
csvfile = argv[0]

features = []
with open(csvfile, 'r', encoding='utf-8') as file:
    reader = csv.DictReader(file)
    for row in reader:
        if 'uid' not in row or row['uid'] == '' :
            print(f'Attribute \'uid\' missing in row {row}')
            continue
        if 'lat' not in row or 'lon' not in row:
            print(f'Attributes \'lat\'/\'lon\' missing in row {row}')
            continue
        properties = { 'uid': row['uid'] }
        for column in ['address', 'type']:
            if column in row and row[column] != '':
                properties[column] = row[column]
        for column in ['max_height', 'max_width', 'max_depth']:
            if column in row and row[column] != '' and row[column].isdigit():
                properties[column] = int(row[column])
        if 'park_and_ride_type' in row and row['park_and_ride_type'] != '':
            properties['park_and_ride_type'] = [row['park_and_ride_type']]
        if 'DHID' in row and row['DHID'] != '':
            properties['external_identifiers'] = { 'type': 'DHID', 'value': row['DHID'] }
        features.append({ 'type': 'Feature', 'properties': properties, 'geometry': { 'type': 'Point', 'coordinates': [float(row['lon']), float(row['lat'])] } })

jsonfile = csvfile.replace('.csv', '.geojson')
jsondata = { 'type': 'FeatureCollection', 'features': features }
with open(jsonfile, 'w', encoding='utf-8') as file:
    json.dump(jsondata, file, ensure_ascii=False, indent=4)