## csv2geojson.py

This script converts a csv file into a geojson file, e.g.:
```bash
python csv2geojson.py kienzler_stuttgart.csv
```
will generate the file `kienzler_stuttgart.geojson`.


## xlsx2geojson.py

This script converts ParkAPI Reference Table in Excel into a geojson file, e.g.:
```bash
python xlsx2geojson.py ulm_sensors parking-sites
```
will generate the file `ulm_sensors.geojson` for `ParkingSites` in the folder `parking-sites`.

```bash
python xlsx2geojson.py ulm_sensors parking-spots
```
will generate the file `ulm_sensors.geojson` for `ParkingSpots` in the folder `parking-spots`.
