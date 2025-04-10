import argparse
import json
import sys
from pathlib import Path

from datetime import datetime, timezone
from typing import Optional, Any

from openpyxl.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.reader.excel import load_workbook
from validataclass.exceptions import ValidationError

from parkapi_sources.exceptions import (
    ImportParkingSiteException,
    ImportParkingSpotException,
)
from parkapi_sources.models import StaticParkingSiteInput, StaticParkingSpotInput

from parkapi_sources.converters.base_converter import (
    ParkingSiteBaseConverter,
    ParkingSpotBaseConverter,
)
from parkapi_sources.converters.base_converter.push import (
    NormalizedXlsxConverter,
    XlsxConverter,
)
from parkapi_sources.models import SourceInfo, ExcelOpeningTimeInput, GeojsonInput
from parkapi_sources.util import ConfigHelper, RequestHelper
from decimal import Decimal
from validataclass.dataclasses import Default, validataclass
from validataclass.validators import DataclassValidator
from parkapi_sources.validators import (
    ExcelNoneable,
    GermanDurationIntegerValidator,
    NumberCastingStringValidator,
)

parser = argparse.ArgumentParser(
    prog="Xlsx2Geojson Conversion Script",
    description="This script helps to convert Excel Reference Table to ParkAPI Geojson",
)

parser.add_argument("source_uid")
parser.add_argument("source_group")
args = parser.parse_args()
source_uid: str = args.source_uid
source_group: str = args.source_group

if not source_group:
    sys.exit("Error: please add a source_uid as first argument")

if not source_group:
    sys.exit("Error: please add a source_group e.g. parking-spots, parking-sites")

file_path: Path = Path(f"sources/{str(source_group)}/{str(source_uid)}.xlsx")
if not file_path.is_file():
    sys.exit(f"Error: please add an Excel file with name '{file_path}'")

purpose_mapping: dict[str, str] = {
    "Auto": "CAR",
    "Fahrrad": "BIKE",
}


def serialize(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    elif isinstance(obj, datetime):
        return obj.isoformat()
    elif hasattr(obj, "value"):  # for Enums
        return obj.value
    elif isinstance(obj, list):
        return [
            serialize(item)
            if not isinstance(item, dict)
            else {k: serialize(v) for k, v in item.items() if v is not None}
            for item in obj
        ]
    elif obj.__class__.__name__ == "UnsetValueType":  # custom check for UnsetValue
        return None
    return str(obj)


def filter_none(data: dict) -> dict:
    return {key: serialize(value) for key, value in data.items() if value is not None}


class Xlsx2GeojsonParkingSites(NormalizedXlsxConverter, ParkingSiteBaseConverter):
    source_info = SourceInfo(
        uid=source_uid,
        name="Convert ParkingSites Reference Table to Geojson",
        has_realtime_data=True,
    )

    supervision_type_mapping: dict[str, str] = {
        True: "YES",
        False: "NO",
    }

    def __init__(self, config: Optional[dict] = None):
        self.config = config or {}
        self.config_helper = ConfigHelper(config=self.config)
        self.request_helper = RequestHelper(config_helper=self.config_helper)
        super().__init__(self.config_helper, self.request_helper)

        # For additional attributes not in Generic Converter and if column names changes
        additional_header_rows: dict[str, str] = {
            "Einfahrtshöhe (cm)": "max_height",
            "Zweck der Anlage": "purpose",
        }
        self.header_row = {
            **{
                key: value
                for key, value in super().header_row.items()
                if value not in additional_header_rows.values()
            },
            **additional_header_rows,
        }

    def map_row_to_parking_site_dict(
        self,
        mapping: dict[str, int],
        row: tuple[Cell, ...],
        **kwargs: Any,
    ) -> dict[str, Any]:
        parking_site_dict = super().map_row_to_parking_site_dict(mapping, row, **kwargs)

        for field in mapping.keys():
            parking_site_dict[field] = row[mapping[field]].value

        parking_site_dict["max_height"] = (
            round(parking_site_dict.get("max_height"))
            if isinstance(parking_site_dict.get("max_height"), float)
            else parking_site_dict.get("max_height")
        )
        parking_site_dict["max_stay"] = (
            round(parking_site_dict.get("max_stay"))
            if isinstance(parking_site_dict.get("max_stay"), float)
            else parking_site_dict.get("max_stay")
        )
        parking_site_dict["opening_hours"] = parking_site_dict["opening_hours"].replace(
            "00:00-00:00", "00:00-24:00"
        )
        parking_site_dict["purpose"] = purpose_mapping.get(
            parking_site_dict.get("purpose")
        )
        parking_site_dict["type"] = self.type_mapping.get(
            parking_site_dict.get("type"), "OFF_STREET_PARKING_GROUND"
        )
        parking_site_dict["supervision_type"] = self.supervision_type_mapping.get(
            parking_site_dict.get("supervision_type"),
        )
        parking_site_dict["static_data_updated_at"] = datetime.now(
            tz=timezone.utc
        ).isoformat()

        return parking_site_dict

    def handle_xlsx(
        self, workbook: Workbook
    ) -> tuple[list[StaticParkingSiteInput], list[ImportParkingSiteException]]:
        worksheet = workbook.active
        mapping: dict[str, int] = self.get_mapping_by_header(next(worksheet.rows))

        static_parking_site_errors: list[ImportParkingSiteException] = []
        static_parking_site_inputs: list[StaticParkingSiteInput] = []

        for row in worksheet.iter_rows(min_row=2):
            # ignore empty lines as LibreOffice sometimes adds empty rows at the end of a file
            if row[0].value is None:
                continue
            parking_site_dict = self.map_row_to_parking_site_dict(
                mapping=mapping,
                row=row,
                column_names=[cell.value for cell in next(worksheet.rows)],
            )

            try:
                static_parking_site_input = filter_none(
                    self.static_parking_site_validator.validate(
                        parking_site_dict
                    ).to_dict()
                )
                static_parking_site_inputs.append(
                    {
                        "type": "Feature",
                        "properties": static_parking_site_input,
                        "geometry": {
                            "type": "Point",
                            "coordinates": [
                                static_parking_site_input["lon"],
                                static_parking_site_input["lat"],
                            ],
                        },
                    }
                )

            except ValidationError as e:
                static_parking_site_errors.append(
                    ImportParkingSiteException(
                        source_uid=self.source_info.uid,
                        parking_site_uid=parking_site_dict.get("uid"),
                        message=f"invalid static parking site data {parking_site_dict}: {e.to_dict()}",
                    )
                )
                continue

        return static_parking_site_inputs, static_parking_site_errors


@validataclass
class ExcelStaticParkingSpotInput(StaticParkingSpotInput):
    uid: str = NumberCastingStringValidator(min_length=1, max_length=256)
    max_stay: Optional[int] = (
        ExcelNoneable(GermanDurationIntegerValidator()),
        Default(None),
    )


class Xlsx2GeojsonParkingSpots(XlsxConverter, ParkingSpotBaseConverter):
    source_info = SourceInfo(
        uid=source_uid,
        name="Convert ParkingSpots Reference Table to Geojson",
        has_realtime_data=True,
    )

    static_parking_spot_validator = DataclassValidator(ExcelStaticParkingSpotInput)
    excel_opening_time_validator = DataclassValidator(ExcelOpeningTimeInput)

    header_row: dict[str, str] = {
        "ID": "uid",
        "Name": "name",
        "Widmung": "type",
        "Längengrad": "lon",
        "Breitengrad": "lat",
        "Zweck der Anlage": "purpose",
        "Geometry": "geojson",
        "Maximale Parkdauer": "max_stay",
        "24/7 geöffnet?": "opening_hours_is_24_7",
        "Öffnungszeiten Mo-Fr Beginn": "opening_hours_weekday_begin",
        "Öffnungszeiten Mo-Fr Ende": "opening_hours_weekday_end",
        "Öffnungszeiten Sa Beginn": "opening_hours_saturday_begin",
        "Öffnungszeiten Sa Ende": "opening_hours_saturday_end",
        "Öffnungszeiten So Beginn": "opening_hours_sunday_begin",
        "Öffnungszeiten So Ende": "opening_hours_sunday_end",
    }

    restricted_to_type_mapping: dict[str, str] = {
        "Ladesäule": "CHARGING",
        "Familie": "FAMILY",
        "Handicap": "DISABLED",
    }

    def __init__(self, config: Optional[dict] = None):
        self.config = config or {}
        self.config_helper = ConfigHelper(config=self.config)
        self.request_helper = RequestHelper(config_helper=self.config_helper)
        super().__init__(self.config_helper, self.request_helper)

    def handle_xlsx(
        self, workbook: Workbook
    ) -> tuple[list[StaticParkingSpotInput], list[ImportParkingSpotException]]:
        worksheet = workbook.active
        mapping: dict[str, int] = self.get_mapping_by_header(next(worksheet.rows))

        static_parking_spot_errors: list[ImportParkingSpotException] = []
        static_parking_spot_features: list[StaticParkingSpotInput] = []

        for row in worksheet.iter_rows(min_row=2):
            # ignore empty lines as LibreOffice sometimes adds empty rows at the end of a file
            if row[0].value is None:
                continue
            parking_spot_dict = self.map_row_to_parking_spot_dict(
                mapping=mapping,
                row=row,
                column_names=[cell.value for cell in next(worksheet.rows)],
            )

            try:
                static_parking_spot_input = filter_none(
                    self.static_parking_spot_validator.validate(
                        parking_spot_dict
                    ).to_dict()
                )
                static_parking_spot_features.append(
                    {
                        "type": "Feature",
                        "properties": static_parking_spot_input,
                        "geometry": {
                            "type": "Point",
                            "coordinates": [
                                static_parking_spot_input["lon"],
                                static_parking_spot_input["lat"],
                            ],
                        },
                    }
                )

            except ValidationError as e:
                static_parking_spot_errors.append(
                    ImportParkingSpotException(
                        source_uid=self.source_info.uid,
                        parking_spot_uid=parking_spot_dict.get("uid"),
                        message=f"invalid static parking spot data {parking_spot_dict}: {e.to_dict()}",
                    )
                )
                continue

        return static_parking_spot_features, static_parking_spot_errors

    def map_row_to_parking_spot_dict(
        self,
        mapping: dict[str, int],
        row: tuple[Cell, ...],
        column_names: list[str],
    ) -> dict[str, Any]:
        parking_spot_raw_dict: dict[str, str] = {}
        for field in mapping.keys():
            parking_spot_raw_dict[field] = row[mapping[field]].value

        parking_spot_dict = {
            key: value
            for key, value in parking_spot_raw_dict.items()
            if not key.startswith("opening_hours_")
        }
        opening_hours_input = self.excel_opening_time_validator.validate(
            {
                key: value
                for key, value in parking_spot_raw_dict.items()
                if key.startswith("opening_hours_")
            }
        )

        restricted_to_type = (
            self.restricted_to_type_mapping.get(
                parking_spot_raw_dict.get("type", "").strip()
            )
            if isinstance(parking_spot_raw_dict["type"], str)
            else None
        )
        restricted_to = {
            "type": restricted_to_type,
            "hours": opening_hours_input.get_osm_opening_hours(),
        }
        parking_spot_dict["restricted_to"] = [restricted_to]
        parking_spot_dict["has_realtime_data"] = self.source_info.has_realtime_data
        parking_spot_dict["purpose"] = purpose_mapping.get(
            parking_spot_dict.get("purpose")
        )
        parking_spot_dict["static_data_updated_at"] = datetime.now(
            tz=timezone.utc
        ).isoformat()

        return parking_spot_dict


workbook = load_workbook(filename=str(file_path.absolute()))
static_parking_inputs, import_parking_exceptions = [], []

if source_group == "parking-sites":
    xlsx2geojson = Xlsx2GeojsonParkingSites()
elif source_group == "parking-spots":
    xlsx2geojson = Xlsx2GeojsonParkingSpots()

static_parking_inputs, import_parking_exceptions = xlsx2geojson.handle_xlsx(workbook)
static_geojson_parking_inputs: GeojsonInput = {
    "type": "FeatureCollection",
    "features": static_parking_inputs,
}

geojson_file = file_path.with_suffix(".geojson")
with geojson_file.open("w") as file:
    json.dump(static_geojson_parking_inputs, file, ensure_ascii=False, indent=4)

print(
    f"Successful with {len(static_parking_inputs)} {source_group} and {len(import_parking_exceptions)} Errors"
)
