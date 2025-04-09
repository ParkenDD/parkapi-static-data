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

from parkapi_sources.exceptions import ImportParkingSiteException
from parkapi_sources.models import StaticParkingSiteInput

from parkapi_sources.converters.base_converter import ParkingSiteBaseConverter
from parkapi_sources.converters.base_converter.push import NormalizedXlsxConverter
from parkapi_sources.models import SourceInfo
from parkapi_sources.models import GeojsonInput
from parkapi_sources.util import ConfigHelper, RequestHelper
from decimal import Decimal

parser = argparse.ArgumentParser(
    prog="Xlsx2Geojson Conversion Script",
    description="This script helps to convert Excel Reference Table to ParkAPI Geojson",
)

parser.add_argument("source_uid")
args = parser.parse_args()
source_uid: str = args.source_uid
file_path: Path = Path(f"{str(source_uid)}.xlsx")

if not file_path.is_file():
    sys.exit(f"Error: please add an Excel file with name '{file_path}'")


class Xlsx2GeojsonConverter(NormalizedXlsxConverter, ParkingSiteBaseConverter):
    config: Optional[dict] = (None,)
    config_helper = ConfigHelper(config=config)
    request_helper = RequestHelper(config_helper=config_helper)
    source_info = SourceInfo(
        uid=source_uid,
        name="Convert Excel Reference Table to Geojson",
        has_realtime_data=False,
    )

    purpose_mapping: dict[str, str] = {
        "Auto": "CAR",
        "Fahrrad": "BIKE",
    }

    supervision_type_mapping: dict[str, str] = {
        True: "YES",
        False: "NO",
    }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, self.config_helper, self.request_helper, **kwargs)
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
        parking_site_dict["purpose"] = self.purpose_mapping.get(
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
                static_parking_site_input = self.filter_none(
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

    def serialize(self, obj):
        if isinstance(obj, Decimal):
            return float(obj)
        elif isinstance(obj, datetime):
            return obj.isoformat()
        elif hasattr(obj, "value"):  # for Enums
            return obj.value
        elif obj.__class__.__name__ == "UnsetValueType":  # custom check for UnsetValue
            return None
        return str(obj)

    def filter_none(self, data: dict) -> dict:
        return {
            key: self.serialize(value)
            for key, value in data.items()
            if value is not None
        }


xlsx2geojson = Xlsx2GeojsonConverter()
workbook = load_workbook(filename=str(file_path.absolute()))
static_parking_site_inputs, import_parking_site_exceptions = xlsx2geojson.handle_xlsx(
    workbook
)
static_geojson_parking_site_inputs: GeojsonInput = {
    "type": "FeatureCollection",
    "features": static_parking_site_inputs,
}

xlsx2geojson = Xlsx2GeojsonConverter()
geojson_file = file_path.with_suffix(".geojson")

with geojson_file.open("w") as file:
    json.dump(static_geojson_parking_site_inputs, file, ensure_ascii=False, indent=4)
print(
    f"Successful with {len(static_parking_site_inputs)} ParkingSites and {len(import_parking_site_exceptions)} Errors"
)
