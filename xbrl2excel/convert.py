import os
import shutil
import json
from datetime import datetime
from openpyxl import load_workbook
import xbrr


def convert_xbrl(output_path, doc_id, doc_path, template_path):
    print(f"Download XBRL file.")
    xbrl_path = xbrr.edinet.api.document.get_xbrl(
                   doc_id, save_dir=doc_path, expand_level="dir")

    xbrl = xbrr.edinet.reader.read(xbrl_path)

    print(f"Read template file.")
    book = load_workbook(template_path)

    print(f"Transfer value from XBRL to EXCEL.")
    target = book["summary"]

    name = xbrl.extract(xbrr.edinet.aspects.Metadata).company_name.value
    period_end = xbrl.extract(xbrr.edinet.aspects.Metadata).fiscal_year_end_date.value
    fiscal_year = xbrl.extract(xbrr.edinet.aspects.Metadata).fiscal_year.value
    target["A1"] = name

    date_ranges = {}
    current_datetime = period_end
    for c in "GFEDC":
        current_date = current_datetime.strftime("%Y-%m-%d")
        date_ranges[current_date] = c
        next_datetime = current_datetime.replace(year=current_datetime.year - 1)
        current_datetime = next_datetime

    print(date_ranges)
    target["C3"] = datetime.strptime(min(list(date_ranges.keys())), "%Y-%m-%d")
    tag_column = "C"
    tag_start = 4
    num_tags = 1

    for i in range(num_tags):
        index = tag_start + i
        cell = f"{tag_column}{index}"
        cell_value = target[cell].value

        if not cell_value or not cell_value.startswith("jp"):
            continue

        tag = cell_value.strip()
        elements = xbrl.find_all(tag)
        if len(elements) == 0:
            target[cell] = None
        else:
            for e in elements:
                v = e.value(label_kind=None).to_dict()
                _v = None
                if v["value"]:
                    scale = 1 if not v["decimals"] else int(v["decimals"])
                    if scale < 0:
                        scale = 10 ** (-scale)
                    else:
                        scale = 1
                    _v = float(v["value"]) / scale

                if v["period"] in date_ranges:
                    target[f"{date_ranges[v['period']]}{index}"] = _v

    file_name = f"{name}_{fiscal_year}年度.xlsx"
    book.save(output_path)
    shutil.rmtree(xbrl_path)
    return output_path
