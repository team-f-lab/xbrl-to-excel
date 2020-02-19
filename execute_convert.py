import os
from xbrl2excel.convert import convert_xbrl


if __name__ == "__main__":
    output_path = "converted.xlsx"
    doc_root = os.path.join(os.path.dirname(__file__), "./data")
    template_path = os.path.join(os.path.dirname(__file__), "./xbrl2excel/template.xlsx")
    convert_xbrl(
        output_path=output_path,
        doc_id="S100DE5C",
        doc_path=doc_root,
        template_path=template_path)
