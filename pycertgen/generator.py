from typing import Literal
from . import ppt_to_pdf, ppt_manager
import pandas as pd
import os


def gen(certified_list_path: str, template_path: str, output_dir: str, key_autogen=False, to_pdf_method: Literal['libreoffice', 'powerpoint'] = None):
    certified_df = pd.read_excel(certified_list_path)

    certified_data = certified_df.reset_index().to_dict(orient='records')

    for certified in certified_data:

        # Call the function to replace text in the PowerPoint presentation
        path = ppt_manager.fill_template(
            template_path, certified, output_dir, key_autogen)

        if to_pdf_method:
            ext = os.path.split(path)[-1]
            output_pdf = path.replace(ext, 'pdf')
            ppt_to_pdf.convert(path, output_pdf, to_pdf_method)
