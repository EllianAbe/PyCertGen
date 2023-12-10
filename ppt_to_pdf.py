from typing import Literal
import subprocess
import os


def convert(input_ppt, output_dir, method: Literal['libreoffice', 'powerpoint'] = 'libreoffice'):
    try:
        if method == 'libreoffice':
            subprocess.run(['libreoffice', '--headless', '--invisible',
                            '--convert-to', 'pdf', '--outdir',
                            os.path.dirname(output_dir), input_ppt],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        if method == 'powerpoint':
            subprocess.run(['powerpnt', '--convert-to', 'pdf', input_ppt,])
            raise NotImplementedError("powerpoint method not implemented yet.")

    except subprocess.CalledProcessError as e:
        print(f"Conversion failed: {e}")
