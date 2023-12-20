#! python3

import argparse
from . import gui, generator


def parse_arguments():
    parser = argparse.ArgumentParser(
        description='Generate PDF certificates from command line.')

    parser.add_argument('--show_gui', help='Show GUI',
                        action='store_true')
    parser.add_argument('--certified_path', type=str,
                        help='Path to the certified list excel file', metavar='f')
    parser.add_argument('--template_path', type=str,
                        help='Path to the PowerPoint template file', metavar='t')
    parser.add_argument('--output_path', type=str,
                        help='Path to the output directory', metavar='o')
    parser.add_argument('--key_autogen', action='store_true', default=False,
                        help='Enable automatic key generation')
    parser.add_argument('--to_pdf_method', type=str, choices=['libreoffice', 'powerpoint'],
                        default='libreoffice', help='Method to convert ppt to pdf', metavar='m')

    args = parser.parse_args()

    if not args.show_gui:
        if not args.certified_path or not args.template_path or not args.output_path:
            parser.error(
                'when --show_gui is false, --certified_path, --template_path and --output_path are required')

    return args


if __name__ == '__main__':
    args = parse_arguments()

    print(args)
    if args.show_gui:
        gui.run()
    else:
        generator.gen(args.certified_path, args.template_path, args.output_path,
                      args.key_autogen, args.to_pdf_method)
