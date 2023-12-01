"""
This script parses CVF or OpenView paper abstracts to a docx file.
"""

import argparse
from utils import lazy_export, parse


def parse_args():
    parser = argparse.ArgumentParser(description='parse CVF or OpenReview paper abstracts to a docx file')
    parser.add_argument('--conference', help='exact abbreviation of conference, '
                                             'support NeurIPS, ICLR, ICML, CVPR, and ICCV so far')
    parser.add_argument('--year', help='year of conference')
    parser.add_argument('--toc', action='store_true', help='add table of contents at the very beginning, '
                                                           'ONLY support Windows system installed MS Word, '
                                                           'because of the `win32com` package')
    parser.add_argument('--pdf', action='store_true', help='additionally generate pdf, '
                                                           'ONLY support Windows system installed MS Word, '
                                                           'because of the `win32com` package')
    parser.add_argument('--wps', action='store_true', help='use WPS to convert pdf, which performs better than MS Word '
                                                           'and can handle large file size')
    parser.add_argument('--select', action='store_true', help='interactively select specific accepted types, '
                                                              'for OpenReview papers')
    parser.add_argument('--all', action='store_true', help='lazy export paper abstracts of all conferences in all '
                                                           'supported years')

    args = parser.parse_args()
    return args


def main():
    args = parse_args()
    lazy_export(args) if args.all else parse(args)


if __name__ == '__main__':
    main()
