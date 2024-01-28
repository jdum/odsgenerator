# Copyright 2021-2024 Jérôme Dumonteil
# Licence: MIT
# Authors: jerome.dumonteil@gmail.com
"""CLI interface to odsgenerator.
"""

import argparse
import sys

import odfdo

from odsgenerator.odsgenerator import __doc__, __version__, file_to_ods

ODFDO_REQUIREMENT = (3, 5, 0)


def check_odfdo_version():
    """Utility to verify we have the minimal version of the odfdo library."""
    if tuple(int(x) for x in odfdo.__version__.split(".")) >= ODFDO_REQUIREMENT:
        return True
    print(  # pragma: no cover
        "Error: odfdo version >= "
        f"{'.'.join(str(x) for x in ODFDO_REQUIREMENT)} is required"
    )
    return False  # pragma: no cover


def main():  # pragma: no cover
    """Read parameters from STDIN and apply the required command.

    Usage:
       odsgenerator [-h] [--version] input_file output_file

    Arguments:
        input_file: Input file containing data in json or yaml format.

        output_file: Output file, .ods file generated from input.

    Use `odsgenerator --help` for more details about input file parameters
    and look at examples in the tests folder.
    """
    if not check_odfdo_version():
        sys.exit(1)
    parser = argparse.ArgumentParser(
        description="odsgenerator, an .ods generator.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--version", action="version", version="%(prog)s " + __version__
    )
    parser.add_argument(
        "input_file", help="input file containing data in json or yaml format"
    )
    parser.add_argument(
        "output_file", help="output file, .ods file generated from input"
    )
    args = parser.parse_args()
    file_to_ods(args.input_file, args.output_file)


if __name__ == "__main__":
    main()  # pragma: no cover
