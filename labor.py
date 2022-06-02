#!/usr/bin/env python
"""
create monthly labor summary
"""

import os
from decimal import Decimal
from datetime import datetime
from datetime import timedelta
from platform import system
import fdb
import click
from dotenv import load_dotenv
from excelopen import ExcelOpenDocument

load_dotenv()  # use os.getenv()


FIELDS = [
]

FORMATS = [
]

WIDTHS = [
]



@click.command()
@click.option('--host',
              '-h',
              envvar='PRODUCTIONHOST',
              default='',
              help='host to connect to'
              )
@click.option('--path',
              '-p',
              envvar='PATH_DIR',
              default='',
              help='Path to save file to'
              )
def cli(host, path):
    """Create spreadsheet with inventory items from fishbowl
    You will want to use: -e Upholstry -e Shipping -e Apparel
    """
    # rows = read_firebird_database(host)
    print(path)


if __name__ == "__main__":
        cli()  # pylint: disable=no-value-for-parameter
