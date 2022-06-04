#!/usr/bin/env python
"""
create monthly labor summary

notes: build list of boats with punches in the last month
       build dict of boats that have punches
       those boats have a dictionary of employeees
       each employee has a total number of hours
"""

import os
import sys
from decimal import Decimal
from datetime import datetime
from datetime import timedelta
from platform import system
import click
import pytds
from dotenv import load_dotenv
from excelopen import ExcelOpenDocument

load_dotenv()  # use os.getenv()


FIELDS = [
]

FORMATS = [
]

WIDTHS = [
]

SQL = """
   SELECT  substring(departmentname, 1,3 ) as departmentname, tp.job_id,
           CASE WHEN job.jobName IS NULL THEN '' ELSE job.jobname END as jobName,
           tp.employee_id, em.lastname, em.firstname,
           tp.inpunch_dt, tp.workingpunch_ts,
           CASE
             WHEN DATEPART(MINUTE, workingpunch_ts) = 45 THEN DATEPART(HOUR, workingpunch_ts) + .75
             WHEN DATEPART(MINUTE, workingpunch_ts) = 30 THEN DATEPART(HOUR, workingpunch_ts) + .5
             WHEN DATEPART(MINUTE, workingpunch_ts) = 15 THEN DATEPART(HOUR, workingpunch_ts) + .25
             WHEN DATEPART(MINUTE, workingpunch_ts) =  0 THEN DATEPART(HOUR, workingpunch_ts)
           END as WorkTime, tp.workingpunch_id,
           tp.inout_id, task.taskname, tp.task_id
     FROM  timeWorkingPunch tp
LEFT JOIN  job on tp.job_id = job.job_id
LEFT JOIN  task on tp.task_id = task.task_id
     JOIN  empMain em ON tp.employee_id = em.employee_id
     JOIN  tblDepartment dp ON tp.department_id  = dp.department_id
    WHERE  tp.inpunch_dt BETWEEN '%s' AND '%s'
      AND  tp.active_yn = 1
      AND  task.taskname IN ('1 Boat Builder', '2 Canvas and Upholstery', '4 Paint', '5 Outfitting', '5 Outfitting - Floorboard')
  --  AND  job.jobname IN ('18056 122')
"""

def format_dates():
    """return formated dates for prior month"""
    today = datetime.now()
    last = datetime(today.year, today.month, 1)-timedelta(days=1)
    first = datetime(last.year, last.month, 1)
    start = first - timedelta(days=365)
    return (first.strftime('%B %Y'),
	    start.strftime('%Y-%m-%d 00:00:00'),
      last.strftime('%Y-%m-%d 23:59:59'),
      last.month,
      last.year)


def get_boats(host, database, user, password, start, finish):
    """placeholder for pytds template"""
    with pytds.connect(server=host,
                       database=database,
                       user=user,
                       password=password,
                       port=1433,
                       tds_version=0x70000000,
                      ) as conn:
        with conn.cursor() as cur:
            _ = cur.execute(SQL % (start, finish))
            rows = cur.fetchall()
    return rows

def get_hulls(rows, month, year):
    """get a list of boats worked on during the month"""
    return set([row[2] 
                for row in rows 
                if row[6].month == month and
                    row[6].year == year and
                    row[2] and
                    row[2][0] < '6'])


@click.command()
@click.option('--host',
              '-h',
              envvar='DBHOST',
              default='',
              help='host to connect to'
              )
@click.option('--database',
              '-d',
              envvar='DATABASE',
              default='',
              help='database to use'
              )
@click.option('--user',
              '-u',
              envvar='SQLUSER',
              default='',
              help='database user'
              )
@click.option('--password',
              '-p',
              envvar='SQLPASSWORD',
              default='',
              help='database pasword'
              )
@click.option('--path',
              '-p',
              envvar='PATH_DIR',
              default='',
              help='Path to save file to'
              )
def cli(host, database, user, password, path):
    """Create spreadsheet with inventory items from fishbowl
    You will want to use: -e Upholstry -e Shipping -e Apparel
    """
    period, start, finish, month, year = format_dates()
    rows = get_boats(host, database, user, password, start, finish)
    hulls = set([row[2] for row in rows if row[6].month == month
                                       and row[6].year == year
                                       and row[2]
                                       and row[2][0] < '6'])
    hulls = sorted(get_hulls(rows, month, year))
    print(hulls)
    _ = (path)
    sys.exit(0)



if __name__ == "__main__":
    cli()  # pylint: disable=no-value-for-parameter
