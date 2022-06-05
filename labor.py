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
    'Employee Name',
    'Job Name',
    'Task Name',
    'Total Hours',
    'Fab',
    'Paint',
    'Canvas',
    'Floor Boards',
    'Outfitting',
]

FORMATS = [
    'General',
    'General',
    'General',
    '0.00',
    '0.00',
    '0.00',
    '0.00',
    '0.00',
    '0.00',
]

WIDTHS = [
    26,
    11,
    24,
    12,
    12,
    12,
    12,
    12,
    12,
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

def build_time(rows, hulls):
    """build hulls/employees/hours for boats with last month activity"""
    boats = {'total': Decimal(0)}
    for row in rows:
      hull = row[2]
      dept = row[11]
      employee = row[4] + ', ' + row[5]
      if row[8]:
          punch = Decimal(row[8])
      else:
          punch = Decimal(0)
      type = row[10]
      if hull not in hulls:
          continue
      if hull not in boats:
          boats[hull] = {'total': Decimal(0)}
      if dept not in boats[hull]:
          boats[hull][dept] = {'total': Decimal(0)}
      if employee not in boats[hull][dept]:
          boats[hull][dept][employee] = Decimal(0)
      if punch == 1:
          boats[hull][dept][employee] -= punch
          boats[hull][dept]['total'] -= punch
          boats[hull]['total'] -= punch
          boats['total'] -= punch
      else:
          boats[hull][dept][employee] += punch
          boats[hull][dept]['total'] += punch
          boats[hull]['total'] += punch
          boats['total'] += punch
  
    return boats

def write_sheet(boats):
    """write sheet to disk"""
    excel = ExcelOpenDocument()
    excel.new('test.xlsx')
    title_font = excel.font(name='Calibri', size=11, bold=True)
    body_font = excel.font(name='Calibri', size=11)

    # set column widths
    for column, width in enumerate(WIDTHS, start=65):
         excel.set_width(chr(column), width)

    # write column names
    for column, field, format, width in zip(range(len(FIELDS)), FIELDS, FORMATS, WIDTHS):
        excel.cell(row=1, column=column+1).value = field
        excel.cell(row=1, column=column+1).font = title_font

    row = 2
    for boat in sorted(boats):
        if boat != 'total':
            dept_row = row - 1
            for column, dept in [(5, '1 Boat Builder'),
                                 (7, '2 Canvas and Upholstery'),
                                 (6, '4 Paint'),
                                 (8, '5 Outfitting - Floorboard'),
                                 (9, '5 Outfitting')]:
                if dept in boats[boat]:
                    old_row = dept_row + 1
                    dept_row += len(boats[boat][dept]) - 1
                    # excel.cell(row=dept_row, column=column).value = boats[boat][dept]['total']
                    excel.cell(row=dept_row, column=column).value = f"=SUM(D{old_row}:D{dept_row})"
                    excel.cell(row=dept_row, column=column).font = body_font
                    excel.cell(row=dept_row, column=column).number_format = r'0.00'
            for dept in sorted(boats[boat]):
                if dept != 'total':
                    for employee in sorted(boats[boat][dept]):
                        if employee != 'total':
                            excel.cell(row=row, column=1).value = employee
                            excel.cell(row=row, column=1).font = body_font
                            excel.cell(row=row, column=2).value = boat
                            excel.cell(row=row, column=2).font = body_font
                            excel.cell(row=row, column=3).value = dept
                            excel.cell(row=row, column=3).font = body_font
                            excel.cell(row=row, column=4).value = boats[boat][dept][employee]
                            excel.cell(row=row, column=4).font = body_font
                            excel.cell(row=row, column=4).number_format = r'0.00'
                            row += 1
            row += 1

    excel.cell(row=row, column=1).value = 'Totals'
    excel.cell(row=row, column=1).font = title_font

    excel.cell(row=row, column=4).value = f"=SUM(D2:D{row-2})"
    excel.cell(row=row, column=4).font = title_font
    excel.cell(row=row, column=4).number_format = r'0.00'

    excel.cell(row=row, column=5).value = f"=SUM(E2:E{row-2})"
    excel.cell(row=row, column=5).font = title_font
    excel.cell(row=row, column=5).number_format = r'0.00'
    
    excel.cell(row=row, column=6).value = f"=SUM(F2:F{row-2})"
    excel.cell(row=row, column=6).font = title_font
    excel.cell(row=row, column=6).number_format = r'0.00'
    
    excel.cell(row=row, column=7).value = f"=SUM(G2:G{row-2})"
    excel.cell(row=row, column=7).font = title_font
    excel.cell(row=row, column=7).number_format = r'0.00'
    
    excel.cell(row=row, column=8).value = f"=SUM(H2:H{row-2})"
    excel.cell(row=row, column=8).font = title_font
    excel.cell(row=row, column=8).number_format = r'0.00'
    
    excel.cell(row=row, column=9).value = f"=SUM(I2:I{row-2})"
    excel.cell(row=row, column=9).font = title_font
    excel.cell(row=row, column=9).number_format = r'0.00'

    excel.cell(row=row, column=10).value = f"=SUM(E{row}:I{row})"
    excel.cell(row=row, column=10).font = title_font
    excel.cell(row=row, column=10).number_format = r'0.00'

    excel.save()


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
    start = '2021-03-01 00:00:00'
    finish = '2022-03-31 23:59:59'
    month = 3
    rows = get_boats(host, database, user, password, start, finish)
    hulls = set([row[2] for row in rows if row[6].month == month
                                       and row[6].year == year
                                       and row[2]
                                       and row[2][0] < '6'])
    hulls = sorted(get_hulls(rows, month, year))
    boats = build_time(rows, hulls)
    write_sheet(boats)
    _ = (path)
    
    sys.exit(0)


if __name__ == "__main__":
    cli()  # pylint: disable=no-value-for-parameter
