"""
    Released under MIT License (MIT)

    Copyright (c) 2020 Vincent Audric

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
    OR OTHER DEALINGS IN THE SOFTWARE.
"""
__license__ = "MIT"
__copyright__ = "Copyright (c) 2020 Vincent Audric"
__author__ = "Vincent Audric"
__version__ = "0.1"
__status__ = "at_your_own_risk"


from pprint import pprint
import os
import time
import datetime
import hashlib, uuid
from urllib.parse import urlparse
import click
from dateutil.relativedelta import relativedelta
import requests
from bs4 import BeautifulSoup


# Start some date
today = datetime.date.today()
default_year = today.year
default_month = today.month - 1 if today.month >  1 else 12

WAIT_TIME = 0.5

@click.command()
@click.option('-u', '--empno', prompt='SWOL Employee Number', type=click.STRING, default='000000', help=f'Employee number (SWOL username) (e.g.: 000000).')
@click.option('-p', '--password', prompt='SWOL Password (encrypted and not stored)', type=click.STRING, confirmation_prompt=True, help=f'Skywest Online (SWOL) Password.', hide_input=True)
@click.option('--start-year', 'start_year', prompt='From year', type=click.INT, default=default_year, help=f'Export from this year. (e.g.: {default_year})')
@click.option('--start-month', 'start_month', prompt='From month', type=click.IntRange(1, 12), default=default_month, help=f'Export from this month. (e.g.: {default_month})')
@click.option('--end-year', 'end_year', prompt='To year', type=click.INT, default=default_year, help=f'Export to this year. (e.g.: {default_year})')
@click.option('--end-month', 'end_month', prompt='To month', type=click.IntRange(1, 12), default=default_month, help=f'Export to this month. (e.g.: {default_month})')
@click.option('-f', '--format', prompt='Block format', type=click.STRING, default='HHMM', help=f'Block time format: HHMM or Decimal.')
@click.option('-o', '--output', prompt='Ouput file name', type=click.STRING, default='export', help=f'Name of output file without extension. (e.g.: export)')
def export_csv(*argv, **kwargs):
    """\b
    Simple program that exports your block times since YEAR MONTH up to last month.
    By nice to SKW! Don't request more than 2 years (24 months) at a time.
    The program won't let you do it.
    Be Nice Throttle (BNT):The script sleeps for 1 second between every page load.
    """

    password = str(kwargs['password'])
    salt = uuid.uuid4().hex
    hashed_password = hashlib.sha512((password + salt).encode('utf-8')).hexdigest()[:15]

    # Only accept 'HHMM' or 'Decimal' formats
    if kwargs['format'] not in ['HHMM', 'Decimal']:
        kwargs['format'] = 'HHMM'

    # click.echo('Password is encrypted and not stored: {}'.format(click.style(hashed_password, fg='blue')))

    # Bid Months
    start = datetime.date(kwargs['start_year'], kwargs['start_month'], 1)
    end = datetime.date(kwargs['end_year'], kwargs['end_month'], 1)

    from_bid_month = str(kwargs['start_year']).zfill(4) + str(kwargs['start_month']).zfill(2)
    to_bid_month = str(kwargs['end_year']).zfill(4) + str(kwargs['end_month']).zfill(2)

    if start > end:
        click.secho('The from BidMonth cannot be after the to BidMonth', fg='red')
        click.echo('From BidMonth: {}'.format(click.style(from_bid_month, fg='red', bold=True)), err=True)
        click.echo('To BidMonth: {}'.format(click.style(to_bid_month, fg='red', bold=True)), err=True)
        return False
    else:
        click.echo('From BidMonth: {}'.format(click.style(from_bid_month, fg='green', bold=True)))
        click.echo('To BidMonth: {}'.format(click.style(to_bid_month, fg='green', bold=True)))
        kwargs.update({
            'start': start,
            'end': end,
        })

    # Pad Employee Number
    kwargs['empno'] = str(kwargs['empno']).zfill(6)

    swol_export_csv(**kwargs)


def swol_export_csv(**kwargs):
    start = kwargs['start']
    end = kwargs['end']
    bid_months = []
    while start <= end:
        bid_month = str(start.year).zfill(4) + str(start.month).zfill(2)
        bid_months.append(bid_month)
        start += relativedelta(months=1)

    if len(bid_months) > 24:
        click.echo('You are unreasonable asking for the last {} months'.format(click.style(str(len(bid_months)), fg='red', bold=True)))
        bid_months = bid_months[-24:]
        click.echo('The new start BidMonth is: {}'.format(click.style(bid_months[0], fg='green')))

    # HTTP Request Headers - Make it look you're human
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en-US,en;q=0.9,fr;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
    }

    swol_url = 'https://www.skywestonline.com'
    swol_my_sked_uri = '/SKYW/SkedPlus/MySchedule.aspx'
    swol_my_sked_params = {'Notify' : 'N'}

    # GET Login Page
    s = requests.Session()
    s.headers = headers
    url = swol_url + swol_my_sked_uri

    try:
        r = s.get(url, params=swol_my_sked_params)
    except Exception as e:
        click.secho("Problem loading the login page. Retry later", err=True, color='red', fg='red')
        return False


    if not r.ok:
        click.secho("Problem loading the login page. Retry later", err=True, color='red', fg='red')
        return False

    # Process Login Form
    bs = BeautifulSoup(r.text, 'html5lib')
    form = bs.find('form')

    # Fetch all input form fields to build the POSTed params
    params = {}
    inputs = bs.form.findAll('input')
    for input in inputs:
        if 'value' in input.attrs:
            params.update({input.attrs['name']: input.attrs['value']})
        else:
            params.update({input.attrs['name']: ''})

    # Authentication
    login = {
        'username': ('ctl02$txtEmpNo', kwargs['empno']),
        'password': ('ctl02$txtPassword', kwargs['password']),
    }

    #  Replace with the provided username and password
    for k, v in login.items():
        params.update({v[0]: v[1]})

    # We wait WAIT_TIME seconds
    time.sleep(WAIT_TIME)

    # POST Login
    s.headers.update({'Referer': r.url})
    o_url = urlparse(r.url)
    url = swol_url + '/'.join(o_url.path.split('/')[:-1]) + '/' + form.attrs['action']
    try:
        r = s.post(url, params)
    except Exception as e:
        click.secho("Problem loading SkedPlus page. Retry later", err=True, color='red', fg='red')
        return False

    if not r.ok:
        click.secho("Problem loading SkedPlus page. Retry later", err=True, color='red', fg='red')
        return False

    # Check that we landed on the SkedPlus page
    if 'Login.aspx' in r.url:
        bs = BeautifulSoup(r.text, 'html5lib')
        click.secho("Problem SWOL Login. Retry later", err=True, color='red', fg='red')
        div_err = bs.find('div', {'id': 'ctl02_divErr'})
        if div_err is not None:
            click.secho("Invalid Username or Password", err=True, color='red', fg='red')
        return False

    # GET Export CSV Files
    s.headers.update({'Referer': r.url})

    # Loop through all the bid months
    counter = 0
    csv_header = ''
    filepath = kwargs['output'] + '_' + kwargs['format'] + '.csv'
    f = open(filepath, 'w')
    for bid_month in bid_months:
        # We wait WAIT_TIME seconds
        time.sleep(WAIT_TIME)

        params = {
            'Format': 'CSV',
            'Block': kwargs['format'],
            'BidMonth': bid_month,
        }

        swol_export_csv_uri = '/SKYW/SkedPlus/Export.aspx'
        url = swol_url + swol_export_csv_uri
        try:
            r = s.get(url, params=params)
        except Exception as e:
            click.secho("Problem loading the CSV export page. Retry later", err=True, color='red', fg='red')
            return False

        if not r.ok:
            click.secho("Problem loading the CSV export page. Retry later", err=True, color='red', fg='red')
            return False

        lines = r.text.splitlines()

        # Write CSV headers only once
        if counter == 0:
            f.write(lines[0].strip() + os.linesep)

        if len(lines) > 1:
            for line in lines[1:]:
                f.write(line.strip() + os.linesep)

        counter += 1

    # Close file
    f.close()

    # Close Session
    s.close()


if __name__ == '__main__':
    export_csv()
