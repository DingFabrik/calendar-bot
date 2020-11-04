#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import print_function

import configparser
import locale
import os.path
import pickle
import smtplib
import ssl
from datetime import datetime, timedelta, date
from email.mime.text import MIMEText
from socket import gaierror

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from pytz import timezone

config = configparser.ConfigParser()
config.read('config.ini')

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

tz_germany = timezone('Europe/Berlin')

now = datetime.utcnow()
# Get Monday of current week
# dateStart is in UTC, attach local timezone later
dateStart = datetime(year=now.year, month=now.month, day=now.day,
                     hour=0, minute=0, second=0) - timedelta(days=now.weekday())
# Get Sunday of current week
# dateEnd is in UTC, attach local timezone later
dateEnd = datetime(year=dateStart.year, month=dateStart.month, day=dateStart.day,
                   hour=23, minute=59, second=59) + timedelta(days=6)


def send_mail(text):
    msg = MIMEText(text, 'plain', 'utf-8')
    msg['Subject'] = 'Termine für KW ' + dateStart.strftime('%V (%d.%m.') + dateEnd.strftime(' bis %d.%m.)')
    msg['From'] = config['Mail Server']['Login']
    msg['To'] = config['Outgoing Mail']['To']
    msg['reply-to'] = config['Outgoing Mail']['ReplyTo']

    try:
        # send your message with credentials specified above
        with smtplib.SMTP(config['Mail Server']['Server'], config['Mail Server']['Port']) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(config['Mail Server']['Login'], config['Mail Server']['Password'])
            server.send_message(msg)
            server.quit()

        # tell the script to report if your message was sent or which errors need to be fixed
        #print('Sent')
    except (gaierror, ConnectionRefusedError):
        print('Failed to connect to the server. Bad connection settings?')
    except smtplib.SMTPServerDisconnected:
        print('Failed to connect to the server. Wrong user/password?')
    except smtplib.SMTPException as e:
        print('SMTP error occurred: ' + str(e))


def fromisoformat(obj, date_to_format):
    return obj.fromisoformat(date_to_format) if date_to_format else date_to_format


def format_garbage_event(event):
    locale.setlocale(locale.LC_TIME, "de_DE")

    start_date = fromisoformat(date, event['start'].get('date'))
    return start_date.strftime('%A, %x') + ' ' + event['summary']


def format_event(event):
    locale.setlocale(locale.LC_TIME, "de_DE")

    out = []

    start_datetime = fromisoformat(datetime, event['start'].get('dateTime'))
    end_datetime = fromisoformat(datetime, event['end'].get('dateTime'))
    start_date = fromisoformat(date, event['start'].get('date'))
    end_date = fromisoformat(date, event['end'].get('date'))

    description = event.get('description')
    summary = event['summary']

    # A delta of one day
    one_day_duration = timedelta(days=1)

    def check_for_day_spanning(start, end, start_format, end_format):
        return start.strftime(start_format) if (end - start) >= one_day_duration \
            else start.strftime(start_format) + " bis " + end.strftime(end_format)

    if start_datetime and end_datetime:
        out.append(
            check_for_day_spanning(start_datetime, end_datetime, '%A, %x %H:%M', '%H:%M')
        )
    elif start_datetime:
        out.append(start_datetime.strftime('%A, %x %H:%M'))
    elif start_date and end_date:
        out.append(
            check_for_day_spanning(start_date, end_date, '%A, %x', '%A, %x')
        )

    out.append(summary)
    if description:
        out.append(description)
    return '\n'.join(out)


def get_events(service, calendar_id):
    """
    Retrieve planning events from Calendar API
    :param service: googleapiclient.discovery.Resource
    :return: List of events
    """
    # dateStart and dateEnd are in UTC, attach local timezone and then make it an ISO compatible timestamp
    time_min = dateStart.astimezone(tz_germany).isoformat()
    time_max = dateEnd.astimezone(tz_germany).isoformat()

    # see https://developers.google.com/resources/api-libraries/documentation/calendar/v3/python/latest/calendar_v3.events.html#list
    #     timeMin: string, Lower bound (exclusive) for an event's end time to filter by. Optional. The default is not to
    #     filter by end time. Must be an RFC3339 timestamp with mandatory time zone offset,
    #     for example, 2011-06-03T10:00:00-07:00, 2011-06-03T10:00:00Z.
    #     Milliseconds may be provided but are ignored. If timeMax is set, timeMin must be smaller than timeMax.
    #
    #     timeMax: string, Upper bound (exclusive) for an event's start time to filter by. Optional. The default is not
    #     to filter by start time. Must be an RFC3339 timestamp with mandatory time zone offset,
    #     for example, 2011-06-03T10:00:00-07:00, 2011-06-03T10:00:00Z.
    #     Milliseconds may be provided but are ignored. If timeMin is set, timeMax must be greater than timeMin.
    events_result = service.events().list(calendarId=calendar_id, timeMin=time_min,
                                          singleEvents=True, timeMax=time_max,
                                          orderBy='startTime').execute()
    return events_result.get('items', [])


def print_calendar_ids(service):
    """
    Prints calendar IDs that this bot has access to.
    :param service: googleapiclient.discovery.Resource
    """
    # Identify calendars:
    cals_result = service.calendarList().list().execute()
    cals = cals_result.get('items', [])

    # Calulate table dimensions
    len_id = len('ID')
    len_summary = len('Summary')
    for cal in cals:
        len_id = max(len(cal['id']), len_id)
        len_summary = max(len(cal['summary']), len_summary)

    sep = '  '
    total_length = (len_id + len(sep) + len_summary)

    # Print actual table of calendars
    print('>>> Calendars <<<'.center(total_length))
    print('-' * total_length)
    print('ID'.ljust(len_id) + sep + 'Summary'.ljust(len_summary))
    print('=' * total_length)
    for cal in cals:
        print(cal['id'].ljust(len_id) + sep + cal['summary'].ljust(len_summary))
    print('-' * total_length)


def main():
    """
    Dingfabrik Calendar Bot.
    Sends the calendar entries for the current week per mail.
    """
    credentials = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    pickle_file = 'token.pickle'
    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            credentials = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(pickle_file, 'wb') as token:
            pickle.dump(credentials, token)

    service = build('calendar', 'v3', credentials=credentials)
    planning_events = get_events(service, config['Calendars']['InternalPlanning'])
    garbage_collection_events = get_events(service, config['Calendars']['GarbageCollection'])

    out = []

    out.append('Hier sind die Termine aus dem internen Planungskalender für diese Woche:')
    out.append('-' * 72)
    out.append('')

    if not planning_events:
        out.append('Keine Termine in dieser Woche.')
    else:
        for event in planning_events:
            et = format_event(event)
            out.append(et)
            out.append('')
            out.append('***'.center(72))
            out.append('')

    out.append('')
    out.append('')
    out.append('Abfuhrtermine diese Woche:')
    out.append('-' * 72)

    if not garbage_collection_events:
        out.append('Keine Abfuhrtermine in dieser Woche.')
    else:
        for event in garbage_collection_events:
            et = format_garbage_event(event)
            out.append(et)

    out_text = '\r\n'.join(out)
    # To see the text that will be send, uncomment the next line
    #print(out_text)

    send_mail(out_text)

    # Remove comment to see a list of calendar IDs
    #print_calendar_ids(service)


if __name__ == '__main__':
    main()
