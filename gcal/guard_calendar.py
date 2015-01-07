#!/opt/gcal/python25/bin/python
try:
  from xml.etree import ElementTree # for Python 2.5 users
except ImportError:
  from elementtree import ElementTree
from datetime import datetime
from email.mime.text import MIMEText
import logging
import logging.handlers
import smtplib
import time
import yaml
import sys
import httplib2
import argparse

# from oauth2client import client
from apiclient import errors
from apiclient.discovery import build
from apiclient.http import MediaFileUpload
from apiclient.errors import HttpError
from oauth2client.client import AccessTokenRefreshError
from oauth2client.client import flow_from_clientsecrets
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.file import Storage
from oauth2client.tools import run
from oauth2client import file, client, tools

class GuardCalendar:

  global client
  global service
  global authorized_creators
  global logger

  def __init__(self, domain):
    self.domain = domain
        
  def GetCreators(self, user):

    authorized_creators = set()

    acl = service.acl().list(calendarId=user).execute()

    logger.debug( "---- Sharing permissions:")
    for rule in acl['items']:
      logger.debug( 'user %s has role: %s' % (rule['id'], rule['role']) )
      if ((rule['role'] == 'owner') or ((rule['role'] == 'writer'))):
          logger.debug("Adding %s to authorized list" % rule['scope']['value'])
          authorized_creators.add(rule['scope']['value'])

    return authorized_creators

  def QueryFutureEvents(self, user, start_date=datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')):

    logger.debug( "---- Finding events after start_date: %s" % start_date )
    
    the_events = []
    page_token = None
    while True:
      events = service.events().list(calendarId=user, pageToken=page_token, timeMin=start_date).execute()
      for event in events['items']:
        logger.debug('Adding %s, %s created by: %s' % (event['id'], event['summary'], event['creator']['email']))
        my_event = {'id':event['id'],
                    'summary': event['summary'],
                    'creator': event['creator']['email'],
                    'start': event['start']['dateTime']
                    }
        the_events.append(my_event)

      page_token = events.get('nextPageToken')
      if not page_token:
        break

    return the_events
  
  def DeleteEvent(self, user, eventId):

    try:
        logger.info("Deleting eventId %s from user: %s" % (eventId, user))
        # service.events().delete(calendarId=user, eventId=eventId).execute()
    except Exception, error:
          logger.error( 'An error occurred in DeleteEvent(): %s' % error )

  def SendEmail(self, event, email):

    try:
        organizer = event['creator']
        start_time = time.strptime(event['start'].split('T')[0], '%Y-%m-%d')
        msgbody = email.get('message').replace('{0}', organizer).replace('{1}', event['summary']).replace('{2}', time.strftime('%m/%d/%Y', start_time))
        msg = MIMEText(msgbody)
        msg['Subject'] = email.get('subject')
        msg['From'] = 'gcal-donot-reply@' + self.domain
        msg['To'] = event['creator']

        mail_server = smtplib.SMTP('smtp.gene.com')
        mail_server.sendmail(msg['From'], msg['To'], msg.as_string())
        mail_server.close()
    except Exception, error:
          logger.error( 'An error occurred in SendEmail(): %s' % error)

  @classmethod
  def gdata_to_datetime(self, gdata_time):
    return datetime.fromtimestamp(time.mktime(time.strptime(gdata_time.split('.')[0],'%Y-%m-%dT%H:%M:%S')))

def main(argv):
    global service
    global authorized_creators
    global logger

    parser = argparse.ArgumentParser(description='Removes unauthorized events from gCal', prog=argv[0])
    parser.add_argument('--loglevel', nargs='?', default='info',
                       help='Set log level (default: info)')

    args = parser.parse_args()

    CLIENT_SECRET = '../../client_secret_gappsadm_genepoc.json' # downloaded JSON file

    # Check https://developers.google.com/drive/scopes for all available scopes
    OAUTH_SCOPE = 'https://www.googleapis.com/auth/calendar'

    storage = Storage('../../storage_gappsadm_genepoc.json')
    credentials = storage.get()
    if not credentials or credentials.invalid:
      flow = client.flow_from_clientsecrets(CLIENT_SECRET, OAUTH_SCOPE)
      credentials = tools.run(flow, storage)

    # Create an httplib2.Http object and authorize it with our credentials
    http = httplib2.Http()
    http = credentials.authorize(http)

    service = build(serviceName='calendar', version='v3', http=http)

    numeric_level = getattr(logging, args.loglevel.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError('Invalid log level: %s' % args.loglevel)
    logging.basicConfig(level=numeric_level)

    # logging.basicConfig()
    LOG_FILENAME = './guard.log'
    logger = logging.getLogger('GuardCalendar')


    # Add the log message handler to the logger
    handler = logging.handlers.RotatingFileHandler(LOG_FILENAME, maxBytes=2048, backupCount=5)

    fmt = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    handler.setFormatter(fmt)
    logger.addHandler(handler)

    logger.info("Log level set to %s" % args.loglevel)

    file = open("config/guard_calendar.yml")
    config = yaml.load(file)
    file.close()

    users = config.get('toguard').split(',')
    theDomain = config.get('domain')

    # cannot_delete_deadline = GuardCalendar.gdata_to_datetime("2010-01-26T01:00:00.000Z")

    for user in users:
      email_address = user + '@' + config.get('domain')
      guard = GuardCalendar(theDomain)
      authorized_creators = guard.GetCreators(email_address)
      future_events = guard.QueryFutureEvents(email_address)
      file = open("config/" + user + ".yml")
      email = yaml.load(file)
      file.close()

      for x in range(len( future_events )):
          my_event = future_events[x]
          if (my_event['creator'] not in authorized_creators):
            logger.info("%s is not allowed to create an event on calendar of: %s" % (my_event['creator'],email_address))
            guard.DeleteEvent(email_address, my_event['id'])
            guard.SendEmail(my_event, email)
            logger.info('notified organizer ' + my_event['creator'] + ' for this deletion')

if __name__ == '__main__':
  main(sys.argv)
