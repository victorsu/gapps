#!/usr/bin/python

import sys
import datetime
import codecs
import httplib2
import time

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


def main(argv):
  parent_id = 0

  CLIENT_SECRET = 'client_secret_vsu.json' # downloaded JSON file

  # Check https://developers.google.com/drive/scopes for all available scopes
  OAUTH_SCOPE = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.appdata',
    'https://www.googleapis.com/auth/drive.apps.readonly'
]

  # Redirect URI for installed apps
  REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'

  storage = Storage('storage_vsu2.json')
  credentials = storage.get()
  if not credentials or credentials.invalid:
    flow = client.flow_from_clientsecrets(CLIENT_SECRET, ' '.join(OAUTH_SCOPE))
    credentials = tools.run(flow, storage)

  # Create an httplib2.Http object and authorize it with our credentials
  http = httplib2.Http()
  http = credentials.authorize(http)

  drive_service = build('drive', 'v2', http=http)

  filename = 'Dashboard Project Name List.xlsx'
  new_filename = 'Dashboard Project Name List.xlsx'

  mime_type="application/vnd.ms-excel"
  new_mime_type="application/vnd.ms-excel"

  title = 'Dashboard Project Name List'
  new_title = 'Dashboard Project Name List'

  new_revision = True
  # upload_type = 'media'

  # Update existing file:
  file_id = '1MlSYNzhskJHcZpTFiUuzAmoGXEKhKD_rFCYVyYhQuSI'

  try:
    # First retrieve the file from the API.
    file = drive_service.files().get(fileId=file_id).execute()

    # File's new metadata.
    file['title'] = title
    file['mimeType'] = mime_type

    # File's new content.
    media_body = MediaFileUpload(
        new_filename, mimetype=new_mime_type, resumable=True)

    # Send the request to the API.
    updated_file = drive_service.files().update(
        fileId=file_id,
        body=file,
        newRevision=new_revision,
        media_body=media_body).execute()
    return updated_file
  except errors.HttpError, error:
    print 'An error occurred: %s' % error
    return None

  """
  # Insert new file

  media_body = MediaFileUpload(filename, mimetype=mime_type, resumable=False)

  body = {
    'title': title,
    'mimeType': mime_type
  }
  # Set the parent folder.
  if parent_id:
    body['parents'] = [{'id': parent_id}]

  try:
    file = drive_service.files().insert(
        body=body,
        convert=True,
        media_body=media_body).execute()

    # Uncomment the following line to print the File ID
    print 'File ID: %s' % file['id']

    return file
  except errors.HttpError, error:
    print 'An error occured: %s' % error
    return None

  """

if __name__ == '__main__':
  main(sys.argv)

