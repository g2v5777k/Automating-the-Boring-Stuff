import pickle
import os.path

from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


def authenticate(_file, drive_path):
    '''
    Uploads files from _files path to drive
    :param _file: path to file to upload
    :param drive_path: drive ids
    :return: None
    '''
    print('Uploading output to Google Drive...')
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/drive']
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open(r'C:\Scripts\RDOF_FieldTracking\token.pickle', 'rb') as token:
            creds = pickle.load(token, encoding='latin1')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r'C:\Scripts\RDOF_FieldTracking\credentials.json', SCOPES)
            creds = flow.run_local_server(port=0) 
        # Save the credentials for the next run
        with open(r'C:\Scripts\RDOF_FieldTracking\token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    fname = os.path.basename(_file)
    file_metadata = {'name': fname, 'parents': [drive_path]}
    media = MediaFileUpload(_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', chunksize=1024*1024, resumable=True)
    request = service.files().create(body=file_metadata,
                                     supportsAllDrives=True,
                                     media_body=media)
    response = None
    while response is None:
        status, response = request.next_chunk()
        if status:
            print("Uploaded %d%%." % int(status.progress() * 100))
    print("Upload complete.")


