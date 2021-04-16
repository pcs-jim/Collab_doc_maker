import os
from tkinter import *
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv

import httplib2
import requests
from googleapiclient import errors
from googleapiclient.discovery import build
from oauth2client.file import Storage
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

"""
Created by: Jaemyong Chang
On: 3/31/2021

The purpose of this is to set up the initial folders and files for a collaborative notetaking study
under the guidance of Professor Mik Fanguy of KAIST university of Daejeon, South Korea.

Example of files and folder creation.
https://drive.google.com/drive/folders/1iRQVclSh7JNxyvgqoO5Sw1s1T6UWhrj5?usp=sharing

Example of the templates.
https://drive.google.com/drive/folders/1z92dOmC0M_jkMv6651mtRQoGS3ABHy2Y?usp=sharing

You'll need to download the 'credentials.json' from your Google account.
"""

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents']

# Get permission via web browser.
gauth = GoogleAuth()
gauth.LoadCredentialsFile('google_credentials.txt')

if gauth.credentials is None:
    gauth.LocalWebserverAuth()
elif gauth.access_token_expired:
    gauth.Refresh()
else:
    gauth.Authorize()
gauth.SaveCredentialsFile('google_credentials.txt')

drive = GoogleDrive(gauth)


# Get authorization.
storage = Storage('google_credentials.txt')
credentials = storage.get()
http = httplib2.Http()
http = credentials.authorize(http)

drive_service_v2 = build('drive', 'v2', http=http)
drive_service_v3 = build('drive', 'v3', http=http)

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secrets.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('docs', 'v1', credentials=credentials)
doc_service = build('docs', 'v1', credentials=creds)

template_folder_id = None
parent_id = None

doc_ids = []
permission_list = []


def search_for_folder():
    """
        Initial search of folders.
    """
    global template_folder_id
    global doc_ids
    template_name = template_folder_name_entry.get()

    page_token = None
    while True:
        response = drive_service_v2.files().list(q="mimeType='application/vnd.google-apps.folder'",
                                              spaces='drive',
                                              fields='nextPageToken, items(id, title)',
                                              pageToken=page_token).execute()
        for file in response.get('items', []):
            # Process change
            if file.get('title') == template_name:
                template_folder_id = file.get('id')
                print('Found file: %s (%s)' % (file.get('title'), file.get('id')))
                break
        page_token = response.get('nextPageToken', None)

        if page_token is None or template_folder_id is not None:
            break

    if template_folder_id is None:
        folder_id_label = Label(root, text='File Not Found')
        folder_id_label.grid(row=1, column=1)
    else:
        template_folder = drive.ListFile({'q': "'%s' in parents and trashed=false" % template_folder_id}).GetList()
        for doc in template_folder:
            print(doc['title'], doc['id'])
            doc_ids.append({'title': doc['title'], 'id': doc['id']})
        template_id_label = Label(root, text=str(len(doc_ids)) + ' doc(s) found.')
        template_id_label.grid(row=1, column=1)


def create_root_folder():
    global parent_id
    folder_name = root_folder_name_entry.get()
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }

    file = drive_service_v3.files().create(body=file_metadata, fields='id').execute()
    parent_id = file.get('id')
    folder_id_label = Label(root, text=parent_id)
    folder_id_label.grid(row=3, column=1)


def create_drive_folder_in_parent(root_id, folder_name):

    # for names in sub_folder_names:
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [root_id]
    }

    file = drive_service_v3.files().create(body=file_metadata, fields='id').execute()

    return file.get('id')


def create_folders():
    global parent_id
    global doc_ids
    global permission_list

    section_names = section_name_entry.get().split()
    group_names = group_folder_name_entry.get().split()

    for section_name in section_names:
        section_id = create_drive_folder_in_parent(parent_id, section_name)
        section_group_permission = {'section': section_name}

        group_dict = {}
        for group_name in group_names:
            group_id = create_drive_folder_in_parent(section_id, group_name)
            group_dict[group_name] = group_id

            for document in doc_ids:
                title = document['title']
                id = document['id']

                body = {
                    'name': title,
                    'parents': [group_id, ]
                }

                try:
                    drive_service_v3.files().copy(fileId=id, body=body).execute()
                except:
                    pass

        section_group_permission['groups'] = group_dict
        permission_list.append(section_group_permission)

    Label(root, text='Folders and Files Created').grid(row=6, column=1)


def callback(request_id, response, exception):
    if exception:
        print(exception)
    else:
        print("Permission ID:", response.get('id'))


batch = drive_service_v3.new_batch_http_request(callback=callback)


def grant_permission():
    global permission_list
    csv_filename = permission_filename_entry.get()

    with open(csv_filename, 'r') as file:
        reader = csv.reader(file)
        for idx, csv_row in enumerate(reader):
            # Ignore header.
            if idx == 0:
                continue

            for i in range(len(permission_list)):
                # Check section and group name.
                if csv_row[1] == permission_list[i]['section'] and csv_row[2] in permission_list[i]['groups']:
                    print(permission_list[i]['groups'])
                    print('\t', 'Send permission to: ' + csv_row[0], permission_list[i]['groups'][csv_row[2]])
                    # Gives permission to a user.
                    user_permission = {
                        'type': 'user',
                        'role': 'writer',
                        'emailAddress': csv_row[0]
                    }
                    batch.add(drive_service_v3.permissions().create(
                        fileId=permission_list[i]['groups'][csv_row[2]],
                        body=user_permission,
                        fields='id'
                    ))
                    batch.execute()


# UI
root = Tk()

root.title('Create Files and Folders')

# Search for the templates folder with the Google Docs templates.
Label(root, text='Template Folder').grid(row=0, column=0)
template_folder_name_entry = Entry(root, width=60, borderwidth=5)
template_folder_name_entry.grid(row=0, column=1)

search_template_btn = Button(root, text='Search', command=search_for_folder)
search_template_btn.grid(row=1, column=0)

# Create parent (root) folder.
Label(root, text='Root Folder').grid(row=2, column=0)
root_folder_name_entry = Entry(root, width=60, borderwidth=5)
root_folder_name_entry.grid(row=2, column=1)

create_parent_btn = Button(root, text='Create Folder', command=create_root_folder)
create_parent_btn.grid(row=3, column=0)

# Create section folders
Label(root, text='Section Name(s)').grid(row=4, column=0)
section_name_entry = Entry(root, width=60, borderwidth=5)
section_name_entry.grid(row=4, column=1)

# Create group folders.
Label(root, text='Group Name(s)').grid(row=5, column=0)
group_folder_name_entry = Entry(root, width=60, borderwidth=5)
group_folder_name_entry.grid(row=5, column=1)

# Create files and folders.
create_folders_btn = Button(root, text='Create Folder(s)', command=create_folders)
create_folders_btn.grid(row=6, column=0)

# Grant permissions from csv file
Label(root, text='CSV Filename').grid(row=7, column=0)
permission_filename_entry = Entry(root, width=60, borderwidth=5)
permission_filename_entry.grid(row=7, column=1)

grant_permission_btn = Button(root, text='Grant Permissions', command=grant_permission)
grant_permission_btn.grid(row=8, column=0)

template_folder_name_entry.insert(END, 'Templates Spring 2021')
root_folder_name_entry.insert(END, 'ABC_Spring2021_Test')
section_name_entry.insert(END, 'Section1 Section2')
group_folder_name_entry.insert(END, 'Group1 Group2')
permission_filename_entry.insert(END, 'term_year_gmail_addresses.csv')


root.mainloop()


