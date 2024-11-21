from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials
import pkg_resources


def upload_file_to_gdrive(folder_id, path, file_name):
    gauth = GoogleAuth()
    # NOTE: if you are getting storage quota exceeded error, create a new service account, and give that service account permission to access the folder and replace the google_credentials.
    gauth.credentials = ServiceAccountCredentials.from_json_keyfile_name(pkg_resources.resource_filename(__name__, "credentials.json"), scopes=['https://www.googleapis.com/auth/drive'])

    drive = GoogleDrive(gauth)

    file = drive.CreateFile({'parents': [{"id": folder_id}], 'title': file_name})

    file.SetContentFile(path + '/' + file_name)
    file.Upload()
