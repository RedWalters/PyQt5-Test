import os
from django.conf import settings
import zipfile
import shutil
import xmltodict

try:
    settings.configure()
except:
    pass

def get_sheet_details(file_path):
    sheets = []        
    file_name = os.path.splitext(os.path.split(file_path)[-1])[0]
    # Make a temporary directory with the file name
    if os.path.exists(os.path.join(settings.MEDIA_ROOT, file_name)):
        shutil.rmtree(os.path.join(settings.MEDIA_ROOT, file_name))
    else:
        pass
    directory_to_extract_to = os.path.join(settings.MEDIA_ROOT, file_name)
    os.mkdir(directory_to_extract_to)

    # Extract the xlsx file as it is just a zip file
    zip_ref = zipfile.ZipFile(file_path, 'r')
    zip_ref.extractall(directory_to_extract_to)
    zip_ref.close()

    # Open the workbook.xml which is very light and only has meta data, get sheets from it
    path_to_workbook = os.path.join(directory_to_extract_to, 'xl', 'workbook.xml')
    with open(path_to_workbook, 'r') as f:
        xml = f.read()
        dictionary = xmltodict.parse(xml)
        for sheet in dictionary['workbook']['sheets']['sheet']:
            sheet_details = {
                'id': sheet['@sheetId'], # can be @sheetId for some versions
                'name': sheet['@name'] # can be @name
            }
            sheets.append(sheet_details)

    # Delete the extracted files directory
    shutil.rmtree(directory_to_extract_to)
    return sheets