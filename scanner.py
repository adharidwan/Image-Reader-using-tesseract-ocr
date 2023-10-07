import cv2
import os
import pytesseract
from PIL import Image
from googleapiclient.discovery import build
from google.oauth2 import service_account


# path ke google sheet api credential json
creds_path = r'C:\Users\lenov\OneDrive\Documents\Project\CardScanner\Google Sheets API JSON\cardscanner-401308-379f6b1456d8.json'

# google sheet api
creds = service_account.Credentials.from_service_account_file(creds_path, scopes=['https://www.googleapis.com/auth/spreadsheets'])
service = build('sheets', 'v4', credentials=creds)
spreadsheet_id = '1Wxp_Sz_3inTGFhcfC04V9IL1QLFixU52PQSsyO50V7w'

#somehow someway pythonny ngga detect tesseract ocr jdi harus dimasukin paksa
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  

# pake ocr buat dapetin textny (coba pake keycolor code biar ngga perlu ocr)
def extract_text_from_image(image):
    text = pytesseract.image_to_string(image)
    return text

# upload ke spreadsheed
def update_spreadsheet(spreadsheet_id, range_name, values):
   body = {
       'values': [values],
   }
   result = service.spreadsheets().values().update(
       spreadsheetId=spreadsheet_id,
       range=range_name,
       valueInputOption='RAW',
       body=body
   ).execute()
   print(f'Updated {result.get("updatedCells")} cells.')

# buka kamera
cap = cv2.VideoCapture(0)

while True:
    # foto
    ret, frame = cap.read()

    # ocr untuk text 
    text_from_card = extract_text_from_image(frame)

    # display in hasil ocr
    cv2.imshow("Scanner Kartu", frame)

    # cek untuk input dan keluar
    if cv2.waitKey(1) & 0xFF == ord('k'):
        range_to_update = 'Nama!B1'  # range spreadsheet
        update_spreadsheet(spreadsheet_id, range_to_update, [text_from_card])
        break

# matikan kamera dan tutup cv
cap.release()
cv2.destroyAllWindows()
