from flask import Flask, request, jsonify
import requests
import os
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
import shutil
from win32com import client
import pythoncom


MONDAY_API_URL = "https://api.monday.com/v2"
MONDAY_FILE_URL = "https://api.monday.com/v2/file"

API_KEY = os.getenv("MONDAY_API_KEY")
BOARD_ID = os.getenv("MONDAY_BOARD_ID")

empty_PPA = "PPA Form.xlsx"

column_ids = {'Vendor':'text_mkv6s9er', 'Date':'date4','Total':'numeric_mkv65026','Inkind':'numeric_mkv6tbbk','Department':'text_mkvft17m','Description':'text_mkve5mct','OrderedBy':'multiple_person_mkvfj0v1','PPA file':'file_mkvm3pf3','Workflow':'color_mkv67wpq'}

load_dotenv()

app = Flask(__name__)

HEADERS = {
    "Authorization" : API_KEY
}


class add_PPA_to_board:

    def __init__(self, board_id, API_KEY):
        self.board_id = board_id
        self.API_KEY = API_KEY
        self.current_invoice_number = 0
        self.current_item_id = 0
        self.current_item_data = {}
        return
    
    def upload_to_monday(self, file_path):
        #create query and send file
        item_id = self.current_item_id
        column_id = column_ids['PPA file']
        file = file_path
        if not item_id or not column_id or not file:
            return jsonify({"error": "Missing item_id, column_id, or file"}), 400
        query = f'''
        mutation ($file: File!) {{
        add_file_to_column (file: $file, item_id: {item_id}, column_id: "{column_id}") {{
            id
        }}
        }}
        '''
        #open pdf to prepare to send
        PPA_PDF_file = open(file_path,'rb')
        multipart_data = {
            'query': (None, query),
            'variables[file]': PPA_PDF_file
        }
        response = requests.post(MONDAY_FILE_URL, headers=HEADERS, files=multipart_data)
        #close pdf
        PPA_PDF_file.close()
        #check for errors
        if response.status_code != 200:
            print(f"HTTP Error while uploading: {response.status_code}")
            print(response.text)
        else:
            response_json = response.json()
            # Check for GraphQL errors inside the response body
            if 'errors' in response_json:
                print("GraphQL Errors found during file upload:")
                for error in response_json['errors']:
                    print(f"- Message: {error.get('message')}")
                    if 'locations' in error:
                        print(f"  Location: {error['locations']}")
                    if 'path' in error:
                        print(f"  Path: {error['path']}")
        print('PDF Uploaded')
        try:
            os.remove(file_path)
            print(f"File '{file_path}' deleted successfully.")
        except Exception as e:
            print(f"An error occurred while deleting PDF file: {e}")
        return "PDF Uploaded"
    
    def get_item(self, item_id):
        self.current_item_id = item_id
        data_good = False
        self.current_item_data = {}
        # Query Monday Board for Item info
        query = {
            'query': f'''
                query {{
                    items(ids: {item_id}) {{
                        id
                        name
                        column_values {{
                            id
                            column{{
                                title
                            }}
                            text
                        }}
                    }}
                }}
            '''
        }
        response = requests.post(MONDAY_API_URL, json=query, headers=HEADERS)
        #parse response for column values and check for errors
        if response.status_code != 200:
            print(f"HTTP Error: {response.status_code}")
            print(response.text)
        else:
            response_json = response.json()
            # Check for GraphQL errors inside the response body
            if 'errors' in response_json:
                print("GraphQL Errors found:")
                for error in response_json['errors']:
                    print(f"- Message: {error.get('message')}")
                    if 'locations' in error:
                        print(f"  Location: {error['locations']}")
                    if 'path' in error:
                        print(f"  Path: {error['path']}")
            else:
                # No GraphQL errors; process the data
                item_data = response_json['data']['items'][0]
                self.current_invoice_number = item_data['name']
                #print(f"Invoice #: {item_data['name']}")
                #print("Column Values:")
                # loop through column_values and extract data needed
                for column in item_data['column_values']:
                    
                    self.current_item_data[column['id']] = (column['text'])
                    print(f"- {column['column']['title']} ({column['id']}): {column['text']}")
                print(self.current_item_data[column_ids['PPA file']])
                if not self.current_item_data[column_ids['PPA file']]:
                    if self.current_item_data[column_ids['Workflow']] == 'PPA Creation':
                        print("Data is good to go")
                        data_good = True
        return data_good
    
    def xlsxtopdf(self, file_path):
        #initialize pythoncom
        excel = client.Dispatch("Excel.Application",pythoncom.CoInitialize())
        #open workbook
        sheets = excel.Workbooks.Open(os.path.abspath(file_path))
        work_sheets = sheets.Worksheets[0]
        #export as pdf
        work_sheets.ExportAsFixedFormat(0, os.path.abspath(file_path[:len(file_path)-4] + 'pdf'))
        #close workbook and exit pythoncom
        sheets.Close()
        excel.Quit()
        print('PDF created successfully')
        try:
            os.remove(file_path)
            print(f"File '{file_path}' deleted successfully.")
        except Exception as e:
            print(f"An error occurred while deleting xlsx file: {e}")
        return

    def create_ppa(self):
        data = self.current_item_data
        new_ppa_name ="PPA Form -" + data[column_ids['Vendor']] + self.current_invoice_number + ".xlsx"
        #destination_file = destination_folder + r"/" + new_ppa_name
        destination_file = new_ppa_name
        #Create the file for ppa
        try:
            shutil.copy(empty_PPA, destination_file)
            print(f"'{empty_PPA}' copied to '{destination_file}' successfully.")
        except FileNotFoundError:
            print(f"Error: '{empty_PPA}' not found.")
        except Exception as e:
            print(f"An error occurred when creating empty PPA form: {e}")
        #open file
        try:
            workbook = load_workbook(destination_file)
            sheet = workbook.active
            sheet["L13"] = data[column_ids['Vendor']]
            sheet["B13"] = data[column_ids['Department']]
            sheet["G19"] = self.current_invoice_number
            sheet["B23"] = 1
            sheet["N23"] = data[column_ids['Total']]
            sheet["C23"] = data[column_ids['Description']]
            sheet["C24"] = "job date: " + data[column_ids['Date']]
            sheet["D40"] = "PPA Prepared by " + data[column_ids['OrderedBy']]
            workbook.save(filename=destination_file)
        except Exception as e:
            print(e)
        else:
            print('PPA Created Succesfully')
        self.xlsxtopdf(destination_file)
        self.upload_to_monday(destination_file[:len(destination_file)-4]+'pdf')
        return True

add = add_PPA_to_board(BOARD_ID, API_KEY)

@app.route("/add_ppa", methods=["POST"])
def add_ppa():
    data = request.json
    if data:
        print("Request recieved\n")
    if not data:
        print("Request empty")
        return ":("
    item_id = data["event"]["itemId"]
    good_data = add.get_item(item_id)
    if good_data == True:
        add.create_ppa()
    else:
        print('Stopped Creation process, bad data or PPA present')
    return 'done'

if __name__ == "__main__":
    app.run(debug=True)