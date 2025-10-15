# Double check configuration before running
# currently only checks signedPPA group on invoice board. does not work for ccppas

import requests
import os
import shutil
import time
from PyPDF2 import PdfMerger
import json
from dotenv import load_dotenv

load_dotenv()


# === CONFIG ===
API_KEY = os.getenv("MONDAY_API_KEY")
BOARD_ID = os.getenv("MONDAY_BOARD_ID")
GROUP_ID = 'group_mkvpt40z'
#FILE_COLUMN_KEYS = {'invoice' : 'file_mkv6n4tv', 'PPA':'file_mkvk90zm'}  # file column key value pairs
FILE_COLUMN_IDS = ['file_mkv6n4tv','file_mkvk90zm']  # file column IDs
OUTPUT_FOLDER = 'monday_files'  # Base output directory

# === CONSTANTS ===
API_URL = 'https://api.monday.com/v2'
HEADERS = {
    'Authorization': API_KEY,
    'Content-Type': 'application/json'
}


def make_graphql_query(query: str, variables: dict = None):
    payload = {'query': query}
    if variables:
        payload['variables'] = variables
    response = requests.post(API_URL, headers=HEADERS, json=payload)
    if response.status_code != 200:
                print(f"HTTP Error: {response.status_code}")
                print(response.text)
    response.raise_for_status()
    return response.json()


def get_items_in_group(board_id: str, group_id: str):
    query = '''
    query ($board_id: ID!, $group_id: String!) {
      boards(ids: [$board_id]) {
        groups(ids: [$group_id]) {
          items_page{
            items {
                id
                name
                assets{
                    id
                    public_url
                    name
                }
                column_values {
                    id
                    value
                }
            }
          }
        }
      }
    }
    '''
    variables = {'board_id': int(board_id), 'group_id': group_id}
    data = make_graphql_query(query, variables)
    # Check for GraphQL errors inside the response body
    if 'errors' in data:
        print("GraphQL Errors found getting items in group:")
        for error in data['errors']:
            print(f"- Message: {error.get('message')}")
            if 'locations' in error:
                print(f"  Location: {error['locations']}")
            if 'path' in error:
                print(f"  Path: {error['path']}")
        return
    else:
        return data['data']['boards'][0]['groups'][0]['items_page']['items']

def download_file(url: str, save_path: str):
    try:
        #print(f"⬇️ Downloading from: {url}")
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            with open(save_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        print(f"✔️ Saved to: {save_path}")
    except requests.exceptions.HTTPError as e:
        print(f"❌ HTTP error: {e}")
    except Exception as e:
        print(f"❌ Other error: {e}")


def process_items(items: list, file_column_ids: list):
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    for item in items:
        item_id = item['id']
        item_name = item['name'].strip().replace(' ', '_').replace('/', '_')
        folder_name = os.path.join(OUTPUT_FOLDER, f"temp")
        #create new temp folder
        os.makedirs(folder_name, exist_ok=True)
        print(f"\n📁 Processing item: {item_name} (ID: {item_id})")
        assets = []
        for column in item['column_values']:
            if column['id'] not in file_column_ids:
                continue
            if not column['value']:
                continue
            try:
                files = json.loads(column['value']).get('files', [])
                for file in files: 
                    asset_id = file['assetId']
                    assets.append(str(asset_id))
            except Exception as e:
                print(f"⚠️ Error processing column '{column['id']}' in item {item_id}: {e}")
        for asset in item['assets']:
            try:
                if(asset['id'] in assets):
                    file_url = asset['public_url']
                    print(f"➡️  Downloading: {asset['name']}")
                    download_file(file_url, os.path.join(folder_name, asset['name']))
            except Exception as e:
                print(f"⚠️ Error downlading file from column '{column['id']}' in item {item_id}: {e}")
        merge_pdfs(folder_name, item['name'] + "_merged.pdf")
        #delete previous temp folder
        try:
            shutil.rmtree(folder_name)
            print(f"Folder '{folder_name}' and its contents deleted successfully.")
        except OSError as e:
            print(f"❌Error: {e.filename} - {e.strerror}.")
            return
        

def merge_pdfs(folder_path,file_name):
    try:
        output_path = os.path.join(OUTPUT_FOLDER,file_name)        
        if os.path.exists(output_path):
            os.remove(output_path)
            print(f"Removed existing file: {output_path}")
        # List all PDF files in the folder
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
        #check if empty
        if not pdf_files:
            print("No PDF files found in the folder.")
            return
        #find the ppa and put it on top
        for i in range(len(pdf_files)):
            print(pdf_files[i])
            if('PPA' in pdf_files[i]):
                pdf_files.insert(0, pdf_files.pop(i))    
                break    
        #initilize merger
        merger = PdfMerger()          
        #merge pdfs
        for pdf in pdf_files:
            #get full file path
            full_path = os.path.join(folder_path, pdf)
            #merge
            try:
                print(f"✅ Adding: {pdf} to {output_path}")
                merger.append(full_path)
            except Exception as e:
                print(f"❌ Skipped {pdf}: {e}")
        #write merged pdf to file and close merger
        merger.write(output_path)
        merger.close()
        print(f"\nAll PDFs merged into: {output_path}")
    except Exception as e:
        print(f"❌ Error with merging in {folder_path}: {e}")
    return

def print_dir(folder_path):
    if not os.path.isdir(folder_path):
        print(f"Directory does not exist: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if not os.path.isdir(folder_path):
            print(f"Directory not found: {folder_path}")
            return
        for filename in os.listdir(folder_path):
            filepath = os.path.join(folder_path, filename)
            if os.path.isfile(filepath):
                try:
                    print(f"Sending to printer: {filepath}")
                    os.startfile(filepath, "print")
                    time.sleep(30)  # Wait a bit to avoid overloading the print queue
                except Exception as e:
                    print(f"Error printing {filepath}: {e}")
    return

if __name__ == '__main__':
    print("🔄 Fetching items from Monday.com group...")
    items = get_items_in_group(BOARD_ID, GROUP_ID)
    print(f"✅ Found {len(items)} items.")
    process_items(items, FILE_COLUMN_IDS)