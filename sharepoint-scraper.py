import os
import time
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Load environment variables
load_dotenv()

# Retrieve environment variables
site_url = os.getenv('SHAREPOINT_SITE_URL')
user = os.getenv('SHAREPOINT_USER')
password = os.getenv('SHAREPOINT_PASSWORD')

# This checks if the variables have been set correctly.
print("URL:", site_url)
print("User:", user)
print("Password:", password)
time.sleep(3)  # Pauses the script for 5 seconds

# Setup Authentication Context
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(user, password):
    ctx = ClientContext(site_url, ctx_auth)
    web = ctx.web
    try:
        ctx.load(web)
        ctx.execute_query()
        print("Connected to SharePoint site:", web.properties['Title'])
    except Exception as e:
        print(f"Failed to connect to SharePoint site. {str(e)}")

    # Set up the download directory
    download_path = Path.home() / 'Downloads' / f'SharePoint_Downloads_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    download_path.mkdir(parents=True, exist_ok=True)

    # Retrieve documents
    library = web.lists.get_by_title('Documents')
    folders = library.root_folder.folders
    try:
        ctx.load(folders)
        ctx.execute_query()
    except Exception as e:
        print(f"Failed to retrieve folders. {str(e)}")

    for folder in folders:
        files = folder.files
        ctx.load(files)
        try:
            ctx.execute_query()
        except Exception as e:
            print(f"Failed to execute query to load files from folder {folder.server_relative_url}: {str(e)}")
        for file in files:
            print(f"Downloading {file.name}...")
            try:
                file.download(str(download_path / file.name))
                ctx.execute_query()
            except Exception as e:
                print(f"Failed to download {file.name}. {str(e)}")
else:
    print(ctx_auth.get_last_error())
