import os
import asyncio
import requests
import aiohttp
from msgraph import GraphServiceClient
from msal import ConfidentialClientApplication
from azure.identity.aio import ClientSecretCredential
from datetime import datetime

# ==== CONFIG ====
client_id = os.getenv("client_id")
tenant_id = os.getenv("tenant_id")
client_secret = os.getenv("client_secret")
team_id = os.getenv("team_id")
channel_id = os.getenv("channel_id")
file_path = os.getenv("file_path", "android/app/build/outputs/apk/production/release/app-production-release.apk")
file_name = os.getenv("name", os.path.basename(file_path))
chunk_size = 5 * 1024 * 1024  # 5 MB
target_folder = os.getenv("target_folder")
app_env = os.getenv("app_env", "staging")
app_version_name = os.getenv("version_name")
app_version_code = os.getenv("version_code")
app_name = None
action = os.getenv("usecase", "upload")

# ==== AUTH ====
print("APK name=",file_name)
authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(
    client_id=client_id,
    authority=authority,
    client_credential=client_secret
)

result = app.acquire_token_for_client(scopes=scopes)

if "access_token" not in result:
    raise Exception(f"Could not obtain token: {result.get('error_description')}")

access_token = result["access_token"]
print("Obtained access token")
credential = ClientSecretCredential(tenant_id, client_id, client_secret)
graph_client = GraphServiceClient(credentials=credential, scopes=scopes)
print("Successfully Authenticated")

async def upload_to_teams():
    # Step 1: Locate Teams Channel Folder
    print("Getting drive id")
    folder = await graph_client.teams.by_team_id(team_id).channels.by_channel_id(channel_id).files_folder.get()
    drive_id = folder.parent_reference.drive_id
    print("Getting folder id")
    folder_id = folder.id

    # Get target folder ID
    items = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(folder_id).children.get()
    target_folder_id = None
    for item in items.value:
        if item.name == target_folder and item.folder:
            target_folder_id = item.id
            break

    if not target_folder_id:
        raise Exception(f"Target folder '{target_folder}' not found in channel.")

    # Step 2: Create upload session
    print("Creating upload session")
    upload_session_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{target_folder_id}:/{file_name}:/createUploadSession"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    upload_body = {
        "item": {
            "@microsoft.graph.conflictBehavior": "rename",
            "name": file_name
        }
    }
    upload_session_resp = requests.post(upload_session_url, headers=headers, json=upload_body)
    upload_session_resp.raise_for_status()
    upload_url = upload_session_resp.json()["uploadUrl"]

    # Step 3: Upload in chunks
    print("Uploading files in chuncks")
    file_size = os.path.getsize(file_path)
    with open(file_path, "rb") as f:
        bytes_uploaded = 0
        chunk_num = 1
        while bytes_uploaded < file_size:
            chunk = f.read(chunk_size)
            start = bytes_uploaded
            end = bytes_uploaded + len(chunk) - 1
            content_range = f"bytes {start}-{end}/{file_size}"

            print(f"[Chunk {chunk_num}] Uploading {content_range}")
            chunk_resp = requests.put(upload_url, headers={
                "Content-Length": str(len(chunk)),
                "Content-Range": content_range
            }, data=chunk)

            if chunk_resp.status_code not in [200, 201, 202]:
                raise Exception(f"Chunk upload failed: {chunk_resp.text}")

            bytes_uploaded += len(chunk)
            chunk_num += 1

    print("✅ Upload complete!")

    web_url = chunk_resp.json().get("webUrl")
    if web_url:
        print(f"File URL: {web_url}")   
    else:
        raise Exception("Upload succeeded but no webUrl returned.")


    github_env = os.getenv("GITHUB_ENV")
    if github_env:
        with open(github_env, "a") as f:
            f.write(f"FILE_URL={web_url}\n")

async def download_apk():    

    # Step 1: Locate Teams Channel Folder

    target_folder = "Android_APKs"
    folder = await graph_client.teams.by_team_id(team_id).channels.by_channel_id(channel_id).files_folder.get()
    drive_id = folder.parent_reference.drive_id
    folder_id = folder.id

     # Get target folder ID
    items = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(folder_id).children.get()
    target_folder_id = None
    for item in items.value:
        if item.name == target_folder and item.folder:
           target_folder_id = item.id
           break

    if not target_folder_id:
        raise Exception(f"Target folder '{target_folder}' not found in channel.")


    child_items = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(target_folder_id).children.get()
    
    if app_version_name is not None and app_version_code is not None:
        app_name = f"android-{app_env}-{app_version_name}({app_version_code}).apk"
        print(f"Looking for artifact: {app_name}")

        item_id = None
        for item in child_items.value:
            if item.name == app_name:
                item_id = item.id
                print(f"✅ Found artifact: {app_name}")
                break

        if not item_id:
            print("❌ Artifact not found, falling back to latest item")
            latest_item = max(
                child_items.value,
                key=lambda item: datetime.fromisoformat(str(item.last_modified_date_time))
            )
            item_id = latest_item.id
            print(f"➡️ Using latest artifact: {latest_item.name}")

    else: 
        print("Env variables missing, fetching latest artifact")
        latest_item = max(
            child_items.value,
            key=lambda item: datetime.fromisoformat(str(item.last_modified_date_time))
        )
        item_id = latest_item.id
        app_name = latest_item.name
        print(f"➡️ Using latest artifact: {latest_item.name}")


    #Step 2. Get the redirect CDN URL
    graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"

    headers = {"Authorization": f"Bearer {access_token}"}

    async with aiohttp.ClientSession() as session:
        async with session.get(graph_url, headers=headers, allow_redirects=False) as resp:
            if resp.status == 302:
                download_url = resp.headers["Location"]
            else:
                raise Exception(f"Unexpected response {resp.status}: {await resp.text()}")

    # Step 3. Figure out total size of file
    async with aiohttp.ClientSession() as session:
        async with session.head(download_url) as resp:
            total_size = int(resp.headers.get("Content-Length", 0))

    # Step 4. Resume logic
    start_byte = 0

    print(f"Starting download at byte {start_byte} of {total_size} into {app_name}")

    async with aiohttp.ClientSession() as session:
        with open(app_name, "ab") as f:
            while True:
                end_byte = start_byte + chunk_size - 1
                range_header = {"Range": f"bytes={start_byte}-{end_byte}"}
                async with session.get(download_url, headers=range_header) as resp:
                    if resp.status in (200, 206):  # 206 = Partial Content
                        chunk = await resp.content.read()
                        if not chunk:
                            break
                        f.write(chunk)
                        start_byte += len(chunk)

                        # Print progress
                        percent = (start_byte / total_size) * 100 if total_size else 0
                        print(f"Downloaded {start_byte}/{total_size} bytes ({percent:.2f}%)")

                    elif resp.status == 416:
                        # Requested range not satisfiable (EOF)
                        break
                    else:
                        raise Exception(f"Download failed: {resp.status} {await resp.text()}")

    print(f"✅ Download complete: {app_name}")

    github_env = os.getenv("GITHUB_ENV")
    if github_env:
        with open(github_env, "a") as f:
            f.write(f"APK_PATH={app_name}\n")

if action == "upload":
    asyncio.run(upload_to_teams())
else:
    asyncio.run(download_apk())
