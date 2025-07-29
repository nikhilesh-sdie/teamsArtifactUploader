import os
import asyncio
import requests
from msgraph import GraphServiceClient
from msal import ConfidentialClientApplication
from azure.identity.aio import ClientSecretCredential

# ==== CONFIG ====
client_id = os.getenv("client_id")
tenant_id = os.getenv("tenant_id")
client_secret = os.getenv("client_secret")
team_id = os.getenv("team_id")
channel_id = os.getenv("channel_id")
file_path = os.getenv("file_path", "android/app/build/outputs/apk/production/release/app-production-release.apk")
file_name = os.getenv("name", os.path.basename(file_path))
chunk_size = 5 * 1024 * 1024  # 5 MB
target_folder = "Android_APKs"

# ==== AUTH ====
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
        with open(os.environ['GITHUB_OUTPUT'], 'a') as fh:
            fh.write(f"Application={web_url}\n")
    else:
        raise Exception("Upload succeeded but no webUrl returned.")

asyncio.run(upload_to_teams())