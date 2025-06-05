from hashlib import sha256
from re import sub, findall, IGNORECASE
from os import mkdir, getenv
from os.path import isdir, join, exists
from argparse import ArgumentParser, ArgumentError, ArgumentTypeError, Namespace
from time import sleep
from datetime import datetime
from csv import DictWriter
from typing import Generator, Any
from requests import Response, get, post
from msal import PublicClientApplication
from rich.console import Console
from rich.theme import Theme
from rich.progress import track
from dotenv import load_dotenv
from mail_folder import MailFolder
from mail_message import MailMessage

# Microsoft Graph URLs
GET_MAILBOX_FOLDERS_URL: str = r'https://graph.microsoft.com/v1.0/users/{{user_id}}/mailFolders?$select=id,displayName,totalItemCount'
GET_FOLDER_MESSAGE_URL: str = r'https://graph.microsoft.com/v1.0/users/{{user_id}}/mailFolders/{{folder_id}}/messages?$select=createdDateTime,id,body,subject,receivedDateTime'
BATCH_REQUEST_URL: str = r'https://graph.microsoft.com/v1.0/$batch'

# Our Rich Text styles
STYLES: dict[str, str] = {'success': 'green', 'caution': 'yellow', 'error': 'red', 'info': 'cyan'}

def main() -> None:
    """Main entry point of the program that handles over-arching logic.
    """
    console: Console = create_console()
    args: dict[str, str] | None = get_args(console)

    if args:
        tenant_id, client_id = get_env_data(args['file'], console)

        if tenant_id and client_id:
            azure_auth_token: str = get_delegated_token(tenant_id, client_id, console)
            scan_and_optionally_delete(azure_auth_token, args['user'], console, args['delete'])
        else:
            console.print(f'Error reading in environment variables. Please ensure you have values set for "AZURE_TENANT_ID" and "DEDUPE_CLIENT_ID"', style='error')

def create_console() -> Console:
    """Creates a rich-text console.

    Returns:
        Console: Returns an instantiated rich-text console object.
    """
    return Console(theme=Theme(STYLES))

def get_args(console: Console) -> dict[str, str]:
    """Gets arguments from the command line.
    
    Args:
        console (Console): Rich-text console for pretty print statements.

    Returns:
        dict[str, str]: A dictionary with file, user, and delete keys.
    """
    parser: ArgumentParser = ArgumentParser()

    parser.add_argument('-u', '-user', help='The user whose mailbox we should deduplicate.', required=True)
    parser.add_argument('-d', '-delete', help='Deletes any duplicates that are found.', action='store_true')
    parser.add_argument('-f', '-file', help="""The env file that contains the tenant ID and client ID. Should contain the following entries:\n
                        CLIENT_ID=\"<CLIENT ID>\"\n
                        TENANT_ID=\"<TENANT ID>\"""", 
                        required=True)
    
    return _validate_args(parser, console)

def _validate_args(parser: ArgumentParser, console: Console) -> dict[str, str]:
    """Validates and retrieves arguments from the argument parser.

    Args:
        parser (ArgumentParser): An instantiated parser from which arguments can be pulled.
        console (Console): Rich-text console for pretty print statements.

    Returns:
        dict[str, str]: A dictionary with file, user, and delete keys.
    """
    parsed_args: dict[str, str] | None = None

    try:
        args: Namespace = parser.parse_args()

        # Were both arguments provided and does the file exist?
        if (args.f and args.u) and exists(args.f):
            parsed_args = {
                'file': args.f,
                'user': args.u,
                'delete': args.d
            }
        else:
            parser.error('Valid arguments not specified.')
    except ArgumentError as err:
        console.print(f'{err.argument_name}: {err.message}', style='error')
        console.print('\nPlease specify -h flag for help for more information.', style='error')
    finally:
        return parsed_args

def get_env_data(env_file: str, console: Console) -> tuple[str | None, str | None]:
    """Retrieves the tenant id and client id variables from ENV file.

    Args:
        env_file (str): The ENV file that should be read from.
        console (Console): Rich-text console for pretty print statements.

    Returns:
        tuple[str | None, str | None]: Returns up to two strings from the ENV file if they were found. Otherwise will return one or more null variables.
    """
    tenant_id: str | None = None
    client_id: str | None = None

    try:
        load_dotenv(dotenv_path=env_file)
        tenant_id = getenv('AZURE_TENANT_ID') or None
        client_id = getenv('DEDUPE_CLIENT_ID') or None
    except Exception as err:
        console.print(f'Error loading env file: {err}')
    finally:
        return (tenant_id, client_id)

def normalize_text(text: str) -> str:
    """Normalizes text for consistent formatting by removing HTML tags, converting to lower case, and stripping leading and trailing white space.

    Args:
        text (str): The text to normalize.

    Returns:
        str: A normalized version of the string.
    """
    text = sub(r'<[^>]+>', '', text)
    text = sub(r'\s+', ' ', text)
    return text.strip().lower()

def extract_links(text: str) -> list[str]:
    """Extracts all hyperlinks from text.

    Args:
        text (str): The text to extract hyperlinks from.

    Returns:
        list[str]: Returns a list of all the hyperlinks that have been extracted.
    """
    # Match href attributes with either single or double quotes
    return findall(r'href=["\'](.*?)["\']', text, flags=IGNORECASE)

def create_headers(token: str, include_app_type: bool = False) -> dict[str, str]:
    """Creates API payload headers.

    Args:
        token (str): M365 Authorization token.
        include_app_type (bool, optional): Determines if application type json is included in the headers. Defaults to False.

    Returns:
        dict[str, str]: Authorization and/or application header information.
    """
    return {
        'Authorization': f'Bearer {token}'
    } if not include_app_type else {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

def compute_hash(subject: str, body: str) -> str:
    """Generates a SHA256 hash of an email by combining the subject, body, and any links found within the mail message.

    Args:
        subject (str): The mail message subject.
        body (str): The mail message body.

    Returns:
        str: A SHA256 hash of the mail message.
    """
    clean_subject: str = normalize_text(subject or '')
    clean_body: str = normalize_text(body or '')
    link_text: str = ''.join(extract_links(body))
    return sha256(f'{clean_subject}|{clean_body}|{link_text}'.encode('utf-8')).hexdigest()

# MSAL: Get delegated access token
def get_delegated_token(tenant_id: str, client_id: str, console: Console) -> str:
    """Retrieves an access token that allows a user to perform actions on mail boxes in a M365 tenant.

    Args:
        tenant_id (str): The ID of the M365 tenant we're trying to access.
        client_id (str): The client ID of the application that will grant us permissions to access mail boxes.
        console (Console): Rich-text console for pretty print statements.

    Raises:
        Exception: Generic error if an access token was unable to be acquired.

    Returns:
        str: The access token that grants permission for all mailbox related actions.
    """
    console.print(r'Acquiring authorization to perform mailbox actions. You should see a browser window open and ask you to login.', style='info')

    scopes: list[str] = ['Mail.ReadWrite', 'Mail.ReadWrite.Shared']
    authority_url: str = f'https://login.microsoftonline.com/{tenant_id}'
    app: PublicClientApplication = PublicClientApplication(client_id, authority=authority_url)
    result: dict[str, Any] | dict[str, str] | dict | Any | dict[Any | str, Any] = app.acquire_token_interactive(scopes=scopes)

    if 'access_token' in result:
        console.print('Token successfully acquired!\n', style='success')
        return result['access_token']
    else:
        raise Exception('Authentication failed.')

def soft_graph_call(url: str, headers: dict[str, str], max_tries: int = 5, json_payload: dict[str, Any] = None) -> Response:
    """Attempts to gracefully reach out to an endpoint and handle various status codes.

    Args:
        url (str): The target endpoint.
        headers (dict[str, str]): The payload headers that should be used.
        max_tries (int, optional): The maximum number of retries before giving up if something goes wrong. Defaults to 5.
        json_payload (dict[str, Any], optional): Optional data that can be sent along to the endpoint. Defaults to None.

    Raises:
        Exception: Raises generic exception if a non-200 status is received to the point max retries has been exhausted.

    Returns:
        Response: Returns a Response object from the requests module if successful.
    """
    try_count: int = 0

    while try_count < max_tries:
        response: Response = get(url=url, headers=headers,timeout=(10,60)) if json_payload is None else post(
            url=url, headers=headers, timeout=(10, 60), json=json_payload)

        match(response.status_code):
            case 200: # Everything is okay
                return response
            
            case 429: # We're getting throttled
                wait_for: int = int(response.headers.get('Retry-After', '5'))
                sleep(wait_for)

            case 502 | 503 | 504: # Transient error
                wait_for: int = (2 ** try_count)
                sleep(wait_for)

            case _:
                print(f'Unhandled error {response.status_code} during graph call to {url}')

        try_count += 1

    raise Exception(f'Microsoft graph call failed after {max_tries} tries.')

def get_mail_folders(token: str, target_user: str) -> list[MailFolder]:
    """Retrieves the target user's mail folders.

    Args:
        token (str): M365 Authorization token.
        target_user (str): Target user whose mailbox is being scanned.

    Returns:
        list[MailFolder]: A list of the user's mail folders from their mailbox.
    """
    headers: dict[str, str] = create_headers(token) 

    url: str | None = GET_MAILBOX_FOLDERS_URL.replace(r'{{user_id}}', target_user)

    all_folders: list[MailFolder] = list()

    while url:
        response: Response = soft_graph_call(url, headers)
        response.raise_for_status()
        data: dict[str, Any] = response.json()
        url: str | None = data.get('@odata.nextLink', None)
        parsed_data: list[MailFolder] = [MailFolder(d['id'], d['displayName'], d['totalItemCount']) for d in data['value']]
        all_folders.extend(parsed_data)

    return all_folders

def fetch_messages(token: str, target_user: str, target_folder: str) -> Generator[MailMessage, None, None]:
    """Fetches mail messages from a specified folder from the target user.

    Args:
        token (str): M365 Authorization token.
        target_user (str): Target user whose mailbox is being scanned.
        target_folder (str): Target mail folder that should be scanned.

    Yields:
        Generator[MailMessage, None, None]: A single mail message at a time for processing.
    """
    headers: dict[str, str] = create_headers(token)

    url: str = GET_FOLDER_MESSAGE_URL.replace(r'{{user_id}}', target_user).replace(r'{{folder_id}}', target_folder)

    while url:
        response: Response = soft_graph_call(url, headers)
        response.raise_for_status()
        data: dict[str, str] | dict[str, list[Any]] = response.json()
        message_batch: list[dict[str, Any]] = data.get('value', [])
        url = data.get('@odata.nextLink', None)

        for message in message_batch:
            yield MailMessage(
                id=message.get('id'),
                subject=message.get('subject', ''),
                body=message.get('body', {}).get('content', ''),
                received=message.get('receivedDateTime', '')
            )

def scan_and_optionally_delete(token: str, user: str, console: Console, should_delete: bool=False):
    """Scans through all of the mail folders of a user in search of duplicate emails. Duplicates are recorded and output to CSVs. Optionally, the emails can be deleted from the mail server.

    Args:
        token (str): M365 Authorization token.
        user (str): Target user whose mailbox is being scanned.
        console (Console): Rich-text console for pretty print statements.
        should_delete (bool, optional): Flag for deleting the emails on the mail server or not. Defaults to False.
    """
    folders: list[MailFolder] = get_mail_folders(token, user)

    for folder in track(folders, description='Mailbox folder progress...'):
        if folder.total_count > 0:
            console.print(f'Scanning {folder.display_name}...', style='info')
            hash_map: dict[str, list[MailMessage]] = dict()

            for message in fetch_messages(token, user, folder.id):
                message_hash: str = compute_hash(message.subject, message.body)

                if message_hash not in hash_map:
                    hash_map[message_hash] = list()

                hash_map[message_hash].append(message)

            console.print('\nCollected messages.', style='success')

            duplicates: list[MailMessage] = list()

            console.print('Sorting items.', style='info')
            for hash, group in hash_map.items():
                if len(group) > 1:
                    # Sort by receivedDateTime, oldest first
                    group_sorted = sorted(group, key=lambda msg: datetime.fromisoformat(msg.received.replace('Z', '+00:00')))
                    # Keep the oldest one, mark the rest for deletion
                    duplicates.extend(group_sorted[1:])

            console.print('Sorting complete.', style='success')

            if len(duplicates) > 0:
                console.print('Dumping duplicates to csv.', style='info')
                write_dupes_to_csv(f'{user}-{folder.display_name}.csv', duplicates)

                if should_delete:
                    console.print('Deleting emails.', style='caution')
                    delete_messages_batched(token, duplicates, user, folder.id, console)
        else:
            console.print(f'\nSkipping {folder.display_name} since it has 0 items.', style='caution')    

def sanitize_file_name(file_name: str, invalid_chars: set[str, str] = {'\\', '/', '\'', ','} ) -> str:
    """Removes illegal characters from file names that might prevent issues when attempting to output to a file.

    Args:
        file_name (str): The file name that should be sanitized. 
        invalid_chars (set[str, str]): The set of characters to remove from the file name. Defaults to {'\', '/', '\'', ','}.

    Returns:
        str: The file name with '-' in place of any invalid characters.
    """
    translation_table: dict[int, str] = {ord(char): '-' for char in invalid_chars}
    return str.translate(file_name, translation_table)

def write_dupes_to_csv(file: str, duplicates: list[MailMessage]) -> None:
    """Outputs every duplicate in a mailbox to a CSV as an audit trail and backup to prevent accidental deletions of critical emails due to misidentification.

    Args:
        file (str): The file that should be output to.
        duplicates (list[MailMessage]): List of all the duplicate emails that are about to be deleted.
    """
    target_folder: str = 'output'

    # Do we have our output folder?
    if not isdir(target_folder):
        mkdir(target_folder)

    output_path: str = join(target_folder, sanitize_file_name(file))

    with open(output_path, 'w', newline='', encoding='utf8') as file_stream:
        writer: DictWriter = DictWriter(file_stream, ['Received', 'Subject', 'Body'])
        writer.writeheader()
        writer.writerows([{'Received': dupe.received, 'Subject': dupe.subject, 'Body': dupe.body} for dupe in duplicates])

def delete_messages_batched(token: str, message_ids: list[MailMessage], user_id: str, folder_id: str, console: Console, batch_size: int = 20) -> None:
    """Deletes emails from the inbox in batches.

    Args:
        token (str): M365 Authorization token.
        message_ids (list[MailMessage]): List of messages to delete.
        user_id (str): Target user whose mailbox is being scanned.
        folder_id (str): The folder containing the messages we're about to delete.
        console (Console): Rich-text console for pretty print statements.
        batch_size (int, optional): How many emails we want to delete in a single batch. Each email in the batch is considered an API call. Defaults to 20 which is Microsoft's current batch limit.
    """
    headers: dict[str, str] = create_headers(token, True)

    # Split into batches of batch_size (at the time of this script, Microsoft limits this to a maximum of 20)
    for i in range(0, len(message_ids), batch_size):
        batch: list[MailMessage] = message_ids[i:i+batch_size]
        requests: list[dict[str, str]] = [
            {
                'id': str(batch_id),
                'method': 'DELETE',
                'url': f'/users/{user_id}/mailFolders/{folder_id}/messages/{message.id}'
            } for batch_id, message in enumerate(batch, start=1)
        ]

        batch_payload = { 'requests': requests }

        response: Response = soft_graph_call(
            BATCH_REQUEST_URL,
            headers=headers,
            json_payload=batch_payload
        )

        result: list[dict[str, Any]] = response.json()
        for item in result.get('responses', []):
            if item['status'] != 204:
                console.print(f"Failed: {item['id']} - Status: {item['status']}", style='error')

# Main execution
if __name__ == '__main__':
    main()