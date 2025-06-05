# Mailbox Deduplicator for Microsoft 365

This is a Python-based command-line tool for scanning a Microsoft 365 user's mailbox and identifying duplicate emails based on a normalized hash of the subject, body, and links. Optionally, it can delete these duplicates in bulk using Microsoft Graph's batch API.

## Features 

- Uses Microsoft Graph API with delegated authentication

- Identifies duplicates based on hashed email content

- Supports dry-run mode (default) for safe testing

- Writes duplicates to CSV per folder for auditing

- Supports batch deletion via -delete flag

- Handles large mailboxes and rate limiting

## Requirements

- Python 3.10+

- Microsoft 365 Business Premium (or higher)

- App Registration in Azure AD (see below)

- Delegated API permissions: Mail.ReadWrite, Mail.ReadWrite.Shared

### Install dependencies:

**For Windows**
```
pip install -r requirements.txt
```

**For Linux**
```
python3 -m pip install -r requirements.txt
```
### Setting Up Azure App Registration
1. Go to the [Entra Portal](https://entra.microsoft.com/#home)
2. Navigate to: Applications -> App registrations -> New registration
3. Register app<br />
    - Name: Mailbox Deduplicator
    - Supported account types: Accounts in this organizational directory only
    - Redirect URI:
        - Public client/native (mobile & desktop)
        - http://localhost

**After registering**
1. Copy Application (client) ID
2. Copy Directory (tenant) ID
3. Click "API permissions" -> "Add a permission" -> "Microsoft Graph" -> "Delegated permissions"
    - Add in 'Mail.ReadWrite"
    - Add in "Mail.ReadWrite.Shared"
## Creating Your .env File
### Example .env:
```
AZURE_TENANT_ID="<your-tenant-id>"
DEDUPE_CLIENT_ID="<your-client-id>"
```
Save it as .env or another file, and provide the path using the -f argument.

## Usage
```
python main.py -u user@domain.com -f .env [-delete]
```
### Arguments:
  - -u, -user: Target user's email address
  - -f, -file: Path to your .env file with Azure credentials
  - -delete: If provided, the script will delete duplicates instead of just logging them

### Example:
***Dry Run Only***
```
python main.py -u jsmith@contoso.com -f ./config.env
```
***Delete confirmed dupes***
```
python main.py -u jsmith@contoso.com -f ./config.env -delete
```
## Output
Duplicates are logged per folder to ./output/<user>-<folder>.csv The CSV contains the Received timestamp, Subject, and Body of the message.

## Notes/Tips
  - Delete is set to false by default, and the -delete flag must be added to perform deletion of items from the mail server.
  - It is highly recommended that you do a run without delete just to make sure that what you expect to be deleted, is what is being deleted.
  - The oldest message per hash is kept and the newer duplicates are deleted.
  - It is possible for very technically different emails to be deleted such as automated reminders from a reminder - we hash links within these messages to try and combat false positives but if an automated message is sending you the same link then technically newer emails are still going to be deleted erroneously.
  - This should work on large mail boxes with tens of thousands of messages.

