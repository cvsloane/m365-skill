#!/usr/bin/env python3
"""
MS365 CLI - Command-line interface for Microsoft 365 via MCP server.
Wraps the @softeria/ms-365-mcp-server for use with Clawdbot skills.
"""

import subprocess
import json
import sys
import argparse

def call_mcp(method: str, params: dict = None) -> dict:
    """
    Call the MCP server via stdio using JSON-RPC protocol.
    
    Args:
        method (str): The MCP tool method name to call
        params (dict, optional): Method parameters. Defaults to None.
        
    Returns:
        dict: Response from the MCP server, parsed as JSON or raw text
        
    Raises:
        subprocess.TimeoutExpired: If the MCP server doesn't respond within 60 seconds
        Exception: For other subprocess or JSON parsing errors
        
    Example:
        >>> result = call_mcp("list_messages", {"top": 5})
        >>> print(result)
        {'value': [{'id': 'msg1', 'subject': 'Hello'}]}
    """
    # Initialize request
    init_msg = json.dumps({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "ms365_cli", "version": "1.0"}
        }
    })

    # Tool call request
    call_msg = json.dumps({
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": method,
            "arguments": params or {}
        }
    })

    input_data = f"{init_msg}\n{call_msg}\n"

    try:
        result = subprocess.run(
            ["npx", "-y", "@softeria/ms-365-mcp-server"],
            input=input_data,
            capture_output=True,
            text=True,
            timeout=60
        )

        # Parse the second line (tool call response)
        lines = result.stdout.strip().split('\n')
        if len(lines) >= 2:
            response = json.loads(lines[1])
            if 'result' in response:
                content = response['result'].get('content', [])
                if content and len(content) > 0:
                    text = content[0].get('text', '')
                    # Try to parse as JSON for pretty printing
                    try:
                        data = json.loads(text)
                        return data
                    except:
                        return {"text": text}
            elif 'error' in response:
                return {"error": response['error']}

        return {"error": "Unexpected response", "raw": result.stdout}

    except subprocess.TimeoutExpired:
        return {"error": "Request timed out"}
    except Exception as e:
        return {"error": str(e)}

def format_output(data: dict, compact: bool = False):
    """
    Format and display MCP server response output.
    
    Args:
        data (dict): The response data from MCP server
        compact (bool, optional): Use compact JSON formatting. Defaults to False.
        
    Example:
        >>> format_output({"value": [{"id": "msg1"}]}, compact=True)
        {"value": [{"id": "msg1"}]}
    """
    if compact:
        print(json.dumps(data, indent=2))
    else:
        print(json.dumps(data, indent=2))

def cmd_login(args):
    """
    Initiate device code flow authentication.
    
    This command starts the interactive authentication process using Microsoft's
    device code flow. The user will be prompted to open a browser and enter
    a code to authenticate their Microsoft account.
    
    Args:
        args: argparse namespace (unused)
        
    Example:
        >>> python3 ms365_cli.py login
        Starting device code login...
        To sign in, use a web browser to open the page https://microsoft.com/devicelogin
        and enter the code XXXXXXXX to authenticate.
    """
    print("Starting device code login...")
    subprocess.run(["npx", "-y", "@softeria/ms-365-mcp-server", "--login"])

def cmd_status(args):
    """
    Check current authentication status.
    
    Verifies if the user is authenticated and can access Microsoft 365 services.
    Returns authentication status and account information.
    
    Args:
        args: argparse namespace (unused)
        
    Example:
        >>> python3 ms365_cli.py status
        {"authenticated": true, "account": "user@example.com"}
    """
    result = call_mcp("verify-login")
    format_output(result)

def cmd_accounts(args):
    """
    List all cached authentication accounts.
    
    Displays all Microsoft accounts that have been cached for authentication.
    Useful for managing multiple accounts.
    
    Args:
        args: argparse namespace (unused)
        
    Example:
        >>> python3 ms365_cli.py accounts
        {"accounts": [{"id": "acc1", "email": "user@example.com"}]}
    """
    result = call_mcp("list-accounts")
    format_output(result)

def cmd_user(args):
    """
    Get current authenticated user information.
    
    Retrieves details about the currently authenticated user including
    display name, email address, and other profile information.
    
    Args:
        args: argparse namespace (unused)
        
    Example:
        >>> python3 ms365_cli.py user
        {"user": {"displayName": "John Doe", "mail": "john@example.com"}}
    """
    result = call_mcp("get-current-user")
    format_output(result)

def cmd_mail_list(args):
    """
    List emails in the user's mailbox.
    
    Retrieves a list of emails from the specified folder or inbox.
    Supports pagination and folder filtering.
    
    Args:
        args: argparse namespace containing:
            top (int, optional): Maximum number of emails to return
            folder (str, optional): Folder ID to list emails from
            
    Example:
        >>> python3 ms365_cli.py mail list --top 10
        {"value": [{"id": "msg1", "subject": "Hello", "from": {"emailAddress": {"address": "sender@example.com"}}}]}
    """
    params = {}
    if args.top:
        params['top'] = args.top
    if args.folder:
        params['folderId'] = args.folder
    result = call_mcp("list-mail-messages", params)
    format_output(result)

def cmd_mail_read(args):
    """
    Read the full content of a specific email.
    
    Retrieves complete email details including subject, body, sender,
    recipients, and attachments for the specified message ID.
    
    Args:
        args: argparse namespace containing:
            id (str): The message ID of the email to read
            
    Example:
        >>> python3 ms365_cli.py mail read AAMkAG...=
        {"id": "AAMkAG...=", "subject": "Meeting Tomorrow", "body": {"content": "Let's meet at 2 PM..."}}
    """
    result = call_mcp("get-mail-message", {"messageId": args.id})
    format_output(result)

def cmd_mail_send(args):
    """
    Send a new email message.
    
    Creates and sends an email to the specified recipients with the given
    subject and body content.
    
    Args:
        args: argparse namespace containing:
            to (str): Recipient email address
            subject (str): Email subject line
            body (str): Email body content
            
    Example:
        >>> python3 ms365_cli.py mail send --to "john@example.com" --subject "Hello" --body "How are you?"
        {"id": "AAMkAG...=", "message": {"subject": "Hello"}}
    """
    body = {
        "message": {
            "subject": args.subject,
            "body": {
                "contentType": "Text",
                "content": args.body
            },
            "toRecipients": [
                {"emailAddress": {"address": args.to}}
            ]
        }
    }
    result = call_mcp("send-mail", {"body": body})
    format_output(result)

def cmd_calendar_list(args):
    """
    List calendar events.
    
    Retrieves upcoming calendar events from the user's calendar.
    Supports filtering by number of events and date ranges.
    
    Args:
        args: argparse namespace containing:
            top (int, optional): Maximum number of events to return
            
    Example:
        >>> python3 ms365_cli.py calendar list --top 5
        {"value": [{"id": "evt1", "subject": "Team Meeting", "start": {"dateTime": "2026-01-15T10:00:00Z"}}]}
    """
    params = {}
    if args.top:
        params['top'] = args.top
    result = call_mcp("list-calendar-events", params)
    format_output(result)

def cmd_calendar_create(args):
    """
    Create a new calendar event.
    
    Creates a calendar event with the specified subject, start/end times,
    optional description, and timezone.
    
    Args:
        args: argparse namespace containing:
            subject (str): Event subject/title
            start (str): Start time in ISO 8601 format
            end (str): End time in ISO 8601 format
            body (str, optional): Event description
            timezone (str, optional): Timezone (default: America/Chicago)
            
    Example:
        >>> python3 ms365_cli.py calendar create --subject "Team Meeting" --start "2026-01-15T10:00:00" --end "2026-01-15T11:00:00"
        {"id": "AAMkAG...=", "subject": "Team Meeting"}
    """
    body = {
        "subject": args.subject,
        "start": {
            "dateTime": args.start,
            "timeZone": args.timezone or "America/Chicago"
        },
        "end": {
            "dateTime": args.end,
            "timeZone": args.timezone or "America/Chicago"
        }
    }
    if args.body:
        body["body"] = {"contentType": "Text", "content": args.body}
    result = call_mcp("create-calendar-event", {"body": body})
    format_output(result)

def cmd_files_list(args):
    """
    List files and folders in OneDrive.
    
    Retrieves files and folders from the specified OneDrive location.
    Can list root directory or specific folder paths.
    
    Args:
        args: argparse namespace containing:
            path (str, optional): Folder path to list (default: root)
            
    Example:
        >>> python3 ms365_cli.py files list --path "Documents"
        {"value": [{"name": "report.pdf", "id": "file1", "file": {"mimeType": "application/pdf"}}]}
    """
    params = {"driveId": "me"}
    if args.path:
        params['driveItemId'] = args.path
    else:
        params['driveItemId'] = "root"
    result = call_mcp("list-folder-files", params)
    format_output(result)

def cmd_tasks_list(args):
    """
    List all To Do task lists.
    
    Retrieves all available task lists in the user's Microsoft To Do account.
    Each list can contain multiple tasks.
    
    Args:
        args: argparse namespace (unused)
        
    Example:
        >>> python3 ms365_cli.py tasks lists
        {"value": [{"id": "list1", "displayName": "Tasks", "isOwner": true}]}
    """
    result = call_mcp("list-todo-task-lists")
    format_output(result)

def cmd_tasks_get(args):
    """
    Get tasks from a specific task list.
    
    Retrieves all tasks from the specified task list, including
    completed and incomplete tasks.
    
    Args:
        args: argparse namespace containing:
            list_id (str): The ID of the task list
            
    Example:
        >>> python3 ms365_cli.py tasks get "list1"
        {"value": [{"id": "task1", "title": "Complete report", "status": "notStarted"}]}
    """
    result = call_mcp("list-todo-tasks", {"todoTaskListId": args.list_id})
    format_output(result)

def cmd_tasks_create(args):
    """
    Create a new task in a task list.
    
    Creates a new task item in the specified task list with an optional
    due date.
    
    Args:
        args: argparse namespace containing:
            list_id (str): The ID of the task list
            title (str): Task title/description
            due (str, optional): Due date in ISO 8601 format
            
    Example:
        >>> python3 ms365_cli.py tasks create "list1" --title "Review budget" --due "2026-01-20"
        {"id": "task1", "title": "Review budget", "status": "notStarted"}
    """
    body = {"title": args.title}
    if args.due:
        body["dueDateTime"] = {"dateTime": args.due, "timeZone": "America/Chicago"}
    result = call_mcp("create-todo-task", {
        "todoTaskListId": args.list_id,
        "body": body
    })
    format_output(result)

def cmd_contacts_list(args):
    """
    List contacts from Outlook.
    
    Retrieves contacts from the user's Outlook contacts folder.
    Supports pagination to limit the number of results.
    
    Args:
        args: argparse namespace containing:
            top (int, optional): Maximum number of contacts to return
            
    Example:
        >>> python3 ms365_cli.py contacts list --top 10
        {"value": [{"id": "contact1", "displayName": "John Doe", "emailAddresses": [{"address": "john@example.com"}]}]}
    """
    params = {}
    if args.top:
        params['top'] = args.top
    result = call_mcp("list-outlook-contacts", params)
    format_output(result)

def cmd_contacts_search(args):
    """
    Search for contacts by name or email.
    
    Searches the user's Outlook contacts using the specified query string.
    Matches against display names, email addresses, and other contact fields.
    
    Args:
        args: argparse namespace containing:
            query (str): Search query string
            
    Example:
        >>> python3 ms365_cli.py contacts search "John"
        {"value": [{"id": "contact1", "displayName": "John Doe", "emailAddresses": [{"address": "john@example.com"}]}]}
    """
    result = call_mcp("search-people", {"search": args.query})
    format_output(result)

def main():
    parser = argparse.ArgumentParser(description="MS365 CLI for Clawdbot")
    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # Auth commands
    login_p = subparsers.add_parser('login', help='Login via device code')
    login_p.set_defaults(func=cmd_login)

    status_p = subparsers.add_parser('status', help='Check auth status')
    status_p.set_defaults(func=cmd_status)

    accounts_p = subparsers.add_parser('accounts', help='List cached accounts')
    accounts_p.set_defaults(func=cmd_accounts)

    user_p = subparsers.add_parser('user', help='Get current user info')
    user_p.set_defaults(func=cmd_user)

    # Mail commands
    mail_p = subparsers.add_parser('mail', help='Mail commands')
    mail_sub = mail_p.add_subparsers(dest='mail_cmd')

    mail_list = mail_sub.add_parser('list', help='List emails')
    mail_list.add_argument('--top', type=int, default=10)
    mail_list.add_argument('--folder', help='Folder ID')
    mail_list.set_defaults(func=cmd_mail_list)

    mail_read = mail_sub.add_parser('read', help='Read email')
    mail_read.add_argument('id', help='Message ID')
    mail_read.set_defaults(func=cmd_mail_read)

    mail_send = mail_sub.add_parser('send', help='Send email')
    mail_send.add_argument('--to', required=True)
    mail_send.add_argument('--subject', required=True)
    mail_send.add_argument('--body', required=True)
    mail_send.set_defaults(func=cmd_mail_send)

    # Calendar commands
    cal_p = subparsers.add_parser('calendar', help='Calendar commands')
    cal_sub = cal_p.add_subparsers(dest='cal_cmd')

    cal_list = cal_sub.add_parser('list', help='List events')
    cal_list.add_argument('--top', type=int, default=10)
    cal_list.set_defaults(func=cmd_calendar_list)

    cal_create = cal_sub.add_parser('create', help='Create event')
    cal_create.add_argument('--subject', required=True)
    cal_create.add_argument('--start', required=True, help='ISO datetime')
    cal_create.add_argument('--end', required=True, help='ISO datetime')
    cal_create.add_argument('--body', help='Event description')
    cal_create.add_argument('--timezone', default='America/Chicago')
    cal_create.set_defaults(func=cmd_calendar_create)

    # Files commands
    files_p = subparsers.add_parser('files', help='OneDrive commands')
    files_sub = files_p.add_subparsers(dest='files_cmd')

    files_list = files_sub.add_parser('list', help='List files')
    files_list.add_argument('--path', help='Folder path')
    files_list.set_defaults(func=cmd_files_list)

    # Tasks commands
    tasks_p = subparsers.add_parser('tasks', help='To Do commands')
    tasks_sub = tasks_p.add_subparsers(dest='tasks_cmd')

    tasks_lists = tasks_sub.add_parser('lists', help='List task lists')
    tasks_lists.set_defaults(func=cmd_tasks_list)

    tasks_get = tasks_sub.add_parser('get', help='Get tasks from list')
    tasks_get.add_argument('list_id', help='Task list ID')
    tasks_get.set_defaults(func=cmd_tasks_get)

    tasks_create = tasks_sub.add_parser('create', help='Create task')
    tasks_create.add_argument('list_id', help='Task list ID')
    tasks_create.add_argument('--title', required=True)
    tasks_create.add_argument('--due', help='Due date (ISO format)')
    tasks_create.set_defaults(func=cmd_tasks_create)

    # Contacts commands
    contacts_p = subparsers.add_parser('contacts', help='Contacts commands')
    contacts_sub = contacts_p.add_subparsers(dest='contacts_cmd')

    contacts_list = contacts_sub.add_parser('list', help='List contacts')
    contacts_list.add_argument('--top', type=int, default=20)
    contacts_list.set_defaults(func=cmd_contacts_list)

    contacts_search = contacts_sub.add_parser('search', help='Search contacts')
    contacts_search.add_argument('query', help='Search query')
    contacts_search.set_defaults(func=cmd_contacts_search)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    if hasattr(args, 'func'):
        args.func(args)
    else:
        parser.print_help()
        sys.exit(1)

if __name__ == '__main__':
    main()
