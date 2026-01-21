"""
Vercel Serverless Function - Jira Webhook Handler
Receives Jira webhook, downloads files from Dropbox, processes press release,
uploads to FTP, and sends Slack notification.
"""

import os
import sys
import json
import tempfile
import shutil
from http.server import BaseHTTPRequestHandler

# Add lib directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'lib'))

from jira_parser import extract_ticket_info_from_webhook, validate_parsed_data
from dropbox_client import DropboxClient
from slack_notify import notify_press_release_ready, notify_error
from press_release import process_press_release, upload_to_ftp, DEFAULT_CONFIG


def load_template(name: str) -> str:
    """Load an HTML template file."""
    template_dir = os.path.join(os.path.dirname(__file__), '..', 'templates')
    template_path = os.path.join(template_dir, name)
    with open(template_path, 'r', encoding='utf-8') as f:
        return f.read()


def handler(request):
    """
    Main webhook handler for Jira automation.

    Expects POST request with Jira webhook payload containing:
    - issue.fields.description with:
      - FILES ON SERVER: folder path in Dropbox
      - LINK EMBEDDED IMAGE TO: tracking URL
      - EMAIL SUBJECT LINE: email subject
    """
    # Only accept POST requests
    if request.method != 'POST':
        return {
            'statusCode': 405,
            'body': json.dumps({'error': 'Method not allowed'})
        }

    # Parse webhook payload
    try:
        body = request.body
        if isinstance(body, bytes):
            body = body.decode('utf-8')
        payload = json.loads(body) if body else {}
    except json.JSONDecodeError as e:
        return {
            'statusCode': 400,
            'body': json.dumps({'error': f'Invalid JSON: {e}'})
        }

    # Extract ticket info
    ticket_info = extract_ticket_info_from_webhook(payload)
    ticket_key = ticket_info.get('key', 'Unknown')

    # Validate required fields
    is_valid, missing_fields = validate_parsed_data(ticket_info)
    if not is_valid:
        error_msg = f"Missing required fields: {', '.join(missing_fields)}"
        notify_error(ticket_key, error_msg)
        return {
            'statusCode': 400,
            'body': json.dumps({'error': error_msg, 'ticket': ticket_key})
        }

    folder_path = ticket_info['folder_path']
    image_url = ticket_info['image_url']
    subject = ticket_info['subject']

    # Create temp directory for processing
    temp_dir = tempfile.mkdtemp(prefix='mouser_pr_')

    try:
        # Download files from Dropbox
        dropbox = DropboxClient()

        # The folder path from Jira might be relative (e.g., "Mouser/2026-01-19_PR_Name")
        # or just the folder name. Try to find it.
        dropbox_path = folder_path
        if not dropbox_path.startswith('/'):
            dropbox_path = '/' + dropbox_path

        # Try direct path first
        try:
            local_folder = dropbox.download_folder(dropbox_path, temp_dir)
        except Exception:
            # Try searching for folder by name
            folder_name = os.path.basename(folder_path)
            found_path = dropbox.find_folder_by_name(folder_name)
            if found_path:
                local_folder = dropbox.download_folder(found_path, temp_dir)
            else:
                raise Exception(f"Could not find folder: {folder_path}")

        # Load templates
        press_release_template = load_template('press_release.html')
        email_template = load_template('email.html')

        # Process press release
        result = process_press_release(
            folder_path=local_folder,
            press_release_template=press_release_template,
            email_template=email_template,
            image_url=image_url,
            subject=subject
        )

        if not result['success']:
            error_msg = '; '.join(result['errors'])
            notify_error(ticket_key, error_msg)
            return {
                'statusCode': 500,
                'body': json.dumps({
                    'error': 'Processing failed',
                    'details': result['errors'],
                    'ticket': ticket_key
                })
            }

        # Upload to FTP
        upload_result = upload_to_ftp(
            result['files_to_upload'],
            DEFAULT_CONFIG,
            result['folder_name'],
            result['month_folder']
        )

        if not upload_result['success']:
            error_msg = f"FTP upload failed: {upload_result['error']}"
            notify_error(ticket_key, error_msg)
            return {
                'statusCode': 500,
                'body': json.dumps({
                    'error': 'Upload failed',
                    'details': upload_result['error'],
                    'ticket': ticket_key
                })
            }

        # Send success notification to Slack
        notify_press_release_ready(
            ticket_key=ticket_key,
            folder_name=result['folder_name'],
            preview_urls=result['preview_urls']
        )

        return {
            'statusCode': 200,
            'body': json.dumps({
                'success': True,
                'ticket': ticket_key,
                'folder': result['folder_name'],
                'preview_urls': result['preview_urls'],
                'files_uploaded': upload_result['uploaded']
            })
        }

    except Exception as e:
        error_msg = str(e)
        notify_error(ticket_key, error_msg)
        return {
            'statusCode': 500,
            'body': json.dumps({
                'error': 'Unexpected error',
                'details': error_msg,
                'ticket': ticket_key
            })
        }

    finally:
        # Clean up temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


# Vercel Python runtime handler
class Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)

        class Request:
            def __init__(self, method, body):
                self.method = method
                self.body = body

        result = handler(Request('POST', body))

        self.send_response(result['statusCode'])
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(result['body'].encode())

    def do_GET(self):
        # Health check endpoint
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({
            'status': 'ok',
            'service': 'mouser-press-release-automation'
        }).encode())
