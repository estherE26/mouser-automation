"""
Parses Jira ticket description to extract press release details.
"""
import re
from typing import Optional


def parse_jira_ticket(description: str) -> dict:
    """
    Extract key fields from Jira ticket description.

    Expected fields:
    - FILES ON SERVER: folder path in Dropbox
    - LINK EMBEDDED IMAGE TO: tracking URL for image
    - EMAIL SUBJECT LINE: email subject

    Returns dict with: folder_path, image_url, subject
    """
    result = {
        'folder_path': None,
        'image_url': None,
        'subject': None,
        'raw_description': description
    }

    # Extract FILES ON SERVER (folder path)
    folder_match = re.search(
        r'FILES ON SERVER:\s*(.+?)(?:\n|$)',
        description,
        re.IGNORECASE
    )
    if folder_match:
        result['folder_path'] = folder_match.group(1).strip()

    # Extract LINK EMBEDDED IMAGE TO (tracking URL)
    image_url_match = re.search(
        r'LINK EMBEDDED IMAGE TO:\s*(.+?)(?:\n|$)',
        description,
        re.IGNORECASE
    )
    if image_url_match:
        result['image_url'] = image_url_match.group(1).strip()

    # Extract EMAIL SUBJECT LINE
    subject_match = re.search(
        r'EMAIL SUBJECT LINE:\s*(.+?)(?:\n|$)',
        description,
        re.IGNORECASE
    )
    if subject_match:
        result['subject'] = subject_match.group(1).strip()

    return result


def validate_parsed_data(data: dict) -> tuple[bool, list[str]]:
    """
    Validate that required fields were extracted.
    Returns (is_valid, list of missing fields).
    """
    missing = []

    if not data.get('folder_path'):
        missing.append('FILES ON SERVER (folder path)')

    if not data.get('image_url'):
        missing.append('LINK EMBEDDED IMAGE TO (tracking URL)')

    if not data.get('subject'):
        missing.append('EMAIL SUBJECT LINE')

    return len(missing) == 0, missing


def extract_ticket_info_from_webhook(payload: dict) -> dict:
    """
    Extract ticket information from Jira webhook payload.

    Jira Automation webhooks typically include:
    - issue.key: Ticket key (e.g., "MW-123")
    - issue.fields.summary: Ticket title
    - issue.fields.description: Ticket description
    """
    issue = payload.get('issue', {})
    fields = issue.get('fields', {})

    ticket_info = {
        'key': issue.get('key', 'Unknown'),
        'title': fields.get('summary', ''),
        'description': fields.get('description', '') or ''
    }

    # Parse the description to extract our fields
    parsed = parse_jira_ticket(ticket_info['description'])

    return {
        **ticket_info,
        **parsed
    }
