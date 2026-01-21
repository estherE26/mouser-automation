"""
Slack notification module for press release automation.
"""
import os
import requests
from typing import Optional


def send_slack_notification(
    message: str,
    webhook_url: Optional[str] = None,
    blocks: Optional[list] = None
) -> bool:
    """
    Send a notification to Slack via webhook.

    Args:
        message: Plain text message (used as fallback)
        webhook_url: Slack webhook URL (defaults to env var)
        blocks: Optional Slack Block Kit blocks for rich formatting

    Returns:
        True if successful, False otherwise
    """
    webhook_url = webhook_url or os.environ.get('SLACK_WEBHOOK_URL')
    if not webhook_url:
        print("Warning: No Slack webhook URL configured")
        return False

    payload = {"text": message}
    if blocks:
        payload["blocks"] = blocks

    try:
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        return True
    except Exception as e:
        print(f"Failed to send Slack notification: {e}")
        return False


def notify_press_release_ready(
    ticket_key: str,
    folder_name: str,
    preview_urls: dict,
    webhook_url: Optional[str] = None
) -> bool:
    """
    Send formatted notification that press release is ready for review.

    Args:
        ticket_key: Jira ticket key (e.g., "MW-123")
        folder_name: Name of the processed folder
        preview_urls: Dict with 'html' and 'email' URLs
        webhook_url: Optional webhook URL override

    Returns:
        True if successful
    """
    html_url = preview_urls.get('html', 'N/A')
    email_url = preview_urls.get('email', 'N/A')

    blocks = [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": "✅ Press Release Ready for Review",
                "emoji": True
            }
        },
        {
            "type": "section",
            "fields": [
                {
                    "type": "mrkdwn",
                    "text": f"*Jira Ticket:*\n{ticket_key}"
                },
                {
                    "type": "mrkdwn",
                    "text": f"*Folder:*\n{folder_name}"
                }
            ]
        },
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": "*Preview Links:*"
            }
        },
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f"• <{html_url}|Web Version>\n• <{email_url}|Email Version>"
            }
        },
        {
            "type": "context",
            "elements": [
                {
                    "type": "mrkdwn",
                    "text": "Review and send URLs to Aaron when ready."
                }
            ]
        }
    ]

    fallback_message = (
        f"Press Release Ready: {ticket_key}\n"
        f"Folder: {folder_name}\n"
        f"Web: {html_url}\n"
        f"Email: {email_url}"
    )

    return send_slack_notification(fallback_message, webhook_url, blocks)


def notify_error(
    ticket_key: str,
    error_message: str,
    webhook_url: Optional[str] = None
) -> bool:
    """
    Send error notification to Slack.

    Args:
        ticket_key: Jira ticket key
        error_message: Description of the error
        webhook_url: Optional webhook URL override

    Returns:
        True if successful
    """
    blocks = [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": "❌ Press Release Processing Failed",
                "emoji": True
            }
        },
        {
            "type": "section",
            "fields": [
                {
                    "type": "mrkdwn",
                    "text": f"*Jira Ticket:*\n{ticket_key}"
                }
            ]
        },
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f"*Error:*\n```{error_message}```"
            }
        },
        {
            "type": "context",
            "elements": [
                {
                    "type": "mrkdwn",
                    "text": "Please check the ticket and try manual processing."
                }
            ]
        }
    ]

    fallback_message = (
        f"Press Release FAILED: {ticket_key}\n"
        f"Error: {error_message}"
    )

    return send_slack_notification(fallback_message, webhook_url, blocks)
