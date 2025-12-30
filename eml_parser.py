import email
import os
from email.header import decode_header
from email.message import Message
from typing import Dict, Optional
import re


def decode_header_value(header_value: str) -> str:
    if not header_value:
        return ""

    decoded_parts = []
    for part, encoding in decode_header(header_value):
        if isinstance(part, bytes):
            if encoding:
                try:
                    decoded_parts.append(part.decode(encoding))
                except (UnicodeDecodeError, LookupError):
                    decoded_parts.append(part.decode('utf-8', errors='ignore'))
            else:
                decoded_parts.append(part.decode('utf-8', errors='ignore'))
        else:
            decoded_parts.append(str(part))

    return "".join(decoded_parts)


def get_email_body(message: Message) -> Optional[str]:
    body = None

    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))

            if content_type == "text/plain" and "attachment" not in content_disposition:
                charset = part.get_content_charset() or 'utf-8'
                payload = part.get_payload(decode=True)
                if payload:
                    try:
                        body = payload.decode(charset)
                        break
                    except (UnicodeDecodeError, LookupError):
                        try:
                            body = payload.decode('utf-8', errors='ignore')
                            break
                        except:
                            pass

        if body is None:
            for part in message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))

                if content_type == "text/html" and "attachment" not in content_disposition:
                    charset = part.get_content_charset() or 'utf-8'
                    payload = part.get_payload(decode=True)
                    if payload:
                        try:
                            html_body = payload.decode(charset)
                            body = html_to_text(html_body)
                            break
                        except (UnicodeDecodeError, LookupError):
                            try:
                                html_body = payload.decode('utf-8',
                                                           errors='ignore')
                                body = html_to_text(html_body)
                                break
                            except:
                                pass
    else:
        content_type = message.get_content_type()
        payload = message.get_payload(decode=True)

        if payload:
            charset = message.get_content_charset() or 'utf-8'
            try:
                body = payload.decode(charset)
            except (UnicodeDecodeError, LookupError):
                try:
                    body = payload.decode('utf-8', errors='ignore')
                except:
                    pass

    return body


def html_to_text(html: str) -> str:
    import re

    html = re.sub(r'<script[^>]*>.*?</script>',
                  '',
                  html,
                  flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<style[^>]*>.*?</style>',
                  '',
                  html,
                  flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<[^>]+>', ' ', html)
    html = re.sub(r'&nbsp;', ' ', html)
    html = re.sub(r'&lt;', '<', html)
    html = re.sub(r'&gt;', '>', html)
    html = re.sub(r'&amp;', '&', html)
    html = re.sub(r'&quot;', '"', html)
    html = re.sub(r'&#39;', "'", html)
    html = re.sub(r'\s+', ' ', html)

    return html.strip()


def parse_eml_file(file_path: str) -> Dict[str, str]:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"EML文件不存在: {file_path}")

    with open(file_path, 'rb') as f:
        message = email.message_from_bytes(f.read())

    subject = decode_header_value(message.get('Subject', ''))
    from_addr = decode_header_value(message.get('From', ''))
    date_str = message.get('Date', '')
    body = get_email_body(message)

    return {
        'file_path': file_path,
        'file_name': os.path.basename(file_path),
        'subject': subject,
        'from': from_addr,
        'date': date_str,
        'body': body or ''
    }


def extract_email_text_from_subject(subject: str) -> str:
    if not subject:
        return ""

    date_pattern = r'(\d{1,2})月(\d{1,2})日[^\d]*(\d{1,2})[:_](\d{1,2})'
    matches = re.findall(date_pattern, subject)

    location_pattern = r'(线上举行|F\d+|A\d+|B\d+|C\d+|会议室|报告厅|讲堂)'
    location_matches = re.findall(location_pattern, subject)

    result_lines = []
    result_lines.append(f"邮件主题: {subject}")

    if matches:
        for match in matches:
            result_lines.append(
                f"时间信息: {match[0]}月{match[1]}日 {match[2]}:{match[3]}")

    if location_matches:
        for loc in location_matches:
            result_lines.append(f"地点信息: {loc}")

    return "\n".join(result_lines)
