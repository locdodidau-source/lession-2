from __future__ import print_function
import os
import datetime as dt
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Phạm vi quyền (Google Calendar)
SCOPES = ['https://www.googleapis.com/auth/calendar']


# ----------------------------------------
# Đăng nhập Google và tạo service
# ----------------------------------------
def dang_nhap_google():
    """
    Đăng nhập Google bằng OAuth2, trả về đối tượng service để thao tác Calendar.
    Yêu cầu có file credentials.json trong cùng thư mục.
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    print("✅ Đăng nhập Google thành công!")
    service = build('calendar', 'v3', credentials=creds)
    return service


# ----------------------------------------
# Tạo sự kiện lặp hàng tuần
# ----------------------------------------
def tao_su_kien(service, mon, phong, giang_vien,
                start_date, end_date, weekday, start_time, end_time,
                reminders=None, prefix="[TKB]"):
    """
    Tạo sự kiện lặp hàng tuần trên Google Calendar.

    Args:
        service: đối tượng Google Calendar service.
        mon (str): tên môn học.
        phong (str): phòng học.
        giang_vien (str): tên giảng viên.
        start_date (str): ngày bắt đầu (dd/mm/YYYY).
        end_date (str): ngày kết thúc (dd/mm/YYYY).
        weekday (int): số thứ (2=Thứ 2 ... 7=Thứ 7).
        start_time (str): giờ bắt đầu (HH:MM).
        end_time (str): giờ kết thúc (HH:MM).
        reminders (list[dict]): danh sách nhắc nhở, ví dụ:
            [{"method": "popup", "minutes": 10}, {"method": "email", "minutes": 1440}]
        prefix (str): tiền tố cho tiêu đề sự kiện.
    Returns:
        str: event_id của sự kiện vừa tạo.
    """
    # Đổi string -> datetime
    start_date = dt.datetime.strptime(start_date.strip(), "%d/%m/%Y").date()
    end_date = dt.datetime.strptime(end_date.strip(), "%d/%m/%Y").date()

    # Google Calendar: 0=Thứ 2 ... 6=Chủ Nhật
    google_weekday = weekday - 2
    if google_weekday < 0:
        google_weekday = 6

    # Tìm ngày bắt đầu khớp với thứ học
    current = start_date
    while current.weekday() != google_weekday:
        current += dt.timedelta(days=1)

    start_dt = dt.datetime.strptime(f"{current.strftime('%d/%m/%Y')} {start_time}", "%d/%m/%Y %H:%M")
    end_dt = dt.datetime.strptime(f"{current.strftime('%d/%m/%Y')} {end_time}", "%d/%m/%Y %H:%M")

    # Body sự kiện
    event = {
        'summary': f"{prefix} {mon}",
        'location': phong,
        'description': f"Giảng viên: {giang_vien}",
        'start': {
            'dateTime': start_dt.isoformat(),
            'timeZone': 'Asia/Ho_Chi_Minh',
        },
        'end': {
            'dateTime': end_dt.isoformat(),
            'timeZone': 'Asia/Ho_Chi_Minh',
        },
        'recurrence': [
            f"RRULE:FREQ=WEEKLY;UNTIL={end_date.strftime('%Y%m%d')}T235959Z"
        ],
        'reminders': {
            'useDefault': False,
            'overrides': reminders if reminders else []
        }
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    event_id = event.get('id')
    print(f"📅 Đã tạo sự kiện: {event.get('summary')} ({event_id})")
    return event_id


# ----------------------------------------
# Xóa toàn bộ sự kiện TKB (có prefix [TKB])
# ----------------------------------------
def xoa_su_kien_tkb(service, prefix="[TKB]"):
    """
    Xóa toàn bộ sự kiện có prefix trong Google Calendar.

    Args:
        service: đối tượng Google Calendar service.
        prefix (str): tiền tố của sự kiện (mặc định "[TKB]").
    Returns:
        int: số sự kiện đã xóa.
    """
    events_result = service.events().list(
        calendarId='primary',
        singleEvents=True,
        orderBy='startTime',
        maxResults=2500
    ).execute()
    events = events_result.get('items', [])

    count = 0
    for event in events:
        if 'summary' in event and event['summary'].startswith(prefix):
            service.events().delete(calendarId='primary', eventId=event['id']).execute()
            count += 1

    print(f"🗑️ Đã xóa {count} sự kiện có prefix '{prefix}'.")
    return count
