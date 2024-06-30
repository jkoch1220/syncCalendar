import datetime
import logging
from O365 import Account, FileSystemTokenBackend
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Konfiguration für Microsoft Graph API
CLIENT_ID = 'DEIN_CLIENT_ID'
CLIENT_SECRET = 'DEIN_CLIENT_SECRET'
TENANT_ID = 'DEIN_TENANT_ID'
SCOPES = ['Calendars.ReadWrite']
TOKEN_PATH = 'o365_token.txt'

# Konfiguration für Google Calendar API
GOOGLE_CREDENTIALS_FILE = 'credentials.json'
GOOGLE_CALENDAR_ID = 'primary'

# Logging konfigurieren
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# O365-Konto authentifizieren
credentials = (CLIENT_ID, CLIENT_SECRET)
token_backend = FileSystemTokenBackend(token_path=TOKEN_PATH, tenant_id=TENANT_ID)
account = Account(credentials, token_backend=token_backend)

if not account.is_authenticated:
    logger.info('Authentifiziere das O365-Konto...')
    account.authenticate(scopes=SCOPES)
    logger.info('O365-Konto erfolgreich authentifiziert.')

schedule = account.schedule()
calendar = schedule.get_default_calendar()

# Google Calendar API authentifizieren
logger.info('Authentifiziere die Google Calendar API...')
google_credentials = service_account.Credentials.from_service_account_file(
    GOOGLE_CREDENTIALS_FILE,
    scopes=['https://www.googleapis.com/auth/calendar']
)

service = build('calendar', 'v3', credentials=google_credentials)
logger.info('Google Calendar API erfolgreich authentifiziert.')

def sync_outlook_to_google():
    # Hole Termine aus dem Outlook-Kalender
    logger.info('Hole Termine aus dem Outlook-Kalender...')
    events = calendar.get_events(start=datetime.datetime.now(), end=datetime.datetime.now() + datetime.timedelta(days=30))
    logger.info(f'{len(events)} Termine gefunden.')

    for event in events:
        google_event = {
            'summary': event.subject,
            'start': {
                'dateTime': event.start.strftime("%Y-%m-%dT%H:%M:%S"),
                'timeZone': 'UTC',
            },
            'end': {
                'dateTime': event.end.strftime("%Y-%m-%dT%H:%M:%S"),
                'timeZone': 'UTC',
            },
            'description': event.body,
        }
        # Erstelle den Termin im Google Kalender
        logger.info(f'Erstelle Termin im Google Kalender: {event.subject}')
        service.events().insert(calendarId=GOOGLE_CALENDAR_ID, body=google_event).execute()
        logger.info(f'Termin erstellt: {event.subject}')

# Synchronisation ausführen
logger.info('Starte Synchronisation von Outlook zu Google Kalender...')
sync_outlook_to_google()
logger.info('Synchronisation abgeschlossen.')
