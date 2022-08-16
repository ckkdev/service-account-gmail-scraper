from scraper import AccountServiceScraper

SCOPES = ['https://mail.google.com/']
SERVICE_ACCOUNT_FILE = 'keyfile.json'
EMAIL_USERS = ['email1@domain.com', 'email2@domain.com', 'email3@domain.com']

AccountServiceScraper(EMAIL_USERS, SCOPES, SERVICE_ACCOUNT_FILE)
