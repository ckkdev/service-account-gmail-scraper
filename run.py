from scraper import AccountServiceScraper

SCOPES = ['https://mail.google.com/']
SERVICE_ACCOUNT_FILE = 'keyfile.json'
EMAIL_USERS = ['genero@ckk.dev', 'dustin@ckk.dev']

AccountServiceScraper(EMAIL_USERS, SCOPES, SERVICE_ACCOUNT_FILE)