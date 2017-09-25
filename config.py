accounts = [

            {"name": "T-Online",
             "server": 'secureimap.t-online.de',
             "port": 993,
             "user": "username@t-online.de",
             "pwd": "your password (if not specified here, MailArchiver will prompt for user password)",
             "ssl": True,
             "ignorefolders": ["INBOX Trash"]
            }

           ]


MAX_FILENAME_LENGTH = 255
MAX_FILENAME_EXTENSION_LENGTH = 6
FILENAME_RE = '[^\w \.\-@\[\]]'
DEFAULT_EXTENSION = ".txt"

# set working directory here:
WORKING_DIR = "/media/usb/"
