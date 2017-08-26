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

MAX_FILENAME_LENGTH = 245

import os
#os.chdir("/tmp")   # change working directory
