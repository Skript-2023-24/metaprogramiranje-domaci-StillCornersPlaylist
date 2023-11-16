require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md
session = GoogleDrive::Session.from_config("config.json")

ws = session.spreadsheet_by_key("1752Q5uKm5vbCU49faGWieQX1zIkrJ9bJZkoMp9yCIJc").worksheets[0]

# Gets content of A2 cell.
p ws[2, 1]  #==> "Test"