# Daniel's Timesheet


A simple Python command line tool to interact with my timesheet (Google Spreadsheet) on Google Drive using [gsheets](https://github.com/xflr6/gsheets).


# Installation
Simply run `./run.sh` to use the tool.


# Usage

Make sure to read the gsheets README in order to create the proper client-secrets.json file for your Google auth account.

Also, be aware, that on the first run, the tool will open your web-browser to authenticate you. After that the auth data is cached in the client-secrets-cache.json file.


# TODO

- add mail sending feature
- improve error handling
- allow editing for entries
- allow better filtering
- allow showing past entries
- allow showing all entries
- minimize reads by caching last accessed data/row, and use that next time the tool is involved??? (problem: cache invalidation?!)


# LICENSE and COPYRIGHT

2017-08-15

Copyright 2017 by Daniel Kurashige-Gollub (daniel@kurashige-gollub.de)


[MIT-License](LICENSE.md)
