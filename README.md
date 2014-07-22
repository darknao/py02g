# py02g #

py02g is a (very) simple application to synchronize your Outlook calendar with your Google account.
For now, it only support one way synchronization (from Outlook to Google).

This is pretty much the same thing as Google Calendar Sync, but with a lot more bugs, and less features :stuck_out_tongue_closed_eyes:

## Features ##
- One way synchronization
- Ability to select your destination calendar
- Crash randomly (I'm working on it)

## Getting started ##

### Dependencies ###
- Python 2.7 (may work on 2.6 idk)
- [Google APIs Client Library](https://developers.google.com/api-client-library/python/start/installation)
- [Python Win32 Extensions (pywin32)](http://sourceforge.net/projects/pywin32/)
- wxPython 2.9.5
- dateutil
- sqlite3

or, you can use the [bundled version](https://github.com/darknao/pyO2g/releases/latest) (with py2exe) 

### Connection to your Google account ###
For that, you'll need an OAuth 2.0 client ID from Google (see [here](https://developers.google.com/console/help/new/#generatingoauth2)).  
Create a Client ID from the [google developers console](https://console.developers.google.com) with the followings parameters :  
**Application Type:** Installed application  
**Installed Application Type:** Other  
Download the JSON key and copy it in the same place than pyO2g (rename it to **client_secrets.json**).

### Configuration ###
Most of the settings can be found in pyO2g.cfg.
For now, there is only 2 settings available.

#### Synchronization interval ###
By default, it's 20 minutes.
You can change that value in the Main section of the configuration file :
```INI
[Main]
syncinterval = 20
```

#### Proxy settings ####
If your Internet connection require a proxy (ie. company proxy), you can set it up here :
```INI
[Proxy]
enabled = yes
host = fqdn_or_ip_of_your_proxy
port = 3128
username = user
password = pass
```

