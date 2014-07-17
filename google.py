# -*- coding: UTF-8 -*-
#
# Copyright (c) 2014 darknao
# https://github.com/darknao/py02g
#
# This file is part of pyO2g.
# 
# pyO2g is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

from oauth2client.client import OAuth2WebServerFlow
from oauth2client.file import Storage
from oauth2client import tools
from oauth2client import client
from apiclient.discovery import build

import logging
import httplib2
import httplib
import os.path

import constants
import database

class cmd_flags(object):
  def __init__(self):
    self.short_url = True 
    self.noauth_local_webserver = False
    self.logging_level = 'DEBUG' 
    self.auth_host_name = 'localhost'
    self.auth_host_port = [8080, 9090]

class GoogleSecretsError(Exception): pass


class Google(object):
    """docstring for Google"""

    h = None # httplib2 object
    flow = None # oAuth2 connection flow
    storage = None # credentials storage
    credentials = None # credentials
    db = None # sqlite3 connector
    log = None # logger
    calId = None # Destination Calendar

    def __init__(self, proxy=None):
        logging.basicConfig()
        self.log = logging.getLogger('Google')
        self.log.setLevel(logging.DEBUG)

        self.db = database.db()

        CLIENT_SECRETS = 'client_secrets.json'
        self.flow = client.flow_from_clientsecrets(CLIENT_SECRETS,
          scope=[
              'https://www.googleapis.com/auth/calendar',
              'https://www.googleapis.com/auth/calendar.readonly',
            ],
            #message=self.secretsMissings(CLIENT_SECRETS)
            )

        """my_proxy_info = httplib2.ProxyInfo(proxy_type =  httplib2.socks.PROXY_TYPE_SOCKS5,
            proxy_host = '127.0.0.1',
            proxy_port = 13337,
            #proxy_user = 'myLogin',
            #proxy_pass = 'myPassword',
            proxy_rdns = True,
            #proxy_user_agent = 'myCompanyUserAgent'
            )"""
        cacert = constants.resource_path('cacerts.txt')
        if os.path.isfile(cacert):
            print "cacert : %s" % cacert
        else:
            cacert = None
        self.h = httplib2.Http(proxy_info=proxy, ca_certs=cacert)
        self.storage = Storage('calendar.dat')
        self.credentials = self.storage.get()
        if self.credentials is None or self.credentials.invalid == True:
            self.credentials = tools.run_flow(self.flow, self.storage, cmd_flags(), http=self.h)

    def secretsMissings(self, filename):
        raise GoogleSecretsError("""WARNING: Please configure Google API credentials!

Please get a JSON key from the APIs Console <https://code.google.com/apis/console>
and rename it to : %s""" % filename)


    def sendEvents(self, calId, evts):
        if self.calId != None:
            c = self.db.cursor()
            # gkey : notasecret        
            http = self.credentials.authorize(self.h)
            service = build("calendar", "v3", http=http)

            for event in evts:
                self.log.debug("insert event [%s]" % (event['summary'], ))
                #log.debug("with id: %s" % (event['myid'], ))
                #log.debug("\r\n%s" % (event,))
                created_event = service.events().insert(calendarId=self.calId, body=event).execute()
                if created_event != None:
                    gid = created_event['id']
                    c.execute('''insert into sync (calId, oid, gid) values (?, ?, ?)''', (calId, event['myid'] ,gid,))
                    self.db.commit()
        else:
            self.log.warning("no google calendar selected!")

    def deleteEvent(self, eventId):
        if self.calId != None:
            c = self.db.cursor()

            http = self.credentials.authorize(self.h)
            service = build("calendar", "v3", http=http)
            try:
                service.events().delete(calendarId=self.calId, eventId=eventId).execute()
            except httplib.BadStatusLine, e:
                self.log.error("%s",e)
            else:
                c.execute('''delete from sync where gid = ?''', (eventId,))
                self.db.commit()
        else:
            self.log.warning("no google calendar selected!")

    def cleanCal(self, calId, appts):
        if self.calId != None:
            c = self.db.cursor()
            c.execute('''select oid,gid from sync where calId = ?''', (calId, ))
            rows = c.fetchall()
            for row in rows:
                if self.inCal(appts, row['oid'].upper()) == False:
                    self.log.debug("deleting eventId [%s]..." % (row['gid'], ))
                    self.deleteEvent(row['gid'])
        else:
            self.log.warning("no google calendar selected!")

    def inCal(self, items, entryId):
        ret = False
        for i in range(1,items.Count +1):
            if items.Item(i).EntryID == entryId:
                ret = True
                break
        return ret



    def listCals(self, force=False):
        c = self.db.cursor()

        ret = []
        calList = []

        # Refresh calendar list
        if force == True:
            http = self.credentials.authorize(self.h)
            service = build("calendar", "v3", http=http)

            lists = service.calendarList().list().execute(http=http)
            index = 0
            if 'items' in lists:
                c.execute('''delete from gcalendar''')
                gcals = lists['items']
                for gcal in gcals:
                    if gcal['accessRole'] == "owner" and gcal['kind'] == "calendar#calendarListEntry":  
                        calList.append((index, gcal['id'], gcal['summary']))
                        index = index + 1
                if index > 0:
                    c.executemany('''insert into gcalendar values(?, ?, ?)''', calList)
                    self.db.commit()

        # Get calendar list
        c.execute('''select * from gcalendar''')
        ret = c.fetchall()

        return ret

# [u'kind', u'foregroundColor', u'defaultReminders', u'primary', u'colorId', u'notificationSettings', u'summary', u'etag', u'backgroundColor', u'timeZone', u'accessRole', u'id']