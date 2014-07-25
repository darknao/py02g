# -*- coding: UTF-8 -*-
#
# Copyright (c) 2014 darknao
# https://github.com/darknao/pyO2g
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
import apiclient.errors

import logging
import httplib2
import httplib
import os.path
import datetime

import constants
import database


class cmd_flags(object):

    def __init__(self):
        self.short_url = True
        self.noauth_local_webserver = False
        self.logging_level = 'DEBUG'
        self.auth_host_name = 'localhost'
        self.auth_host_port = [8080, 9090]


class GoogleSecretsError(Exception):
    pass


class Google(object):
    """docstring for Google"""

    h = None  # httplib2 object
    flow = None  # oAuth2 connection flow
    storage = None  # credentials storage
    credentials = None  # credentials
    db = None  # sqlite3 connector
    log = None  # logger
    calId = None  # Destination Calendar

    def __init__(self, proxy=None):
        self.log = logging.getLogger(__name__)
        self.log.setLevel(logging.DEBUG)

        self.db = database.db()

        CLIENT_SECRETS = 'client_secrets.json'
        self.flow = client.flow_from_clientsecrets(
            CLIENT_SECRETS,
            scope=[
                'https://www.googleapis.com/auth/calendar',
                'https://www.googleapis.com/auth/calendar.readonly',
            ],
            # message=self.secretsMissings(CLIENT_SECRETS)
            )

        cacert = constants.resource_path('cacerts.txt')
        if os.path.isfile(cacert):
            self.log.debug("cacert : %s" % cacert)
        else:
            cacert = None
        self.h = httplib2.Http(proxy_info=proxy, ca_certs=cacert)
        self.storage = Storage('calendar.dat')
        self.credentials = self.storage.get()
        if self.credentials is None or self.credentials.invalid:
            self.credentials = tools.run_flow(self.flow, self.storage,
                                              cmd_flags(), http=self.h)

        # Validate the connection by listing calendars
        self.listCals(force=True)

    def secretsMissings(self, filename):
        raise GoogleSecretsError(
            """WARNING: Please configure Google API credentials!\r\n"""
            """\r\n"""
            """Please get a JSON key from the APIs Console"""
            """<https://code.google.com/apis/console>"""
            """and rename it to : %s""" % filename)

    def sendEvents(self, calId, evts):
        if self.calId is not None:
            c = self.db.cursor()
            http = self.credentials.authorize(self.h)
            service = build("calendar", "v3", http=http)

            for event in evts:

                # log.debug("with id: %s" % (event['myid'], ))
                # log.debug("\r\n%s" % (event,))
                myid = event.pop('myid')
                if 'updateID' in event:
                    # update
                    self.log.debug("update event [%s]" % (event['summary'], ))
                    updateID = event.pop('updateID')
                    try:
                        old_event = service.events().get(
                            calendarId=self.calId, eventId=updateID).execute()
                        event['sequence'] = old_event['sequence']
                        updated_event = service.events().update(
                            calendarId=self.calId,
                            eventId=updateID,
                            body=event).execute()
                    except apiclient.errors.HttpError, e:
                        self.log.debug("error: %s\r\nupdating event: %s"
                                       % (e.content, event,))
                        raise
                    if updated_event is not None:
                        now = datetime.datetime.now()
                        c.execute('''update sync set lastUpdated=?'''
                                  '''where gid=? and calId=?''',
                                  (now, updateID, calId,))
                        self.db.commit()

                else:
                    # new
                    self.log.debug("insert event [%s]" % (event['summary'], ))
                    try:
                        created_event = service.events().insert(
                            calendarId=self.calId, body=event).execute()
                    except apiclient.errors.HttpError, e:
                        self.log.debug("error: %s\r\ncreating event: %s"
                                       % (e.content, event,))
                        raise
                    if created_event is not None:
                        gid = created_event['id']
                        now = datetime.datetime.now()
                        c.execute(
                            '''insert into sync'''
                            '''(lastUpdated, calId, oid, gid)'''
                            '''values (?, ?, ?, ?)''',
                            (now, calId, myid, gid,))
                        self.db.commit()
        else:
            self.log.warning("no google calendar selected!")

    def deleteEvent(self, eventId):
        if self.calId is not None:
            c = self.db.cursor()

            http = self.credentials.authorize(self.h)
            service = build("calendar", "v3", http=http)
            try:
                service.events().delete(
                    calendarId=self.calId,
                    eventId=eventId).execute()
            except httplib.BadStatusLine, e:
                self.log.error("%s", e)
                # add an exception for HttpError 410
                # "Resource has been deleted"

            else:
                c.execute('''delete from sync where gid = ?''', (eventId,))
                self.db.commit()
        else:
            self.log.warning("no google calendar selected!")

    def cleanCal(self, calId, appts):
        if self.calId is not None:
            c = self.db.cursor()
            c.execute('''select oid,gid from sync where calId = ?''',
                      (calId, ))
            rows = c.fetchall()
            for row in rows:
                if not self.inCal(appts, row['oid'].upper()):
                    self.log.debug("deleting eventId [%s]..." % (row['gid'], ))
                    self.deleteEvent(row['gid'])
        else:
            self.log.warning("no google calendar selected!")

    def inCal(self, items, entryId):
        ret = False
        for i in range(1, items.Count + 1):
            if items.Item(i).EntryID == entryId:
                ret = True
                break
        return ret

    def listCals(self, force=False):
        c = self.db.cursor()

        ret = []
        calList = []

        # Refresh calendar list
        if force:
            http = self.credentials.authorize(self.h)
            service = build("calendar", "v3", http=http)

            lists = service.calendarList().list().execute(http=http)
            index = 0
            if 'items' in lists:
                c.execute('''delete from gcalendar''')
                gcals = lists['items']
                for gcal in gcals:
                    if (gcal['accessRole'] == "owner"
                            and gcal['kind'] == "calendar#calendarListEntry"):
                        calList.append((index, gcal['id'], gcal['summary']))
                        index = index + 1
                if index > 0:
                    c.executemany('''insert into gcalendar values(?, ?, ?)''',
                                  calList)
                    self.db.commit()

        # Get calendar list
        c.execute('''select * from gcalendar''')
        ret = c.fetchall()

        return ret
