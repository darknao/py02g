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

import win32com.client
import datetime, pytz
import logging
from collections import deque
import copy

import database

olFriday = 32 # Friday
olMonday = 2 # Monday
olSaturday = 64 # Saturday
olSunday = 1 # Sunday
olThursday = 16 # Thursday
olTuesday = 4 # Tuesday
olWednesday = 8 # Wednesday

weekDays = deque((olMonday, olTuesday, olWednesday, olThursday, olFriday, olSaturday, olSunday))
weekDaysN = deque((0,1,2,3,4,5,6))
dow={d:i for i,d in 
         enumerate('Mon,Tue,Wed,Thu,Fri,Sat,Sun'.split(','))}

class Outlook(object):
    """docstring for Outlook"""
    def __init__(self):
        logging.basicConfig()
        self.log = logging.getLogger('pyO2gcal')
        self.log.setLevel(logging.DEBUG)

        self.outlookCOM = win32com.client.Dispatch("Outlook.Application")

        self.db = database.db()

    def getCalendars(self, extra = None):
        cals = []
        ns = self.outlookCOM.GetNamespace("MAPI")
        cals.append(ns.GetDefaultFolder(9))
        if extra != None and extra != "" :
            r = ns.CreateRecipient(extra)
            cals.append(ns.GetSharedDefaultFolder(r,9))
        return cals

    def getNextRecDate(self, appt):
        nDate = datetime.datetime(1980,1,1)
        now = datetime.datetime.now()
        eDate = datetime.datetime(now.year, now.month, now.day, appt.Start.hour, appt.Start.minute)
        rPattrn = appt.GetRecurrencePattern()
        week = copy.copy(weekDays)
        weekN = copy.copy(weekDaysN)
        week.rotate(-now.weekday())
        weekN.rotate(-now.weekday())
        #print "gNRD: week %s, weekN %s, pattern: %d" %(repr(week), repr(weekN), rPattrn.DayOfWeekMask)
        # find next occurrence in the next 7 days
        for i in range(0, 7):
            if rPattrn.DayOfWeekMask & week[i]:
                nDate = self.next_dow(eDate,(weekN[i],))
                print "Next occurence : %s" % (nDate.isoformat(),)
                break
        return nDate

    def next_dow(self, d, days):
        while d.weekday() not in days:
            d += datetime.timedelta(1)
        return d 


    def getAppt(self, appts):
        events = []
        c = self.db.cursor()
        for i in range(1,appts.Count+1):
            appt=appts.Item(i)

            now = datetime.datetime.now()
            dStart = datetime.datetime(appt.Start.year, appt.Start.month, appt.Start.day, appt.Start.hour, appt.Start.minute)
            if appt.IsRecurring:
                rDate = self.getNextRecDate(appt)
                if rDate != None:
                    dStart = rDate
            if dStart >= now or (appt.IsRecurring and self.getNextRecDate(appt) >= now) and appt.ResponseStatus in (3,0,1,2) and appt.MeetingStatus in (0,1,3) :
                c.execute('''select gid from sync
                         where oid=?''', (appt.EntryID.lower(),))
                r = c.fetchall()
                if len(r) <= 0:
                    events.append(self.createEvent(appt))
        return events

    def createEvent(self, appt): 
        tz=pytz.timezone('Europe/Paris')
        now = datetime.datetime.now()
        dStart = datetime.datetime(appt.Start.year, appt.Start.month, appt.Start.day, appt.Start.hour, appt.Start.minute)
        dEnd = datetime.datetime(appt.End.year, appt.End.month, appt.End.day, appt.End.hour, appt.End.minute)
        if appt.IsRecurring:
            rDate = self.getNextRecDate(appt)
            if rDate != None:
                dStart = rDate
                dEnd = dEnd.replace(year=rDate.year, month=rDate.month, day=rDate.day)
        event = {
          'summary': appt.Subject,
          'location': appt.Location,
          'start': {
            'dateTime': tz.localize(dStart).isoformat()
            # 'date' : for all day event
          },
          'end': {
            'dateTime': tz.localize(dEnd).isoformat()
            # 'date' : for all day event
          },
          'description' : appt.Body,
          'myid' : appt.EntryID.lower()
          #'attendees': [
          #  {
          #    'email': 'attendeeEmail',
          #  },
          #  # ...
          #],
        }
        if appt.ReminderSet:
            event['reminders'] = { "useDefault": False, "overrides" : [{"method" : "popup", "minutes" : appt.ReminderMinutesBeforeStart}]}

        return event


