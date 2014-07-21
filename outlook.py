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

import pywintypes
import winerror
import datetime, pytz
import dateutil.rrule

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

if win32com.client.gencache.is_readonly == True:
    #allow gencache to create the cached wrapper objects
    win32com.client.gencache.is_readonly = False
    
    # under p2exe the call in gencache to __init__() does not happen
    # so we use Rebuild() to force the creation of the gen_py folder
    win32com.client.gencache.Rebuild()
    
    # NB You must ensure that the python...\win32com.client.gen_py dir does not exist
    # to allow creation of the cache in %temp%

msc = win32com.client.constants
rrule = dateutil.rrule

FREQNAMES = ['YEARLY','MONTHLY','WEEKLY','DAILY','HOURLY','MINUTELY','SECONDLY']

def rrule_to_string(rule):
    output = []
    
    parts = ['RRULE:FREQ='+FREQNAMES[rule._freq]]
    if rule._interval != 1:
        parts.append('INTERVAL='+str(rule._interval))
    if rule._wkst:
        parts.append('WKST='+str(rule._wkst))
    if rule._count:
        parts.append('COUNT='+str(rule._count))
    
    for name, value in [
            ('BYSETPOS', rule._bysetpos),
            ('BYMONTH', rule._bymonth),
            ('BYMONTHDAY', rule._bymonthday),
            ('BYYEARDAY', rule._byyearday),
            ('BYWEEKNO', rule._byweekno),
            ('BYDAY', rule._byweekday),
            ]:
        if value:
            if name == 'BYDAY':
                parts.append(name+'='+','.join(str(rrule.weekdays[v]) for v in value))
            else:
                parts.append(name+'='+','.join(str(v) for v in value))

    output.append(';'.join(parts))
    return '\n'.join(output)


class Outlook(object):
    """docstring for Outlook"""
    def __init__(self):
        self.log = logging.getLogger(__name__)
        self.log.setLevel(logging.DEBUG)

        self.createCOM()

        self.db = database.db()

    def createCOM(self):
        success = False
        try:
            self.outlookCOM = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            success = True
        except pywintypes.com_error, e:
            self.outlookCOM = None

        return success

    def isAlive(self):
        alive = False
        if self.outlookCOM != None:
            try:
                # Try something
                self.outlookCOM.GetNamespace("MAPI")
                alive = True
            except pywintypes.com_error, e:
                if e[0] in (winerror.REGDB_E_CLASSNOTREG, -2147023174):
                    # Try to recrate the connection
                    if self.createCOM():
                        alive = True
                else:
                    raise e
        else:
            if self.createCOM():
                alive = True
        return alive


    def getCalendars(self, extra = None):
        cals = []
        if self.isAlive():
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
        week = copy.copy(weekDays)
        weekN = copy.copy(weekDaysN)
        week.rotate(-now.weekday())
        weekN.rotate(-now.weekday())

        rPattrn = appt.GetRecurrencePattern()

        if rPattrn.RecurrenceType == msc.olRecursDaily:
            # Every N rPattrn.Interval
            startDate = datetime.datetime(appt.Start.year, appt.Start.month, appt.Start.day, appt.Start.hour, appt.Start.minute)
            if startDate >= now:
                nDate = startDate
            else:
                if rPattrn.NoEndDate == False:
                    # Either Occurrences or PatternEndDate
                    endDate = datetime.datetime(rPattrn.PatternEndDate.year, rPattrn.PatternEndDate.month, rPattrn.PatternEndDate.day,
                        rPattrn.PatternEndDate.hour, rPattrn.PatternEndDate.minute)
                    if endDate < now:
                        # event expired
                        return nDate

                nextday = (startDate.weekday() + rPattrn.Interval) % 7
                nDate = self.next_dow(eDate,(nextday,))

        elif rPattrn.RecurrenceType == msc.olRecursWeekly:
            #print "gNRD: week %s, weekN %s, pattern: %d" %(repr(week), repr(weekN), rPattrn.DayOfWeekMask)
            # find next occurrence in the next 7 days
            for i in range(0, 7):
                if rPattrn.DayOfWeekMask & week[i]:
                    nDate = self.next_dow(eDate,(weekN[i],))
                    #print "Next occurence : %s" % (nDate.isoformat(),)
                    break
        elif rPattrn.RecurrenceType == msc.olRecursMonthly:
            pass
        elif rPattrn.RecurrenceType == msc.olRecursMonthNth:
            pass
        elif rPattrn.RecurrenceType == msc.olRecursYearly:
            pass
        elif rPattrn.RecurrenceType == msc.olRecursYearNth:
            pass



        return nDate

    def next_dow(self, d, days):
        while d.weekday() not in days:
            d += datetime.timedelta(1)
        return d 

    def getRRULE(self, event):
        # dtstart=None,
        # bysetpos=None,
        #bymonth=None, bymonthday=None, byyearday=None, byeaster=None,
        #byweekno=None, byweekday=None,
        #byhour=None, byminute=None, bysecond=None
        RRULE = None 
        until = None
        count = None
        byweekday = None
        if event.IsRecurring:
            rp = event.GetRecurrencePattern()

            if rp.RecurrenceType == msc.olRecursDaily:
                freq = rrule.DAILY
            elif rp.RecurrenceType == msc.olRecursWeekly:
                freq = rrule.WEEKLY
            elif rp.RecurrenceType == msc.olRecursMonthly:
                freq = rrule.MONTHLY
            elif rp.RecurrenceType == msc.olRecursMonthNth:
                freq = rrule.MONTHLY
            elif rp.RecurrenceType == msc.olRecursYearly:
                freq = rrule.YEARLY
            elif rp.RecurrenceType == msc.olRecursYearNth:
                freq = rrule.YEARLY
            else:
                raise ValueError ("unknown recurrence type: %s" % rp.RecurrenceType )

            if rp.NoEndDate == False:
                until = datetime.datetime(rp.PatternEndDate.year, rp.PatternEndDate.month, rp.PatternEndDate.day,
                    rp.PatternEndDate.hour, rp.PatternEndDate.minute)
                count = rp.Occurrences
            
            interval = rp.Interval

            if rp.DayOfWeekMask > 0:
                byweekday = ()
                for i in range(0,7):
                    if rp.DayOfWeekMask & weekDays[i]:
                        byweekday += (rrule.weekdays[i],)

        RRULE = rrule.rrule(freq, count=count, interval=interval, until=until, byweekday=byweekday)

        return RRULE
        #rrule_to_string(RRULE)



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
            if dStart >= now or appt.IsRecurring and appt.ResponseStatus in (3,0,1,2) and appt.MeetingStatus in (0,1,3) :
                c.execute('''select lastUpdated, gid from sync
                         where oid=?''', (appt.EntryID.lower(),))
                r = c.fetchall()
                if len(r) <= 0:
                    # event not syncronized yet
                    events.append(self.createEvent(appt))
                else:
                    # event found: check last modification date
                    lastModification = datetime.datetime(appt.LastModificationTime.year, appt.LastModificationTime.month, appt.LastModificationTime.day,
                        appt.LastModificationTime.hour, appt.LastModificationTime.minute, appt.LastModificationTime.second )
                    self.log.debug("event [%s] last mod: %s" % (appt.Subject, lastModification))
                    if lastModification > r[0]['lastUpdated']:
                        # update event (or remove / recreate)
                        self.log.debug("event [%s] need update" % (appt.Subject,))
                        updatedEvent = self.createEvent(appt)
                        updatedEvent['updateID'] = r[0]['gid']
                        events.append(updatedEvent)
        return events

    def createEvent(self, appt): 
        tz=pytz.timezone('Europe/Paris') # TODO: Add a custom settings for this
        now = datetime.datetime.now()
        dStart = datetime.datetime(appt.Start.year, appt.Start.month, appt.Start.day, appt.Start.hour, appt.Start.minute)
        dEnd = datetime.datetime(appt.End.year, appt.End.month, appt.End.day, appt.End.hour, appt.End.minute)
        event = {
          'summary': appt.Subject,
          'location': appt.Location,
          'start': {
            'dateTime': dStart.isoformat(),
            'timeZone': 'Europe/Paris'
            # 'date' : for all day event
          },
          'end': {
            'dateTime': dEnd.isoformat(),
            'timeZone': 'Europe/Paris',
            # 'date' : for all day event
          },
          'description' : appt.Body,
          'myid' : appt.EntryID.lower(),
          'attendees': [
            {
              'email': 'francois.f.andrieu.external@airbus.com',
            },
            # ...
          ],
        }
        if appt.IsRecurring:
            recurrence = self.getRRULE(appt)
            if recurrence is not None:
                event['recurrence'] = [rrule_to_string(recurrence)]
            #rDate = self.getNextRecDate(appt)
            #if rDate != None:
            #    dStart = rDate.replace(hour=now.hour+1)
            #    dEnd = dEnd.replace(year=rDate.year, month=rDate.month, day=rDate.day,hour=now.hour+2)
            #    event['start']['dateTime'] = dStart.isoformat()
            #    event['end']['dateTime'] = dEnd.isoformat()
        if appt.ReminderSet:
            event['reminders'] = { "useDefault": False, "overrides" : [{"method" : "popup", "minutes" : appt.ReminderMinutesBeforeStart}]}

        return event


