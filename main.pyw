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

import wx
import logging
import time
import sys
import ConfigParser
import win32com.client
import pywintypes
import winerror
import re
import datetime
import httplib2
import httplib
import copy
from oauth2client import clientsecrets
import apiclient.errors

import constants
import images
import outlook
import google
from taskbar import TaskBarIcon


def oCalNiceName(cal):
    # print "cal:%s" % cal.FolderPath
    regxp = re.match(r"\\\\(.*)\\", cal.FolderPath)
    return regxp.groups()[0]


class MainFrame(wx.Frame):
    outlook = None
    retry = 0

    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, wx.ID_ANY, title,
                          pos=(600, 300),
                          style = (wx.CAPTION | wx.MINIMIZE_BOX |
                                   wx.RESIZE_BORDER | wx.SYSTEM_MENU |
                                   wx.CLOSE_BOX),
                          size=(350, 280))
        busy = wx.BusyInfo("Initialization, please wait...")
        self.cfg = ConfigParser.ConfigParser()
        self.cfg.read(constants.CFGFILE)

        self.SetIcon(images.getIconIcon())
        self.tbicon = TaskBarIcon(self)
        menuBar = wx.MenuBar()
        menu = wx.Menu()
        helpmenu = wx.Menu()

        syncNow = menu.Append(wx.ID_ANY, "&Sync Now\tCtrl+S",
                              "Force synchronization")
        menu.AppendSeparator()

        quit = wx.MenuItem(menu, wx.ID_EXIT, 'E&xit\tCtrl+Q',
                           'Quit the Application')
        # quit.SetBitmap(wx.Image('stock_exit-16.png',
        #                 wx.BITMAP_TYPE_PNG).ConvertToBitmap())
        menu.AppendItem(quit)

        about = helpmenu.Append(wx.ID_ANY, "&About")

        self.Bind(wx.EVT_MENU, self.OnTimer, syncNow)
        self.Bind(wx.EVT_MENU, self.OnTimeToClose, quit)
        self.Bind(wx.EVT_MENU, self.OnAbout, about)
        self.Bind(wx.EVT_CLOSE, self.OnTimeToClose)
        self.Bind(wx.EVT_ICONIZE, self.OnMinimize)

        menuBar.Append(menu, "&File")

        menuBar.Append(helpmenu, "&Help")
        self.SetMenuBar(menuBar)

        self.CreateStatusBar()
        self.SetStatusText("Idle")

        panel = wx.Panel(self)

        self.destCalText = wx.StaticText(panel, wx.ID_ANY, "Google calendar:")
        self.srcCalText = wx.StaticText(panel, wx.ID_ANY, "Outlook calendars:")

        self.debugText = wx.StaticText(panel, wx.ID_ANY, "")

        self.log = wx.TextCtrl(
            panel, wx.ID_ANY, "",
            size=(200, 100),
            style=wx.TE_MULTILINE | wx.TE_READONLY)

        sizer = wx.BoxSizer(wx.VERTICAL)

        self.getOutlookCals()

        proxy = None
        if self.cfg.getboolean('Proxy', 'enabled'):
            proxy = httplib2.ProxyInfo(
                proxy_type=httplib2.socks.PROXY_TYPE_HTTP_NO_TUNNEL,
                proxy_host=self.cfg.get('Proxy', 'host'),
                proxy_port=self.cfg.getint('Proxy', 'port'),
                proxy_user=self.cfg.get('Proxy', 'username'),
                proxy_pass=self.cfg.get('Proxy', 'password'),
                proxy_rdns=True)

        destCalBox = wx.BoxSizer(wx.HORIZONTAL)
        destCalBox.Add(self.destCalText, 0,
                       wx.ALIGN_CENTER_VERTICAL | wx.ALL, 1)

        self.listDestCal = wx.ComboBox(panel, wx.ID_ANY,
                                       size=(150, -1), style=wx.CB_READONLY)
        destCalBox.Add(self.listDestCal, 0,
                       wx.GROW | wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.Bind(wx.EVT_COMBOBOX, self.OnGcalSelect, self.listDestCal)

        reloadDestCalBtn = wx.Button(panel, wx.ID_ANY, 'reload')
        destCalBox.Add(reloadDestCalBtn, 0, wx.ALIGN_CENTER_VERTICAL, 5)
        sizer.Add(destCalBox, 0, wx.EXPAND | wx.ALL, 1)

        self.Bind(
            wx.EVT_BUTTON,
            lambda event: self.reloadDestCal(event, force=True),
            reloadDestCalBtn)

        try:
            self.google = google.Google(proxy=proxy)
            self.reloadDestCal()
        except httplib2.ServerNotFoundError, e:
            del busy
            dlg = wx.MessageDialog(
                self,
                "%s\r\nCheck your Internet connection "
                "and/or your proxy settings." % e,
                "Connection error!",
                wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            sys.exit(-1)
        except clientsecrets.InvalidClientSecretsError, e:
            del busy
            dlg = wx.MessageDialog(
                self,
                """WARNING: Google API credentials not found!\r\n"""
                """\r\n"""
                """Please get a JSON key from the APIs Console"""
                """<https://code.google.com/apis/console>"""
                """and rename it to client_secrets.json""",
                "%s" % e,
                wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()
            sys.exit(-1)
        except httplib2.socks.HTTPError, e:
            if e.message[0] == 407:  # Proxy Authentication Required
                dlg = wx.MessageDialog(
                    self,
                    """Proxy authentication error!\r\n"""
                    """Please check your proxy credentials""",
                    "%s" % e.message[1],
                    wx.OK | wx.ICON_ERROR)
                del busy
                dlg.ShowModal()
                dlg.Destroy()
                sys.exit(-1)
            else:
                raise

        sizer.Add(self.srcCalText, 0, wx.EXPAND | wx.ALL, 1)
        for oCal in self.oCals:
            chkb = wx.CheckBox(panel, wx.ID_ANY,
                               oCalNiceName(oCal), pos=(10, 10))
            chkb.Enable(False)
            chkb.SetValue(1)
            sizer.Add(chkb, 0, wx.EXPAND | wx.ALL, 1)
            # sizer.Add(copy.copy(chkb),0,wx.EXPAND | wx.ALL ,1)

        try:
            interval = self.cfg.getint('Main', 'syncInterval')
        except (ConfigParser.NoSectionError, ConfigParser.NoOptionError), e:
            interval = 30  # default: 30 minutes

        if len(self.oCals) <= 0:
            sizer.Add(
                wx.StaticText(panel, wx.ID_ANY, "no calendar!"),
                0, wx.EXPAND | wx.ALL, 1)

        self.timer = wx.Timer(self)
        self.retryTimer = wx.Timer(self)
        self.timer.Start(interval * 60 * 1000)  # convert to ms
        self.Bind(wx.EVT_TIMER, self.OnTimer)

        line = wx.StaticLine(panel, wx.ID_ANY, size=(20, -1),
                             style=wx.LI_HORIZONTAL)
        sizer.Add(line, 0,
                  wx.GROW | wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.TOP, 5)

        sizer.Add(self.debugText, 0, wx.ALL, 1)

        sizer.Add(self.log, 4, wx.ALL | wx.EXPAND, 1)

        panel.SetSizer(sizer)
        panel.Layout()

        del busy
        busy = wx.BusyInfo("Syncing...")
        self.OnTimer()
        del busy
        self.tbicon.ShowBalloon(
            "%s" % (title,),
            "%s is running..." % (constants.APPNAME,),
            100)

    def getOutlookCals(self):
        if self.outlook is None:
            self.outlook = outlook.Outlook()

        self.oCals = self.outlook.getCalendars(
            self.cfg.get('Outlook', 'extraCal'))

    def OnTimeToClose(self, evt):
        self.tbicon.OnTaskBarQuit(None)
        self.Destroy()

    def OnAbout(self, evt):
        # NOTE: Crash with wxPython3.0.0 & py2exe
        # (use wxPython2.9.5 until it's fixed)
        about = wx.AboutDialogInfo()
        about.Name = constants.APPNAME
        about.Version = "v%s" % constants.VERSION
        about.Copyright = "(c) 2014 darknao / rBus Radio Team"
        about.WebSite = "https://github.com/darknao/pyO2g"
        wx.AboutBox(about)

    def OnTimer(self, evt=None):
        self.PushStatusText("Syncing...")
        # SYnc stuff here...
        try:
            self.syncMyCal()
        except httplib.BadStatusLine:
            # google doing shit? let's try again
            if self.retry < constants.MAX_RETRY:
                self.log.AppendText("Google not responding."
                                    "Trying again in 5secs...\r\n")
                self.retryTimer.Start(
                    (constants.BASE_RETRY_TIME * self.retry
                        + constants.BASE_RETRY_TIME),
                    wx.TIMER_ONE_SHOT)
                self.retry += 1
            else:
                self.log.AppendText("Google not responding. Giving up.\r\n")
        except httplib2.socks.HTTPError, e:
            if e.message[0] == 407:  # Proxy Authentication Required
                self.log.AppendText("Proxy Authentication error!"
                                    "Check your credentials.\r\n")
            else:
                self.log.AppendText("HTTP Error: %s", e.message[1])
        else:
            self.retry = 0
        self.PopStatusText()

    def syncMyCal(self):
        # TODO: use wx.lib.delayedresult for that
        now = datetime.datetime.now()
        if self.google.calId is not None:
            if self.outlook.isAlive():
                self.getOutlookCals()
                self.log.AppendText("%s: Starting sync...\r\n" % now)
                for oCal in self.oCals:
                    calId = oCal.EntryID
                    self.log.AppendText("o Syncing cal: %s\r\n"
                                        % (oCalNiceName(oCal),))
                    calItems = oCal.Items

                    calItems.Sort("[Start]", True)
                    calItems.IncludeRecurrences = "True"
                    evts = self.outlook.getAppt(calItems)
                    self.log.AppendText(" -> Cleaning events...\r\n")

                    try:
                        self.google.cleanCal(calId, oCal.Items)
                        self.log.AppendText(" -> Syncing %s new events\r\n"
                                            % (len(evts),))
                    except apiclient.errors.HttpError, e:
                        self.log.AppendText("gAPI error: %s\r\n" % (e,))

                    if len(evts) > 0:
                        try:
                            self.google.sendEvents(calId, evts)
                        except apiclient.errors.HttpError, e:
                            self.log.AppendText("gAPI error: %s\r\n" % (e,))
                now = datetime.datetime.now()
                self.log.AppendText("%s: Sync completed\r\n" % now)
                self.log.AppendText(
                    "next sync in %d minutes\r\n" %
                    round(self.timer.GetInterval() / 60 / 1000, 2))
            else:
                self.log.AppendText("Outlook not running!\r\n")
        else:
            self.log.AppendText("Please select a Google calendar!\r\n")

    def OnGcalSelect(self, evt):
        cBox = evt.GetEventObject()
        if cBox.GetId() == self.listDestCal.GetId():
            calIndex = cBox.GetCurrentSelection()
            self.debugText.SetLabel(
                "gCalendar selected: %s: %s [%s]" %
                (calIndex, cBox.GetValue(), self.gCals[calIndex]['calId']))
            self.cfg.set('Google', 'calId', self.gCals[calIndex]['calId'])
            self.google.calId = self.gCals[calIndex]['calId']
            with open(constants.CFGFILE, 'w') as configfile:
                self.cfg.write(configfile)

    def reloadDestCal(self, evt=None, force=False):
        if evt is not None:
            btn = evt.GetEventObject()
            btn.Disable()
        self.PushStatusText("Fetching Google Calendars list...")
        selectedId = 0
        gcalId = self.cfg.get('Google', 'calId')
        self.gCals = self.google.listCals(force=force)
        if len(self.gCals) > 0:
            self.listDestCal.Clear()
            for gCal in self.gCals:
                self.listDestCal.Append(gCal['description'])
                if gCal['calId'] == gcalId:
                    selectedId = gCal['id']
                    self.google.calId = gcalId
            self.listDestCal.SetSelection(selectedId)
            if evt is not None:
                btn.Enable()

        self.PopStatusText()

    def OnMinimize(self, evt):
        self.Show(False)


app = wx.App(redirect=True, filename="logfile.txt")

frame = MainFrame(None, "%s v%s" % (constants.APPNAME, constants.VERSION))
app.SetTopWindow(frame)
# frame.Show()

app.MainLoop()
