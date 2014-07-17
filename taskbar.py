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

import wx
import images
import sys
import constants


class TaskBarIcon(wx.TaskBarIcon):
    TBMENU_RESTORE = wx.NewId()
    TBMENU_CLOSE   = wx.NewId()
    TBMENU_QUIT  = wx.NewId()
    TBMENU_SYNC = wx.NewId()

    def __init__(self, frame):
        wx.TaskBarIcon.__init__(self)
        self.frame = frame

        icon = self.MakeIcon(images.getIconImage())
        self.SetIcon(icon, constants.APPNAME)
        self.imgidx = 1

        self.Bind(wx.EVT_TASKBAR_LEFT_DCLICK, self.OnTaskBarToggle)
        self.Bind(wx.EVT_MENU, self.OnTaskBarActivate, id=self.TBMENU_RESTORE)
        self.Bind(wx.EVT_MENU, self.OnTaskBarClose, id=self.TBMENU_CLOSE)
        self.Bind(wx.EVT_MENU, self.OnTaskBarQuit, id=self.TBMENU_QUIT)
        self.Bind(wx.EVT_MENU, frame.OnTimer, id=self.TBMENU_SYNC)

    def __del__(self):
        self.RemoveIcon()

    def CreatePopupMenu(self):
        menu = wx.Menu()
        menu.Append(self.TBMENU_SYNC,   "Sync Now")
        menu.AppendSeparator()
        menu.Append(self.TBMENU_RESTORE, "Restore window")
        menu.Append(self.TBMENU_CLOSE,   "Hide window")
        menu.AppendSeparator()
        menu.Append(self.TBMENU_QUIT, "Quit %s" % (constants.APPNAME,))

        return menu


    def MakeIcon(self, img):
        if "wxMSW" in wx.PlatformInfo:
            img = img.Scale(16, 16)
        elif "wxGTK" in wx.PlatformInfo:
            img = img.Scale(22, 22)
        icon = wx.IconFromBitmap( img.ConvertToBitmap() )
        return icon


    def OnTaskBarActivate(self, evt):
        if self.frame.IsIconized():
            self.frame.Iconize(False)
        if not self.frame.IsShown():
            self.frame.Show(True)
        self.frame.Raise()


    def OnTaskBarToggle(self, evt):
        if self.frame.IsIconized():
            self.frame.Iconize(False)
        if not self.frame.IsShown():
            self.frame.Show(True)
            self.frame.Raise()
        else:
            self.frame.Show(False)


    def OnTaskBarClose(self, evt):
        self.frame.Show(False)

    def OnTaskBarQuit(self, evt):
        self.RemoveIcon()
        if evt is not None:
            self.frame.Close()
        sys.exit(0)


