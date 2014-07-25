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

import sqlite3
import logging

import constants

sql = sqlite3.connect(
    constants.DBFILE,
    detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
sql.row_factory = sqlite3.Row

schema = {
    'sync': {
        'lastUpdated':  'timestamp',
        'calId':        'text',
        'oid':          'text',
        'gid':          'text',
        },
    'gcalendar': {
        'id':           'integer',
        'calId':        'text',
        'description':  'text',
        },
    }

checked = False
logging.basicConfig(
    filename='logfile.txt',
    format='%(asctime)-23s - %(levelname)-7s - %(name)s - %(message)s')
log = logging.getLogger(__name__)
log.setLevel(logging.DEBUG)


class DatabaseError(Exception):
    pass


def db():
    if not checked:
        checkDB()
    return sql


def createTable(table):
    colsList = ""
    for colname, coltype in schema[table].iteritems():
        colsList += "%s %s," % (colname, coltype)
    colsList = colsList.rstrip(',')
    stmt = "CREATE TABLE %s (%s)" % (table, colsList)

    c = sql.cursor()
    c.execute(stmt)
    sql.commit()
    log.debug("table %s created." % (table,))


def checkDB():
    # Check DB integrity
    global checked

    c = sql.cursor()
    error = 0
    log.info("checking db...")
    for table, cols in schema.iteritems():
        log.debug("checking table %s..." % table)
        c.execute("pragma table_info(%s)" % (table,))
        rows = c.fetchall()
        if len(rows) <= 0:
            log.error("table %s not found! Creating..." % (table,))
            createTable(table)
        else:
            for colname, coltype in cols.iteritems():
                found = False
                for row in rows:
                    if row['name'] == colname:
                        if row['type'] != coltype:
                            log.error("%s: column %s has wrong type"
                                      "(%s instead of %s)"
                                      % (table, colname, row['type'], coltype))
                        else:
                            found = True
                            break
                if not found:
                    log.error("%s: column %s not found or wrong type (%s)"
                              % (table, colname, coltype))
                    error += 1
    if error == 0:
        checked = True
        log.info("database OK")
    else:
        raise DatabaseError("Database corrupted!")
    return checked
