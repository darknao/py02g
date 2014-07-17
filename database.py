import sqlite3

import constants

sql = sqlite3.connect(constants.DBFILE)
sql.row_factory = sqlite3.Row

schema = {
    'sync' : {
        'date' : 'text',
        'calId' : 'text',
        'oid' : 'text',
        'gid' : 'text',
        },
    'gcalendar' : {
        'id' : 'integer',
        'calId' : 'text',
        'description' : 'text',
        },
    }

checked = False

class DatabaseError(Exception): pass

def db():
    if checked == False:
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
    print "table %s created." % (table,)

def checkDB():
    # Check DB integrity
    global checked

    c = sql.cursor()
    error = 0
    print "checking db..."
    for table, cols in schema.iteritems():
        print "checking table %s..." % table
        c.execute("pragma table_info(%s)" % (table,))
        rows = c.fetchall()
        if len(rows) <= 0:
            print "ERROR: table %s not found! Creating..." % (table,)
            createTable(table)
        else:
            for colname, coltype in cols.iteritems():
                found = False
                for row in rows:
                    if row['name'] == colname and row['type'] == coltype:
                        found = True
                        break
                if found == False:
                    print "ERROR in %s: column %s not found or wrong type (%s)" % (table, colname, coltype)
                    error += 1
    if error == 0:
        checked = True
        print "database OK"
    else:
        raise DatabaseError("Database corrupted!")
    return checked