# -*- coding: utf-8 -*-
import sqlite3

class Connection:
    fileName = "ConditionPark.db"

    def __init__(self):
        self.basa = sqlite3.connect(self.fileName)
        self.cursor = self.basa.cursor()

    def SelectTable(self, table, collums):
        coll = str(collums[0])
        for i in range(1, len(collums)):
            coll += ", " + str(collums[i])
        self.sql = """SELECT %s FROM %s;""" % (coll, table)
        print(self.sql)
        self.cursor.execute(self.sql)

        self.result = self.cursor.fetchall()
        return self.result

    def Update(self, table, collumn, value, id, index):
        #UPDATE client SET FIO='Абрахам абрамвам вумаван' WHERE id_client=60;
        self.sql = """UPDATE %s SET %s='%s' WHERE %s=%s;""" % (str(table), str(collumn), str(value), str(id), str(index))
        print(self.sql)
        self.cursor.execute(self.sql)
        self.basa.commit()


    def Delete(self, table, id, index):
        #DELETE FROM client WHERE id_client='60';
        self.sql = """DELETE FROM %s WHERE %s='%s' """ % (str(table), str(id), str(index))
        print(self.sql)
        self.cursor.execute(self.sql)
        self.basa.commit()

    def Create(self, table, collumn, values):
        coll = str(collumn[0])
        val = str(values[0])
        for i in range(1, len(collumn)):
            coll += ", " + str(collumn[i])
            val += ", '" + str(values[i]) + "'"

        self.sql = """INSERT INTO %s(%s) VALUES (%s);""" % (table, coll, val)

        print(self.sql)

        self.cursor.execute(self.sql)
        self.basa.commit()

    def Selection(self, selection):
        self.sql = selection
        print(self.sql)

        self.cursor.execute(self.sql)
        self.result = self.cursor.fetchall()

        self.basa.commit()

        return self.result





