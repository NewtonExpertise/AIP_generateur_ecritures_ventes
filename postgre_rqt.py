import pyodbc
from DAO_PostgreSQL import BDDPostgreSQL


class Stat_outils(object):
    def __init__(self, cursor):
        self.cursor = cursor

    def InfoInsertPJDBPostgre(self, collab, horodata, dossier, compte, document, nomref):
        
        sql = """
        INSERT INTO instadoc (collab,horodat,dossier,compte,document, nomref)
        VALUES (%s,%s,%s,%s,%s,%s)
        """
        data = [collab, horodata, dossier, compte, document, nomref]
        self.cursor.execute(sql, data)

        return True

    def espion_postgre(self, collab, horodata, dossier='', base='', args=''):

        sql = """
        INSERT INTO espion (collab,horodat,dossier,base, args)
        VALUES (%s,%s,%s,%s,%s)
        """
        argument = ";".join(args)
        data = [collab, horodata, dossier, base, argument]
        print(sql, data)
        self.cursor.execute(sql, data)
        return True


    def UtilisationTotal(self):
       
        sql = """
        SELECT COUNT(dossier)
        FROM instadoc
        """
        self.cursor.execute(sql)
        RT = self.cursor.fetchall()

        return RT