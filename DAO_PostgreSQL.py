import psycopg2
import logging
import datetime
import sys

class BDDPostgreSQL(object):
    def __init__(self):
        self.cursor = ""
        self.conx = ""

    def connection(self):

        db_name='outils'
        db_user='admin'
        db_host='10.0.0.17'
        db_password='Zabayo@@'
        db_port='5432'

        try:
            self.conx = psycopg2.connect(database=db_name, user=db_user,host=db_host, password=db_password, port=db_port)
            print("connextion ok")
            self.cursor = self.conx.cursor()

        except (Exception, psycopg2.InterfaceError) as error:
            logging.error(f"Echec connexion : {error}")
            return 0
        except (Exception, psycopg2.DatabaseError) as error:
            logging.error(f"Echec database : {error}")
            return 0
        return 1

    def close(self,):
        logging.info('fermeture de la base')
        self.conx.commit()
        self.conx.close()



if __name__ == '__main__':

    bd = BDDPostgreSQL()
    bd.connection()
    x = bd.cursor
    x.execute("SELECT version();")
    RT = x.fetchone()
    print(RT)
