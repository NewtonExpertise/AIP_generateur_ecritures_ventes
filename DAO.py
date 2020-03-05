import pyodbc
import logging
import sys

class BDDAccess():
    def __init__(self):
        self.cursor = ""
                


    def connection(self,chem_base):
        self.chem_base = chem_base.lower()


        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + \
            self.chem_base
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(self.chem_base))
            self.cursor = self.conx.cursor()

        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))
        return self.cursor
        
    def close(self,):
        logging.info('fermeture de la base')
        self.conx.commit()
        self.conx.close()






if __name__ == '__main__':


    chem_base = "//srvquadra/Qappli/Quadra/DATABASE/cpta/DC/016245/qcompta.mdb"
    bddAccess = BDDAccess()
    bddAccess.connection(chem_base)
    sql = """
            SELECT
            RaisonSociale, DebutExercice, FinExercice,
            PeriodeValidee, PeriodeCloturee
            FROM Dossier1
        """
    cur = bddAccess.cursor
    cur.execute(sql)
    print(cur.fetchall())
    bddAccess.close()