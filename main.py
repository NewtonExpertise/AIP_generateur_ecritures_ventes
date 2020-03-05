import xlwings as xw
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from openfileGUY import Ui_MainWindow
import sys
import tempfile
import logging
from datetime import datetime
from postgre_rqt import Stat_outils
from DAO_PostgreSQL import BDDPostgreSQL
from AIPV2 import get_cell_sheet ,get_Nbligne ,get_data,datas_ectritures, plan


logging.basicConfig(filename= os.path.join(tempfile.gettempdir(),'aip.log'), level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %H:%M:%S')



class AIPImport(QtWidgets.QMainWindow):
    def __init__(self, title="default", Parent=None):
        super(AIPImport, self).__init__(Parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.tempfile = ""
        self.pathxlsx = ""
        self.pathfilexlsx = ""
        self.ui.bouton_valider.setEnabled(False)
        self.ui.bouton_valider.setStyleSheet("background-color:white;color:grey;")
        # Récupère le fichier temporaire où est stocké les infos de la dernière utilisation
        for f in os.listdir(tempfile.gettempdir()):
            if f.startswith('aip_ing_thc_urba_newton_expertise'):
                self.tempfile = tempfile.gettempdir()+'\\'+f
        if self.tempfile == "":
            x = tempfile.NamedTemporaryFile(
                prefix='aip_ing_thc_urba_newton_expertise', suffix='.txt', delete=False)
            self.tempfile = x.name

        # # # # # # # # # # # # # # # # # # # # # # # # #
        # Définition des signaux de la fenêtre.         #
        self.ui.openfile.clicked.connect(self.openfile)
        self.ui.bouton_valider.clicked.connect(self.valide)

    def openfile(self):
        # Dernier path utilisé :
        with open(self.tempfile, "r") as f:
            save = f.read()
            if save!="" and os.path.isdir(save):
                self.pathxlsx = save
            else:
                self.pathxlsx = "c://"
        filexlsx = QtWidgets.QFileDialog.getOpenFileName(self, "open", self.pathxlsx)
        self.pathfilexlsx = os.path.normcase(filexlsx[0])
        os.system('start excel.exe '+'"'+self.pathfilexlsx+'"')
        logging.basicConfig(filename= os.path.join(os.path.dirname(self.pathfilexlsx),'aip.log'), level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %H:%M:%S')
        # QtWidgets.QMessageBox().warning(self, "AIP - Import", "Merci d'attendre qu'Excel s'ouvre complètement avant de lancer le traitement.", QtWidgets.QMessageBox.Ok).button(QtWidgets.QMessageBox.Yes).setText('boutjoufzkjrbg')
        msgBox = QtWidgets.QMessageBox()
        msgBox.setWindowTitle('AIP - Import')
        msgBox.setIcon(QtWidgets.QMessageBox.Warning)
        msgBox.setText("Le fichier Excel sélectionné va s'ouvrir. Merci d'attendre l'ouverture complète avant de lancer la génération des feuilles Excel.")
        msgBox.setInformativeText('Veillez à activer la modification.')
        msgBox.setStandardButtons(QtWidgets.QMessageBox.Yes)
        msgBox.button(QtWidgets.QMessageBox.Yes).setText('Suivant')
        msgBox.exec_()
        self.ui.bouton_valider.setEnabled(True)
        self.ui.bouton_valider.setStyleSheet("background-color:#F37d00;color:white;")
        if os.path.splitext(filexlsx[0])[1]== '.xlsx':
            with open(self.tempfile, "w") as f:
                f.write(str(os.path.split(filexlsx[0])[0]))


    def valide(self):
        #on test si nous pouvons instancié le wb si ce n'est pas le cas -> msg d'erreur demandant d'ôter la protection.
        try:
            wb = xw.Book(self.pathfilexlsx)
            infoMsgBox = QtWidgets.QMessageBox()
            infoMsgBox.setIcon(QtWidgets.QMessageBox.Information)
            infoMsgBox.setWindowTitle("AIP - Import")
            infoMsgBox.setText("Traitement lancé, vos onglets vont apparaître d'ici quelques secondes.")
            infoMsgBox.setStandardButtons(QtWidgets.QMessageBox.Yes)
            infoMsgBox.button(QtWidgets.QMessageBox.Yes).setText('Suivant')
            infoMsgBox.exec_()
            self.ui.bouton_valider.setEnabled(False)
            self.ui.bouton_valider.setStyleSheet("background-color:white;color:grey;")
            msg = self.generat_import()
            msgstring=[]
            for message in msg:
                msgstring+= message
            if msgstring:
                BDDPostgre = BDDPostgreSQL()
                Postgreok = BDDPostgre.connection()
                if Postgreok:
                    stat = Stat_outils(BDDPostgre.cursor)
                    args_postgre = ["AIP_import",os.path.basename(self.pathfilexlsx)]
                    stat.espion_postgre(os.getlogin(),  datetime.now().strftime('%Y-%m-%d %H:%M:%S'), args=args_postgre)
                QtWidgets.QMessageBox().information(self, "AIP - Import", '\n'.join(msgstring), QtWidgets.QMessageBox.Ok)
            else:
                x = ["Le fichier sélectionné ne correspond pas au format habituel.", "Merci de faire un retour auprès du service informatique."]
                QtWidgets.QMessageBox().information(self, "AIP - Import", '\n'.join(x))
                logging.debug(f"Le fichier ne correspond pas au format attendu. : {self.pathfilexlsx}")
            
        except:
            x = ["Veuillez vérifier que le fichier Excel est accessible à la modification (hors lecture seul, mode protégé etc.)", "Si le problème persiste merci de faire un retour auprès du service informatique."]
            QtWidgets.QMessageBox().warning(self, "AIP - Import", '\n'.join(x))
        

        
    def generat_import(self):
        wb = xw.Book(self.pathfilexlsx)

        self.import_non_conforme = []
        msg = []
        CellSheet, code_dossier = get_cell_sheet(wb, "N° FACT")
        for sheet , coordoneecellule in CellSheet.items():
            nbligne = get_Nbligne(wb.sheets[sheet],coordoneecellule['col'])
            dataBrut = get_data(wb.sheets[sheet], coordoneecellule['row'], coordoneecellule['col'], nbligne)
            plan_dossier = plan(code_dossier[sheet])
            factures_manquantes, ecritures, nbclient, nbmontant, nbdoublon  = datas_ectritures(dataBrut, plan_dossier)

            len(factures_manquantes)

            x=0
            while x!=len(factures_manquantes):
                
                ecritures[1+x][11]=factures_manquantes[x]
                x+=1
            
            wb.sheets.add(name=('Import_'+sheet), after=wb.sheets[-1].name)
            ws = wb.sheets.active
            ws.range(1,1).value = ecritures
            # Autoformatage des cellules
            ws.autofit()
            prmFile = ['[ECRITURES]',
            "Numero=4",
            "RadicalClients=",
            "RadicalFournisseurs=",
            "RadicalClientsDest=",
            "RadicalFournisseursDest=",
            "CollectifDefautClients=41100000",
            "CollectifDefautFournisseurs=40100000",
            "Libelle=5",
            "NumPiece=8",
            "Debit=6",
            "Credit=7",
            "CodeJournal=2",
            "DateEcriture=3",
            "PositionJJEcr=1",
            "PositionMMEcr=4",
            "PositionAAEcr=9",
            "LongDateEcr=10",
            ]
            prm = open(os.path.join(self.pathxlsx,'QExport.prm'),"a")
            prm.write('\n'.join(prmFile))
            prm.close()
            
            msg.append(["Traitement terminé. "+sheet+" : ",
                    "Comptes client à affecter : "+str(nbclient),
                    "Montants à contrôler : "+str(nbmontant),
                    "Doublons à contrôler : "+str(nbdoublon), 
                    "Factures manquante à contrôler : "+str(len(factures_manquantes)),
                    ""])

        return msg




if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    w = AIPImport()
    w.show()
    sys.exit(app.exec_())
