import xlwings as xw
import pprint
from datetime import datetime
import calendar
from DAO import BDDAccess
from quadraenv import QuadraSetEnv

pp = pprint.PrettyPrinter(indent=4)
listeSheet = ['ING','AIP','THC','URBA', 'BIM']

# wb = xw.Book(r'V:\Mathieu\ventes_aip\test.xlsx')
# numéros quadra
AIPQuadra = "016422"
INGQuadra = "016421"
THCQuadra = "016423"
URBAQuadra = "016799"
BIMQuadra = "000983"
# numéros de comptes
ht55 = '70610100'
ht10 = '70610500'
ht20 = '70610600'
tva55 = '44571200'
tva10 = '44571500'
tva20 = '44571600'
# code journal
codeJDV = 'VE'



def get_cell_sheet(wb , cell_value):
    """
    workBook et valeur de la première cellule du tableau a localiser
    retourn tous les nom de feuille et a localisation de la 1ere cellule
    (ligne,colone)
    """
    CellSheet = {}
    code_dossier = {}
    for ws in wb.sheets:
        for row in range(1,15):
            for col in range(1,10):
                if ws.range(row,col).value == cell_value:
                    if ws.name=='ING':
                        code_dossier[ws.name] = INGQuadra
                        CellSheet[ws.name]={"row":row,"col":col}
                    if ws.name=='AIP':
                        code_dossier[ws.name] = AIPQuadra
                        CellSheet[ws.name]={"row":row,"col":col}
                    if ws.name=='THC':
                        code_dossier[ws.name] = THCQuadra
                        CellSheet[ws.name]={"row":row,"col":col}
                    if ws.name=='URBA':
                        code_dossier[ws.name] = URBAQuadra
                        CellSheet[ws.name]={"row":row,"col":col}
                    if ws.name=='BIM':
                        code_dossier[ws.name] = BIMQuadra
                        CellSheet[ws.name]={"row":row,"col":col}

                    # CellSheet[ws.name]={"row":row,"col":col}
    return CellSheet , code_dossier

def get_Nbligne(ws, col):
    """
    Prend un workSheet et la localisation de la première colone d'un tableau.
    Retourn la longeur du tableau en nombre de ligne.
    """
    return ws.cells(ws.api.rows.count, col).end(-4162).row

def get_data(ws, row, col, nbligne):
    """
    Prend un workSheet, la localisation de la première cellule d'un tableau row et col , nb de ligne
    """
    datasbrut = ws.range((row,col),(nbligne,col+30)).value
    datas = {}
    indexfact = datasbrut[0].index('N° FACT')
    indexobjet = datasbrut[0].index('OBJET')
    indexclient = datasbrut[0].index('CLIENT')
    indexHT = datasbrut[0].index('MONTANT HT')
    indexTVA55 = datasbrut[0].index('TVA 5.5 %')
    indexTVA20 = datasbrut[0].index('TVA 20 %')
    indexTVA10 = datasbrut[0].index('TVA 10 %')
    indexTTC = datasbrut[0].index('TTC')
    i=1
    for data in  datasbrut:
        try:
            float(data[indexTTC])
        except:
            continue
        if data[indexfact] and data[indexclient]:
            if data[indexTVA55]:
                TVA = data[indexTVA55]
                TxTVA = 1.055
            if data[indexTVA10]:
                TVA = data[indexTVA10]
                TxTVA = 1.1
            if data[indexTVA20]:
                TVA = data[indexTVA20]
                TxTVA = 1.2
            datas[i]={'fact' : data[indexfact],
                    'objet' : data[indexobjet],
                    'client' : data[indexclient],
                    'HT' : data[indexHT],
                    'TVA' : TVA,
                    'TxTVA': TxTVA,
                    'TTC' : data[indexTTC]}
        else:
            continue
        i+=1
    return datas

def controle_montant(ht, tva, ttc, taux):
    """
    controle de la cohérence des montant TTC HT TVA
    """
    x=''
    try:
        float(ht)
        float(tva)
        float(ttc)
        float(taux)
        x = True
    except:
        x = False
    if x:
        ttcInterval = [round(ttc+0.01,2),ttc,round(ttc-0.01,2)]

        if round(float(ht)+float(tva),2) in ttcInterval and round((float(ht)*taux),2) in ttcInterval:
            return True
        else:
            return False
    else:
        return False

def affectation_compte(taux):
    """
    affect un compte en fonction de son taux de tva
    """

    if taux == 1.1:
        return {'ht':ht10, 'tva' :tva10}
    elif taux == 1.2:
        return {'ht':ht20, 'tva' :tva20}
    elif taux == 1.055:
        return {'ht':ht55, 'tva' :tva55}
    else:
        return {'ht':None, 'tva' :None}

def datas_ectritures(datadict , plan_comptable):
    """
    prend en argument un dictionnaire : 
    dict[clé]{n°fact, libClt, libEcriture, HT, TVA55, TVA10, TVA20, TTC,}
    """

    incremente_fact=0
    list_doublon = []
    mois = 0
    ctrl_doublons = []
    liste_import = []
    facture_manquante= []
    liste_import.append(['N°ECRITURE','JOURNAL','DATE','COMPTE','LIBELLE','DEBIT','CREDIT','PIECE', "COMPTE CLT", "MONTANT", "DOUBLON", "FAC MANQUANTE", "Veillez à supprimer la ligne 1 avant de générer votre import. QExport.prm a été généré, il ne vous reste qu’à l’appeler via l’onglet « compléments » pour générer votre Qimport."]) #entête
    nbclient = 0
    nbmontant = 0
    nbdoublon = 0
    for cle , val in datadict.items():
        client = ""
        montant = ""
        doublon = ""
        #controle des doublons 
        if val['fact'] not in ctrl_doublons:
            ctrl_doublons.append(val['fact'])
            doublon = None
        else:
            doublon = "x"
            list_doublon.append(val['fact'])


        # affectation des comptes en fontion des taux de TVA
        compte = affectation_compte(val['TxTVA'])
        horstaxe = {'compte':compte['ht'],'montant':val['HT']}
        tva = {'compte':compte['tva'],'montant':val['TVA']}

        # date de la facture.
        datefact = datetime.strptime(val['fact'][:5],'%y.%m')
        datefact = datefact.replace(day=calendar.monthrange(datefact.year,datefact.month)[1])

        # libellé
        libelle = val['objet'][:15]+' '+val['client'][:14]

        # Affectation compte comptable en fonction de la correspondance des libellés du PlanC. et des infos du tableau excel
        compte_client=""
        for compte , info_client in plan_comptable.items():
            if info_client['intitule'] == val['client']:
                compte_client = compte

        # controle client , montant, doublons
        if not controle_montant(val['HT'], val['TVA'], val['TTC'], val['TxTVA']):
            montant = "x"
        else:
            montant = None

        if not compte_client:
            client = "x"
        else:
            client = None




        #Controle factures manquantes
        annee_fact = val['fact'].split('.')[0]
        mois_fact = int(val['fact'].split('.')[1])
        num_fact = int(val['fact'].split('.')[2])
        if mois_fact != mois: # si le mois de l'itération précédante coorrespond
            mois = mois_fact
            incremente_fact = 1
            if incremente_fact != num_fact:
                # tant que le num fact != increment on ajout la fact manquante
                while num_fact > incremente_fact:
                    if val['fact'] in list_doublon:
                        incremente_fact += 1
                    else:
                        facture_manquante.append(annee_fact+'.'+str(mois_fact).zfill(2)+'.'+str(incremente_fact).zfill(2))
                        incremente_fact += 1
                incremente_fact += 1
            else:
                incremente_fact += 1
        else:
            if incremente_fact != num_fact:
                # tant que le num fact != increment on ajout la fact manquante
                while num_fact > incremente_fact:
                    if val['fact'] in list_doublon:
                        incremente_fact += 1
                    else:
                        facture_manquante.append(annee_fact+'.'+str(mois_fact).zfill(2)+'.'+str(incremente_fact).zfill(2))
                        incremente_fact += 1
                incremente_fact += 1
            else:
                incremente_fact += 1

        if val['TTC'] > 0 :
            liste_import.append([cle, codeJDV, datefact, compte_client, libelle, val['TTC'], None, val['fact'], client , montant , doublon , None, None]) #ligne compte TTC
            liste_import.append([cle, codeJDV, datefact, horstaxe['compte'], libelle, None, horstaxe['montant'], val['fact'], client , montant , doublon , None, None]) #ligne compte HT
            liste_import.append([cle, codeJDV, datefact, tva['compte'], libelle, None, tva['montant'], val['fact'], client , montant , doublon , None, None]) #ligne compte TVA
        else:
            liste_import.append([cle, codeJDV, datefact, compte_client, libelle, None, -val['TTC'], val['fact'], client , montant , doublon , None, None]) #ligne compte TTC
            liste_import.append([cle, codeJDV, datefact, horstaxe['compte'], libelle, -horstaxe['montant'], None, val['fact'], client , montant , doublon , None, None]) #ligne compte HT
            liste_import.append([cle, codeJDV, datefact, tva['compte'], libelle, -tva['montant'], None, val['fact'], client , montant , doublon , None, None]) #ligne compte TVA


        if client == 'x':
            nbclient +=1
        if montant == 'x':
            nbmontant +=1
        if doublon == 'x':
            nbdoublon +=1

    liste_import


    return facture_manquante, liste_import , nbclient, nbmontant, nbdoublon

def plan(code_client):
    quadraenv = QuadraSetEnv()
    quadraenv.Millesime(code_client)[0][1]
    BDD = BDDAccess()
    cursor = BDD.connection(quadraenv.Millesime(AIPQuadra)[0][1])
    plan={}
    # Plan comptable
    sql = """
        SELECT
        Numero, Type, Intitule, NbEcritures, ProchaineLettre
        FROM Comptes
    """
    cursor.execute(sql)
    for num, typ, intit, nbecr, lettr in cursor.fetchall():
        plan.update(
            {
                num : {
                    "type" : typ,
                    "intitule" : intit,
                    "nbecr" : nbecr,
                    "lettrage" : lettr
                }
        })

    BDD.close()

    return plan




if __name__ == "__main__":
    import pprint
    pp = pprint.PrettyPrinter(indent=4)

    wb = xw.Book(r'V:\Mathieu\ventes_aip\ressources\test.xlsx')
    dataBrut = {}


    CellSheet, code_dossier = get_cell_sheet(wb, "N° FACT")

    for sheet , coordoneecellule in CellSheet.items():
   
        nbligne = get_Nbligne(wb.sheets[sheet],coordoneecellule['col'])

        dataBrut = get_data(wb.sheets[sheet], coordoneecellule['row'], coordoneecellule['col'], nbligne)

        plan_dossier = plan(code_dossier[sheet])

        factures_manquantes, ecritures = datas_ectritures(dataBrut, plan_dossier)
  
 



    # ##################
    # ##pour quadra env


    # AIPQuadra = "016422"
    # INGQuadra = "016421"
    # THCQuadra = "016423"
    # URBAQuadra = "016799"

    
    # quadraenv = QuadraSetEnv()
    # quadraenv.Millesime(AIPQuadra)[0][1]
    # BDDAccess = BDDAccess()
    # BDDAccess.connection(quadraenv.Millesime(AIPQuadra)[0][1])

    # quadracompta = QueryCompta(BDDAccess.cursor)
    # quadracompta.connect()




