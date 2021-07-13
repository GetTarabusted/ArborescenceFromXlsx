# Made by Olivier Sordoillet @ Cloudixio

# Les imports qui vont bien
import pandas as pd
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import client as AOclient, file as AOfile, tools

def main():
    DRIVE = OAuth()

    # Variables
    # ID des dossiers root
    # folderid_partenaires = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    folderid_paris = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    folderid_lyon = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

    # Création des dataframes par villes et création des dossiers
    dfs = getDataframes("C:/Path/to/excel.xlsx")
    clientList_paris = generateClientList(dfs[1])
    createFolders(DRIVE, clientList_paris, folderid_lyon, "lyon")
    clientList_lyon = generateClientList(dfs[0])
    createFolders(DRIVE, clientList_lyon, folderid_paris, "paris")

def OAuth():
    ''' Permet l'authentification et gère les droits de modification du compte google '''
    SCOPES = 'https://www.googleapis.com/auth/drive'
    store = AOfile.Storage('storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = AOclient.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    DRIVE = build('drive', 'v3', http=creds.authorize(Http()))
    return DRIVE

def getDataframes(path):
    '''Grabs the xls path and output the clients in an array of dataframes. We consider one city per sheet of excel'''
    xls = pd.ExcelFile(path)
    df_paris = pd.read_excel(xls, sheet_name="PGC Paris")
    df_lyon = pd.read_excel(xls, sheet_name="PGC Lyon")
    df_paternaire = pd.read_excel(xls, sheet_name="Partenaire")
    dataframes = [df_paris, df_lyon, df_paternaire]
    return dataframes

def generateClientList(df):
    '''
    Convert the dataframe in dict with name:email and get rids of duplicates \n
    The top row of each sheet much be labelised with Account.Name for the name of the client and Account.Owner.Email for the email of the sales rep associated
    '''
    clients = {}
    df.drop_duplicates(inplace= True) # Inplace let the method be destructive
    for index, row in df.iterrows():
        clients[row["Account.Name"]] = row["Account.Owner.Email"]
    print("Nombre de clients:", len(clients))
    return clients

def createFolder(DRIVE, foldername, parent_id):
    '''creates a folder'''
    file_metadata = { 
        'name': foldername,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id],
        }
    client_folder = DRIVE.files().create(body=file_metadata, fields= 'id').execute() # Fields arguments is the response stored given by the API
    return client_folder

def createFolders(DRIVE, clientList, root_id, city):
    '''
    create the folder tree and give the right permissions \n
    clientList : dict of client client:emailowner,
    root_id : id of the root folder where to create the arborescence
    city : "lyon" or "paris"
    '''
    # Liste qui sert de boucle pour créer les dossiers
    folders = ["02 - Conformité", "03 - KYC", "04 - Contrat", "01 - Propal"]
    # Dictionnaire avec compte et id de permission d'écriture associé
    permissions_list = {
        "vente.pgc.paris" : "XXXXXXXXXXXXXXXXXXXX",
        "vente.pgc.lyon" : "YYYYYYYYYYYYYYYYYYYY",
        }
     # On les stocke pour ne pas les appeler à chaque fois, et sert pour suivre l'avancement du déploiement
    compteur_client = 1
    len_client = len(clientList)

    # Création des dossiers pour chaque client
    for client in clientList:
        # Tracker du nb de clients
        print("Client %i of %i" % (compteur_client, len_client))
        compteur_client += 1

        # Création du dossier client
        client_folder = createFolder(DRIVE, client, root_id)
        print('Folder created: %s' % client)
        print("Folder's ID: %s" % client_folder.get('id'))

        # Suppression du droit pour le permissionId renseigné (PGC vente paris ou lyon)
        DRIVE.permissions().delete(fileId= client_folder.get('id'), permissionId = permissions_list['vente.pgc.'+city]).execute()
        print("Permission deleted for %s" % 'vente.pgc.'+city)

        # Ajout de la permission en reader pour le proprio SF 
        createPermission(DRIVE, client_folder.get('id'), clientList[client], 'reader')
        print("User %s has now access to this folder \n" % clientList[client])

        # Pour chaque client, on crée les 4 mêmes dossiers stockés dans folders
        # On va stocker les id des dossiers qu'on créé pour y modifier les droits après
        folder_ids = []
        for folder in folders:
            subfolder = createFolder(DRIVE, folder, client_folder.get('id'))
            print("Subfolder created: %s" % folder)
            folder_ids.append(subfolder.get('id'))
        updatePermissions(DRIVE, folder_ids, city, clientList[client])



def updatePermissions(DRIVE, idList, city, SF_EmailAdress):
    """
    Takes the IDs in idList and apply the permissions to each of them, depending on the city
    SF_EmailAdress: email adress of the sales rep associated with the client so he can be added as writer on Propal
    On part du principe que tout le monde est en LECTURE de base SAUF Back-Office donc aucun update à faire: permissions().create pour augmenter/créer l'accès,
    permissions().delete pour supprimer l'accès et permissions().update sert juste pour diminuer l'accès ce qui n'est jamais le cas ici
    """
    permissions_list = { #Insert permissions IDs here
        "direction.commerciale.lyon" : "AAAAAAAAAAAAAAAAAAAA",
        "direction.commerciale.paris" : "BBBBBBBBBBBBBBBBBBBB",
        "risk.compliance" : "CCCCCCCCCCCCCCCCCCCC",
        "middle-office" : "DDDDDDDDDDDDDDDDDDDD",
        "back-office" : "EEEEEEEEEEEEEEEEEEEE",
        "finance" : "FFFFFFFFFFFFFFFFFFFF"
        }
    batch = DRIVE.new_batch_http_request(callback=callback)
    # Conformité
    batch.add(DRIVE.permissions().delete(fileId= idList[0], permissionId = permissions_list["finance"]))
    batch.add(DRIVE.permissions().create(fileId= idList[0], body = {"role" : "writer", "type": "group", "emailAddress" : "emailAdress@riskcompliance.com"}, sendNotificationEmail = False))
    # KYC
    batch.add(DRIVE.permissions().delete(fileId= idList[1], permissionId = permissions_list["finance"]))
    batch.add(DRIVE.permissions().create(fileId= idList[1], body = {"role" : "writer", "type": "group", "emailAddress" : "emailAdress@riskcompliance.com"}, sendNotificationEmail = False))
    batch.add(DRIVE.permissions().create(fileId= idList[1], body = {"role" : "writer", "type": "group", "emailAddress" : "emailAdress@middleoffice.com"}, sendNotificationEmail = False))
    # Contrat
    batch.add(DRIVE.permissions().create(fileId= idList[2], body = {"role" : "writer", "type": "group", "emailAddress" : "emailAdress@middleoffice.com"}, sendNotificationEmail = False))
    # Propal
    batch.add(DRIVE.permissions().create(fileId= idList[3], body = {"role" : "writer", "type": "user", "emailAddress" : SF_EmailAdress}, sendNotificationEmail = False))
    batch.add(DRIVE.permissions().create(fileId= idList[3], body = {"role" : "writer", "type": "group", "emailAddress" : "emailAdressd@directioncommercialecity.com"}, sendNotificationEmail = False))
    batch.add(DRIVE.permissions().delete(fileId= idList[3], permissionId = permissions_list["risk.compliance"]))
    batch.add(DRIVE.permissions().delete(fileId= idList[3], permissionId = permissions_list["back-office"]))
    batch.add(DRIVE.permissions().delete(fileId= idList[3], permissionId = permissions_list["finance"]))

    # Lancement du batch
    print("Submitting permissions...")
    batch.execute()
    print("Permissions OK")
    print("\n------------------------\n")

def createPermission(DRIVE, fileId, email, role):
    """Accorde une permission sur fileId à l'email renseigné"""
    permission = {
        "type": "group",
        "role": role,
        "emailAddress": email,
        }
    return DRIVE.permissions().create(fileId = fileId, body = permission, sendNotificationEmail = False).execute()

def callback(request_id, response, exception):
    if exception:
        raise Exception(exception) 

if __name__ == "__main__":
    main()