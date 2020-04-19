'''
Created on Apr 18, 2020

@author: jason murphy
'''


import json
import requests
import datetime
from urllib.parse import  urlencode
import urllib
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
import logging
import configparser
from cryptography.fernet import Fernet
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import base64


full_mailrule_list = []
accessToken = "null"
options = {}

def loadConfFile():
    global options
    aConfigParser = configparser.RawConfigParser()
    try:
        aConfigParser.read('conf.ini')
        fileKey = aConfigParser.get('Basic','fileKey')
        file = aConfigParser.get('Basic','file')
        outputDir = aConfigParser.get('Basic','outputDir')
        logDir = aConfigParser.get('Basic','logDir')
        clientid = aConfigParser.get('office365settings','clientid')
        tenant_id = aConfigParser.get('office365settings','tenant_id')
        ms_graph_scope = aConfigParser.get('office365settings','ms_graph_scope')
        
        
    except:
        print("Something bad happened when reading conf.ini. Failing to local mode")

    options['fileKey'] = fileKey
    options['file'] = file
    options['outputDir'] = outputDir
    options['logDir'] = logDir
    options['clientid'] = clientid
    options['tenant_id'] = tenant_id
    options['ms_graph_scope'] = ms_graph_scope

    #Add in code options
    saltString = "4534kdsSSSSaewreDFf888"
    options['salt'] = saltString.encode()
    
    return(options)

def deriveKey(salt, fileKeyInBytes):
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000, backend=default_backend())
    theKey = base64.urlsafe_b64encode(kdf.derive(fileKeyInBytes))
    #f = Fernet(theKey)
    
    return(theKey)


def decodePassword(encodedPassword,theKey):
    #print("decodePassword")
    f = Fernet(theKey)
    decodedBytes = f.decrypt(encodedPassword.encode())
    decodedString = decodedBytes.decode()
    return(decodedString)


def encodeKeyFile(keyFile,theKey):
    aConfigParser = configparser.RawConfigParser()
    aConfigParser.read(keyFile)
    f = Fernet(theKey)
    #cycle through they keys section and check if they are all encoded. If not then encode.
    for key in aConfigParser['keys']:  
        #Check if the last character is an "="
        #print (aConfigParser.get('keys',key)[0:3])
        if not aConfigParser.get('keys',key)[0:3] == "gAA":
            #print(aConfigParser.get('keys',key))
            #print ("password for "+key+" is not encoded, so lets encode it")
            token = f.encrypt(aConfigParser.get('keys',key).encode())
            aConfigParser['keys'][key]=token.decode()
            with open(options['file'], 'w') as newconfigFile:
                aConfigParser.write(newconfigFile)
            #print (token)
            #print(f.decrypt(token))
            stringa = f.decrypt(token)
            realstring = stringa.decode()

        
        else:
            print("password for "+key+" Is already encoded")
            #print("password for "+key+" Is already encoded and it's: " + f.decrypt(aConfigParser.get('keys',key).encode()).decode() )
            #print(aConfigParser.get('keys',key))    
            #print(f.decrypt(aConfigParser.get('keys',key).encode()).decode)  

def getMailRulesBatch(listX):
    global accessToken
    global full_mailrule_list
    milliseconds = int(round(time.time() * 1000))
    return_user_list = []
    print ("Started getMailRules Function")
    logging.info("Started GetMailRulesBatch Function at "+str(milliseconds))
    
    
#     print ("IN THREAD PRINT LIST OT USE")
#     for r in listX:
#         print(r)
    
    
    with open(options['outputDir']+"RuleList="+str(milliseconds)+".json", mode='w', encoding='utf-8') as fileTemp:
        count = 1
        batchUsers = []
        totalUserCount = len(listX)
        userCount = 1
        for user in listX:
            #Add the mailrules list to the user object with the id being use in the batch
            mail={'mailrules': str(count)}
            user.update(mail)
            
            #Create the JSON Batch request
            url = "/users/"+user['id']+"/mailFolders/inbox/messageRules"
            singleUserRequest = {"id": count, "method": "GET", "url": url }
            batchUsers.append(singleUserRequest)
            
            #Count is the size of the batch
            if count == 20 or userCount==totalUserCount:
                batchUserRequest={"requests": batchUsers}
                header={"Authorization": "Bearer {}".format(accessToken), "Accept": "application/json", "Content-Type": "application/json"}
                response = requests.post("https://graph.microsoft.com/v1.0/$batch", headers=header, json=batchUserRequest)
                if 'error' in response.json():
                    if response.json()['error']['message'] == 'Access token has expired.':
                        accessToken=getNewToken()
                        header={"Authorization": "Bearer {}".format(accessToken), "Accept": "application/json", "Content-Type": "application/json"}
                        response = requests.post("https://graph.microsoft.com/v1.0/$batch", headers=header, json=batchUserRequest)
                
                #Reset some stuff
                count = 0
                batchUsers = []
                batchUserRequest={}
                
                #MERGE Mailrules into the Users in listX
                response_dict = response.json()['responses']
                for each in response_dict:
                    for userX in listX:
                        if "mailrules" in userX and userX['mailrules'] == each['id']:
                            #Update the user object with the mail rules
                            #check if Error is in the response
                            if 'error' in each['body']:
                                mail={'mailrules': {'code' : each['body']['error']['code'], 'message' : each['body']['error']['message']}}
                                userX.update(mail)
                            else:
                                userX.update(mailrules = each['body']['value'])
                                
                            json.dump(userX,fileTemp)
                            fileTemp.write("\n")
                            return_user_list.append(userX)
                            break; #Needed
            
                
            
            count = count + 1
            userCount = userCount + 1
        full_mailrule_list.append(return_user_list)
        return(return_user_list)
            

    
def getNewToken():
    global accessToken
    print ("TOKEN EXPIRED GET NEW TOKEN")
    logging.info("Token Expired, get new token")
    clientsecret = options['clientsecret']
    clientid = options['clientid']
    tenant_id = options['tenant_id']
    ms_graph_scope = options['ms_graph_scope']
    
    url = "https://login.microsoftonline.com/"+ tenant_id +"/oauth2/v2.0/token"
    #dataPost={"client_id: "+clientid+", scope: " + ms_graph_scope +",client_secret: "+ clientsecret +", grant_type: client_credentials"}
    
    dataPost={'client_id':clientid,
               'scope':ms_graph_scope,
               'client_secret':clientsecret,
                'grant_type':'client_credentials'}
    
    response = requests.post(url, data=dataPost)
    #print (response.json())
    
    resp_dict = response.json()
    accessToken=resp_dict['access_token']
    expires_in=resp_dict['expires_in']

    #print ("extracted values")
    #print ("access Token= " + accessToken)
    #print ("expires_in= "+ str(expires_in))

    with open("o365accesstoken", mode='r+') as fileAT:
        fileAT.write(accessToken)
        fileAT.truncate()
    return accessToken


def main():
    global accessToken
    global full_mailrule_list
    sys.stdout.reconfigure(encoding='utf-8')
    currentTime=datetime.datetime.now()
    print("Job Started " + str(currentTime))
    

    
    #Load up conf.ini into options
    options = loadConfFile()
    outputDir = options['outputDir']
    logDir = options['logDir']
    
    today = currentTime.strftime("%m-%d-%Y-%H%M")
    logging.basicConfig(filename=logDir+today+".log",level=logging.INFO)
    logging.info(str(currentTime) + ' start main')
    
    keyFile = options['file']
    temp = options['fileKey']
    fileKeyInBytes = temp.encode()
    salt=options['salt']   
    
       
    theKey = deriveKey(salt, fileKeyInBytes)
   
    #Check if all the passwords in the keyfile are encoded
    encodeKeyFile(keyFile, theKey)

    #Load client secret and add to options{}
    aConfigParser = configparser.RawConfigParser()
    aConfigParser.read(keyFile)
    encodedPassword = aConfigParser.get('keys','clientsecret')
    decodedPassword = decodePassword(encodedPassword, theKey)
    options['clientsecret'] = decodedPassword

    isValidAccessToken=False
    #Load saved access token
    with open("o365accesstoken", mode='r+') as fileAT:
        accessToken=fileAT.read()
        #reset pointer to beginning of file
        fileAT.seek(0)
    
    header = {"Authorization": "Bearer {}".format(accessToken)}
    response = requests.get("https://graph.microsoft.com/v1.0", headers=header)
    #print (response.status_code)
    #print (response.headers)
    #print (response.json())
        
    
    
    #If response code is 401 then get a new access token
    if  response.status_code == 401:
        #print ("GETTING NEW ACCESS TOKEN")
        resp_dict = response.json()
        print (resp_dict['error']['message'])
        
        accessToken = getNewToken()
        isValidAccessToken=True
    
    if response.status_code == 200:
        print ("TOKEN IS STILL GOOD!")
        isValidAccessToken=True
        
       
    
    if isValidAccessToken:
        print ("TOKEN VALID, SO RUNNING STUFF!")
        logging.info("Token Valid, Running Stuff")
        header = {"Authorization": "Bearer {}".format(accessToken)}
        

        #global full_mailrule_list
        #FIRST STEP ---- Get all the users and store them
        response = requests.get("https://graph.microsoft.com/v1.0/users", headers=header)
        print (response.status_code)
        print (response.headers)
        
        resp_dict = response.json()
        full_user_list_working = []
        
        
        with open(options['outputDir']+"UserList.json", mode='w') as fileShortLived:
            tempDict = resp_dict['value']
            for user in tempDict:
                json.dump(user,fileShortLived)
                fileShortLived.write("\n")
                full_user_list_working.append(user)
            count = 1
            while '@odata.nextLink' in resp_dict:
                count = count + 1
                print ("Page " + str(count))
                nextUrl = resp_dict['@odata.nextLink']
                response = requests.get(nextUrl, headers=header)
                resp_dict = response.json()
                tempDict = resp_dict['value']
                for user in tempDict:
                    json.dump(user,fileShortLived)
                    fileShortLived.write("\n")
                    full_user_list_working.append(user)
        
        #SECOND STEP --- Create threads to work off the user_dict
        #Pull some off the full_user_dict_working and send to a thread to be worked on
        count = 1
        print ("Full User Thread length= " + str(len(full_user_list_working)))
        logging.info("Number of Users to pull mail rules for: " + str(len(full_user_list_working)))
        with ThreadPoolExecutor(max_workers=36) as executor:
            while full_user_list_working != []:
                x = 1000
                list_send_to_thread = []
                while x > 0 and full_user_list_working != []:
                    tempuser = full_user_list_working.pop()
                    list_send_to_thread.append(tempuser)
                    x = x - 1
                future = executor.submit(getMailRulesBatch,(list_send_to_thread))
            
            #print (count)
            count = count + 1
        
        #Print the full mail rule list
        with open(options['outputDir']+"MailRuleList"+today+".json", mode='w') as fileTemp: 
            for eachJob in full_mailrule_list:
                for eachUser in eachJob:
                    json.dump(eachUser,fileTemp)
                    fileTemp.write("\n")
if __name__ == '__main__':
    main()