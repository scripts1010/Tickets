import requests, json, urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url= 'https://10.100.107.225'

# Tipo Text es el numero 1, Tipo Numerico es el numero 2 y Tipo Lista de Valores es 4

def sessionTokenArcher(url):
        headers = {"Accept" : "application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" , "Content-Type" : "application/json"}
        user="apitest"
        password="Cyberark1!"
        # user = "sysadmin"
        # password = "C4landriA!"
        body = {
                "InstanceName":"dev",
                "Username":user,
                "Password":password
                }
        response = requests.post(url + '/platformapi/core/security/login', verify=False, data=json.dumps(body) ,headers=headers)
        #print (response)
        #print (response.content)
        sessionToken = response.json()['RequestedObject']['SessionToken']
        print (sessionToken)
        return sessionToken
        
#sessionTokenArcher('https://10.100.107.225')
#levelId: 214
#fieldId : 15353
#AppId: 417
def createRecord(SessionToken,url): 
    headers = {"Accept" : "application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" , "Content-Type" : "application/json", "Authorization": "Archer session-id=" + SessionToken}
    body={
        "Content": {
            "LevelId" : 1448,
            "FieldContents":
            {
                "31976":
                    {
                        "Type" : "1", 
                        "Value" : "Hola, como andan? Version 2.0", 
                        "FieldId" : "31976"
                    }
                    }
                    }
        }            
    response = requests.post(url + '/platformapi/core/content', verify=False, data=json.dumps(body) ,headers=headers)
    print (response.content)
    isSuccessful = response.json()['IsSuccessful']
    print (isSuccessful)
    if not isSuccessful:
       errorMessage = response.json()['ValidationMessages'][0].get('ResourcedMessage')
       print (str(errorMessage))


def getApplication(SessionToken, url):
    # dentro de aplicaciones, desde IE, apoyo el mouse sobre el app de ticket y me dice el id de la app. No es lo mismo que level id
    appId = '1632'
    headers = {"Accept" : "application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" , "Content-Type" : "application/json", "Authorization": "Archer session-id=" + SessionToken, "X-Http-Method-Override":"GET"}
    body = {}
    response = requests.post(url + '/platformapi/core/system/fielddefinition/application/' + appId, verify=False, data=json.dumps(body), headers=headers)
    print(response)
    print(response.content)
    for field in response.json():
        print ("Campo")
        print (field['RequestedObject'].get('Name'),field['RequestedObject'].get('Type'), field['RequestedObject'].get('Id') )
        print ("Campo")

createRecord(sessionTokenArcher(url),url)
#getApplication(sessionTokenArcher(url), url)

