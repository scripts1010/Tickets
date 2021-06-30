#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import poplib
import ConfigParser
import sys
import urllib3
import json
import imaplib
import email
import socket
from email import parser
from datetime import date
from email.header import decode_header
import xml.etree.ElementTree as ET
import traceback
import random
import logging
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

configParser = ConfigParser.RawConfigParser()
#configFilePath = sys.argv[1]
configFilePath = "./env.cfg"
configParser.read(configFilePath)

ticketsIds = {
    # Ultimo numero de url (/apps/ArcherApp/Home.aspx#search/70/75/542/false/default/368)
    "levelId": 218,
    "asunto": 24100,
    "cc": 24095,
    "de": 22508,
    "ddc": 23864,
    "para": 24094,
    "realmentececticket": 24066,
    "rticketno": 72098,
    "trabierto": 24111,
    "trabiertosi": 72137,
    "cerrarticket": 23854,
    "cticketno": 72100,
    "estado": 23852,
    "estadocerrado": 71959,
    "reabrirt": 23852,
    "rticketsi": 72129,
    "delacola": 17022,
    # "clienteTicket": 24499, Test
    "clienteTicket": 25954,
    "servicios": 66631,
    "soc": 72123,
    "soporte": 66629,
    "instalacion": 66628,
    "mesadeayuda": 66630,
    # "socclaro": 73372, Test
    "socclaro": 74620
}
ddcIds = {
    # Ultimo numero de url (/apps/ArcherApp/Home.aspx#search/70/75/542/false/default/368)
    "levelId": 2266,
    "adjunto": 23875,
    "asunto": 23874,
    "cuerpo": 23876,
    "de": 23877,
    "para": 23878,
    "cc": 23879,
    "obs": 23881,
    "paramail": 23883,
    "ccmail": 23884,
    "paramailref": 25531,
    "ccmailref": 25533
}


def contentIdDetalle(sessionToken, mail):

    try:
        userSessionHeader2 = {'Content-Type': 'text/xml;charset=utf-8',
                              'SOAPAction': 'http://archer-tech.com/webservices/ExecuteSearch'}
        searchOptions = \
            """<?xml version="1.0" encoding="utf-8"?>
                    <SearchReport>
                        <DisplayFields>
                            <DisplayField>531</DisplayField>
                            <DisplayField>579</DisplayField>
                        </DisplayFields>
                        <PageSize>50</PageSize>
                        <Criteria>
                            <Keywords />
                            <Filter>
                                <OperatorLogic />
                                <Conditions>
                                    <TextFilterCondition>
                                        <Field>579</Field>
                                        <Operator>Contains</Operator>
                                        <Value>""" + mail + """</Value>
                                    </TextFilterCondition>
                                </Conditions>
                            </Filter>
                            <ModuleCriteria>
                                <Module>84</Module>
                                <IsKeywordModule>True</IsKeywordModule>
                                <BuildoutRelationship>Union</BuildoutRelationship>
                                <SortFields>
                                    <SortField>
                                        <Field>531</Field>
                                        <SortType>Ascending</SortType>
                                    </SortField>
                                </SortFields>
                            </ModuleCriteria>
                        </Criteria>
                    </SearchReport>"""

        body2 = \
            """<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
            <ExecuteSearch xmlns="http://archer-tech.com/webservices/">
            <sessionToken>""" + sessionToken + """</sessionToken>
            <searchOptions><![CDATA[""" + searchOptions + """]]></searchOptions>
            <pageNumber>1</pageNumber>
            </ExecuteSearch>
            </soap:Body>
            </soap:Envelope>
            """

        response = requests.post(baseurl + '/ws/search.asmx',
                                 verify=False, data=body2, headers=userSessionHeader2)
        tree2 = ET.fromstring(response.content)
        reg = tree2.find(
            './/{http://archer-tech.com/webservices/}ExecuteSearchResult').text
        regf = reg.replace("""<?xml version="1.0" encoding="utf-16"?>""",
                           """<?xml version="1.0" encoding="utf-8"?>""")
        regf2 = regf.encode('utf-8', errors='ignore')
        lstUsuarios = ET.fromstring(regf2)
        record = lstUsuarios.find('Record')
        if record is None:
            print(None)
        else:
            contentId = record.get('contentId')
            return contentId
    except:
        logging.error("Ha sucedido un error en: 'contentIdDetalle'")


def verificarAsunto(sessionToken, asunto):
    try:
        userSessionHeader2 = {'Content-Type': 'text/xml;charset=utf-8',
                              'SOAPAction': 'http://archer-tech.com/webservices/ExecuteSearch'}
        searchOptions = \
            """<?xml version="1.0" encoding="utf-8"?><SearchReport><DisplayFields><DisplayField>16719</DisplayField><DisplayField>24100</DisplayField></DisplayFields><PageSize>50</PageSize><Criteria><Keywords /><Filter><OperatorLogic /><Conditions><TextFilterCondition><Field>24100</Field><Operator>Contains</Operator><Value>""" + asunto + \
            """</Value></TextFilterCondition></Conditions></Filter><ModuleCriteria><Module>425</Module><IsKeywordModule>True</IsKeywordModule><BuildoutRelationship>Union</BuildoutRelationship><SortFields><SortField><Field>16719</Field><SortType>Ascending</SortType></SortField></SortFields></ModuleCriteria></Criteria></SearchReport>"""

        body2 = \
            """<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
            <ExecuteSearch xmlns="http://archer-tech.com/webservices/">
            <sessionToken>""" + sessionToken + """</sessionToken>
            <searchOptions><![CDATA[""" + searchOptions + """]]></searchOptions>
            <pageNumber>1</pageNumber>
            </ExecuteSearch>
            </soap:Body>
            </soap:Envelope>
            """
        body2 = body2.encode("utf-8")
        response = requests.post(baseurl + '/ws/search.asmx',
                                 verify=False, data=body2, headers=userSessionHeader2)

        tree2 = ET.fromstring(response.content)
        reg = tree2.find(
            './/{http://archer-tech.com/webservices/}ExecuteSearchResult').text
        regf = reg.replace("""<?xml version="1.0" encoding="utf-16"?>""",
                           """<?xml version="1.0" encoding="utf-8"?>""")
        regf2 = regf.encode('utf-8')
        lstUsuarios = ET.fromstring(regf2)
        record = lstUsuarios.find('Record')
        if record is None:
            print(None)
        else:
            contentId = record.get('contentId')
            return contentId
    except:
        logging.error("Ha sucedido un error en: 'verificarAsunto'")


def idCliente(sessionToken, idAppCliente, emailCliente):
    try:
        userSessionHeader2 = {'Content-Type': 'text/xml;charset=utf-8',
                              'SOAPAction': 'http://archer-tech.com/webservices/ExecuteSearch'}
        searchOptions = \
            """<?xml version="1.0" encoding="utf-16"?><SearchReport><DisplayFields><DisplayField>2972</DisplayField></DisplayFields><PageSize>50</PageSize><Criteria><Keywords /><Filter><OperatorLogic /><Conditions><TextFilterCondition><Field>25953</Field><Operator>Contains</Operator><Value>""" + emailCliente + \
            """</Value></TextFilterCondition></Conditions></Filter><ModuleCriteria><Module>191</Module><IsKeywordModule>True</IsKeywordModule><BuildoutRelationship>Union</BuildoutRelationship><SortFields><SortField><Field>2972</Field><SortType>Ascending</SortType></SortField></SortFields></ModuleCriteria></Criteria></SearchReport>"""

        body2 = \
            """<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
            <ExecuteSearch xmlns="http://archer-tech.com/webservices/">
            <sessionToken>""" + sessionToken + """</sessionToken>
            <searchOptions><![CDATA[""" + searchOptions + """]]></searchOptions>
            <pageNumber>1</pageNumber>
            </ExecuteSearch>
            </soap:Body>
            </soap:Envelope>
            """

        body2 = body2.encode("utf-8")
        response = requests.post(baseurl + '/ws/search.asmx',
                                 verify=False, data=body2, headers=userSessionHeader2)
        print(response.content)
        tree2 = ET.fromstring(response.content)
        reg = tree2.find(
            './/{http://archer-tech.com/webservices/}ExecuteSearchResult')
        if reg is None:
            return None
        else:
            reg = reg.text
        regf = reg.replace("""<?xml version="1.0" encoding="utf-16"?>""",
                           """<?xml version="1.0" encoding="utf-8"?>""")
        regf2 = regf.encode('utf-8')
        lstUsuarios = ET.fromstring(regf2)
        record = lstUsuarios.find('Record')
        if record is None:
            print(None)
        else:
            contentId = record.get('contentId')
            return contentId
    except:
        logging.error("Ha sucedido un error en: 'idCliente'")


def postAttachment(url, sessionToken, attachmentName, attachmentBytes):
    try:
        url = url+'/api/core/content/attachment'
        data = {
            "AttachmentName": attachmentName,
            "AttachmentBytes": attachmentBytes
        }
        headers = {
            'Accept': 'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Authorization': 'Archer session-id='+sessionToken,
            'Content-Type': 'application/json'
        }
        response = requests.post(url, verify=False, headers=headers, json=data)
        return response.json()
    except:
        logging.error("Ha sucedido un error en: 'postAttachment'")


def existeTicketId(sessionToken, ticketId):
    try:

        userSessionHeader2 = {'Content-Type': 'text/xml;charset=utf-8',
                              'SOAPAction': 'http://archer-tech.com/webservices/ExecuteSearch'}
        print(ticketId)
        searchOptions = \
            """<?xml version="1.0" encoding="utf-16"?><SearchReport><DisplayFields><DisplayField>16719</DisplayField><DisplayField>24088</DisplayField><DisplayField>23852</DisplayField><DisplayField>16743</DisplayField><DisplayField>16715</DisplayField><DisplayField>23874</DisplayField><DisplayField>23871</DisplayField></DisplayFields><PageSize>50</PageSize><Criteria><Keywords /><Filter><OperatorLogic /><Conditions><TextFilterCondition><Field>16719</Field><Operator>Contains</Operator><Value>""" + \
            ticketId + """</Value></TextFilterCondition></Conditions></Filter><ModuleCriteria><Module>425</Module><IsKeywordModule>True</IsKeywordModule><BuildoutRelationship>Union</BuildoutRelationship><SortFields><SortField><Field>16719</Field><SortType>Ascending</SortType></SortField><SortField><Field>23874</Field><SortType>Ascending</SortType></SortField></SortFields></ModuleCriteria></Criteria></SearchReport>"""

        body2 = \
            """<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
            <ExecuteSearch xmlns="http://archer-tech.com/webservices/">
            <sessionToken>""" + sessionToken + """</sessionToken>
            <searchOptions><![CDATA[""" + searchOptions + """]]></searchOptions>
            <pageNumber>1</pageNumber>
            </ExecuteSearch>
            </soap:Body>
            </soap:Envelope>
            """
        body2 = body2.encode("utf-8")
        response = requests.post(baseurl + '/ws/search.asmx',
                                 verify=False, data=body2, headers=userSessionHeader2)
        tree2 = ET.fromstring(response.content)
        print(tree2)
        reg = tree2.find(
            './/{http://archer-tech.com/webservices/}ExecuteSearchResult').text
        regf = reg.replace("""<?xml version="1.0" encoding="utf-16"?>""",
                           """<?xml version="1.0" encoding="utf-8"?>""")
        regf2 = regf.encode('utf-8')
        existeTicket = ET.fromstring(regf2)
        record = existeTicket.find('Record')
        if record is None:
            return None
        else:
            contentId = record.get('contentId')
            return contentId
    except:
        logging.error("Ha sucedido un error en: 'existeTicketId'")


def createJSONDetalle(msg, subject, fromm, to, tomail, cc, ccmail, obs):
    try:
        aux = []
        paravalues = []
        ccvalues = []
        auxtomail = []
        auxccmail = []
        levelId = ddcIds['levelId']
        detalle = ddcIds['cuerpo']
        asunto = ddcIds['asunto']
        de = ddcIds['de']
        para = ddcIds['para']
        ccc = ddcIds['cc']
        obss = ddcIds['obs']
        mailpara = ddcIds['paramail']
        mailcc = ddcIds['ccmail']
        paramailref = ddcIds['paramailref']
        ccmailref = ddcIds['ccmailref']
        tomail = tomail.replace(";Nova", "")

        for mail in tomail.split(";"):
            if "@" in mail:
                auxtomail.append(mail)
                contentId = contentIdDetalle(sessionToken, mail)
                if contentId is not None:
                    paravalues.append({"ContentId": contentId})

        for mail in ccmail.split(";"):
            if mail:
                if "@" in mail:
                    auxccmail.append(ccmail)
                    contentId = contentIdDetalle(sessionToken, mail)
                    if contentId is not None:
                        ccvalues.append({"ContentId": contentId})

        auxtomail = list(dict.fromkeys(auxtomail))
        auxccmail = list(dict.fromkeys(auxccmail))
        tomail = ';'.join(auxtomail)
        ccmail = ';'.join(auxccmail)
        data = {
            "Content": {
                "LevelId": levelId,
                "FieldContents": {
                    str(detalle): {
                        "Type": 1,
                        "Value": msg,
                        "FieldId": str(detalle)
                    },
                    str(asunto): {
                        "Type": 1,
                        "Value": subject,
                        "FieldId": str(asunto)
                    },
                    str(de): {
                        "Type": 1,
                        "Value": fromm,
                        "FieldId": str(de)
                    },
                    str(para): {
                        "Type": 1,
                        "Value": to,
                        "FieldId": str(para)
                    },
                    str(ccc): {
                        "Type": 1,
                        "Value": cc,
                        "FieldId": str(ccc)
                    },
                    str(mailpara): {
                        "Type": 1,
                        "Value": tomail,
                        "FieldId": str(mailpara)
                    },
                    str(mailcc): {
                        "Type": 1,
                        "Value": ccmail,
                        "FieldId": str(mailcc)
                    },
                    str(obss): {
                        "Type": 1,
                        "Value": obs,
                        "FieldId": str(obss)
                    },
                    str(paramailref): {
                        "Type": 9,
                        "Value": paravalues,
                        "FieldId": str(paramailref)
                    },
                    str(ccmailref): {
                        "Type": 9,
                        "Value": ccvalues,
                        "FieldId": str(ccmailref)
                    }
                }
            }
        }
        return data
    except:
        logging.error("Ha sucedido un error en: 'createJSONDetalle'")


def createJSONTicket(moduleId, msg, mail):
    try:
        subforms = []
        values = []
        valasunto = decode_header(msg['subject'])[0][0].decode("iso-8859-1")
        decola = []
        paracola = []
        cccola = []
        valde = mail
        valpara = mail
        valcc = mail
        valmailcc = ''
        if msg['TO']:
            valpara = ",".join(getEmails(msg['to'])['emails'])
            paracola = getEmails(msg['to'])['emails']
            tonames = valpara
            toemails = ";".join(getEmails(msg['to'])['emails'])
        if msg['CC']:
            cccola = getEmails(msg['cc'])['emails']
            valcc = ",".join(getEmails(msg['CC'])['emails'])
            valmailcc = ";".join(getEmails(msg['CC'])['emails'])
        if msg['FROM']:
            decola = getEmails(msg['from'])['emails']
            valde = ",".join(decola)
            valmailde = ";".join(decola)
        levelId = ticketsIds['levelId']
        ddc = ticketsIds['ddc']
        asunto = ticketsIds['asunto']
        cc = ticketsIds['cc']
        de = ticketsIds['de']
        para = ticketsIds['para']
        delacola = ticketsIds['delacola']
        hasHTML = False
        body = ''
        obs = ''
        attNames = []
        attBase64 = []

        data = {
            "Content": {
                "LevelId": levelId,
                "FieldContents": {
                    str(ddc): {
                        "Type": 9,
                        "Value": subforms,
                        "FieldId": str(ddc)
                    },
                    str(asunto): {
                        "Type": 1,
                        "Value": valasunto,
                        "FieldId": str(asunto)

                    },
                    str(cc): {
                        "Type": 1,
                        "Value": valmailcc,
                        "FieldId": str(cc)
                    },
                    str(de): {
                        "Type": 1,
                        "Value": valmailde,
                        "FieldId": str(de)
                    },
                    str(para): {
                        "Type": 1,
                        "Value": toemails,
                        "FieldId": str(para)
                    },
                    str(delacola): {
                        "Type": 4,
                        "Value": {
                            'ValuesListIds': [],
                            'OtherText': None
                        },
                        "FieldId": str(delacola)
                    }

                }
            }
        }

        if emailservicios in decola or emailservicios in paracola or emailservicios in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['servicios']]
        if emailsoc in decola or emailsoc in paracola or emailsoc in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['soc']]

        if emailinstalacion in decola or emailinstalacion in paracola or emailinstalacion in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['instalacion']]

        if emailmesadeayuda in decola or emailmesadeayuda in paracola or emailmesadeayuda in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['mesadeayuda']]

        if emailsoporte in decola or emailsoporte in paracola or emailsoporte in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['soporte']]

        if emailsocclaro in decola or emailsocclaro in paracola or emailsocclaro in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['socclaro']]

        return [data, subforms, values, body]
    except:
        logging.error("Ha sucedido un error en: 'createJSONTicket'")


def createJSONActualizarAdjunto(ddcId, attachments, msg):
    try:
        levelId = ddcIds['levelId']
        adjunto = ddcIds['adjunto']
        detalle = ddcIds['cuerpo']
        values = attachments

        data = {
            "Content": {
                "Id": ddcId,
                "LevelId": levelId,
                "FieldContents": {
                    str(detalle): {
                        "Type": 1,
                        "Value": msg,
                        "FieldId": str(detalle)
                    },
                    str(adjunto): {
                        "Type": 11,
                        "Value": values,
                        "FieldId": adjunto
                    }
                }
            }
        }
        return data
    except:
        logging.error("Ha sucedido un error en: 'createJSONActualizarAdjunto'")


def createJSONCrearAdjunto(baseurl, contentId, moduleId, value):
    try:
        levelId = ddcIds['levelId']
        adjunto = ddcIds['adjunto']
        registro = getContentById(baseurl, sessionToken, contentId)
        values = registro.json()[
            'RequestedObject']['FieldContents'][str(adjunto)]['Value']
        if values:
            values.append(value)
        else:
            values = [value]

        data = {
            "Content": {
                "Id": contentId,
                "LevelId": levelId,
                "FieldContents": {
                    str(adjunto): {
                        "Type": 11,
                        "Value": values,
                        "FieldId": adjunto
                    }
                }
            }
        }
        return data
    except:
        logging.error("Ha sucedido un error en: 'createJSONCreateAdjunto'")


def updateJSONTicket(baseurl, contentId, moduleId, subformId):
    try:
        levelId = ticketsIds['levelId']
        ddc = ticketsIds['ddc']
        rticket = ticketsIds['realmentececticket']
        rticketno = ticketsIds['rticketno']
        trabierto = ticketsIds['trabierto']
        trabiertosi = ticketsIds['trabiertosi']
        cticket = ticketsIds['cerrarticket']
        cticketno = ticketsIds['cticketno']
        rrticket = ticketsIds['realmentececticket']
        rrticketsi = ticketsIds['rticketsi']
        estado = ticketsIds['estado']
        ecerrado = ticketsIds['estadocerrado']

        registro = getContentById(baseurl, sessionToken, contentId)
        # print registro.json()
        subforms = registro.json(
        )['RequestedObject']['FieldContents'][str(ddc)]['Value']
        if subforms:
            subforms.append({"ContentID": subformId})
        else:
            subforms = [{"ContentID": subformId}]

        data = {
            "Content": {
                "Id": contentId,
                "LevelId": levelId,
                "FieldContents": {
                    str(ddc): {
                        "Type": 9,
                        "Value": subforms,
                        "FieldId": str(ddc)
                    },
                    str(rticket): {
                        "Type": 4,
                        "Value": {
                            'ValuesListIds': [],
                            'OtherText': None
                        },
                        "FieldId": str(rticket)
                    },
                    str(cticket): {
                        "Type": 4,
                        "Value": {
                            'ValuesListIds': [],
                            'OtherText': None
                        },
                        "FieldId": str(cticket)
                    },
                    str(trabierto): {
                        "Type": 4,
                        "Value": {
                            'ValuesListIds': [],
                            'OtherText': None
                        },
                        "FieldId": str(trabierto)
                    }

                }
            }
        }

        # print data['Content']['FieldContents'][str(rticket)]
        # for val in registro.json()['RequestedObject']['FieldContents'][str(estado)]['Value']['ValuesListIds']:
        # print val
        # print ecerrado
        # if val == ecerrado:
        #    data['Content']['FieldContents'][str(rticket)]['Value']['ValuesListIds']=[int(rticketno)]
        #    data['Content']['FieldContents'][str(cticket)]['Value']['ValuesListIds']=[int(cticketno)]
        #    data['Content']['FieldContents'][str(trabierto)]['Value']['ValuesListIds']=[int(trabiertosi)]

        # print data
        return data
    except:
        logging.error("Ha sucedido un error en: 'updateJSONTicket'")


def postContent(url, sessionToken, content):
    try:
        url = url+'/api/core/content'
        data = content
        headers = {
            'Accept': 'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Authorization': 'Archer session-id='+sessionToken,
            'Content-Type': 'application/json'
        }
        response = requests.post(url, verify=False, headers=headers, json=data)

        return response
    except:
        logging.error("Ha sucedido un error en: 'postContent'")


def putContent(url, sessionToken, content):
    try:
        url = url+'/api/core/content'
        data = content
        headers = {
            'Accept': 'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Authorization': 'Archer session-id='+sessionToken,
            'Content-Type': 'application/json'
        }
        response = requests.put(url, verify=False, headers=headers, json=data)
        return response
    except:
        logging.error("Ha sucedido un error en: 'putContent'")


def getContentById(url, sessionToken, contentId):
    try:
        print("Content ID")
        print(contentId)
        print("Content ID")
        url = url+'/api/core/content/' + str(contentId)
        headers = {
            'Accept': 'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Authorization': 'Archer session-id='+sessionToken,
            'Content-Type': 'application/json',
            'X-Http-Method-Override': 'GET'
        }
        response = requests.post(url, verify=False, headers=headers)
        return response
    except:
        logging.error("Ha sucedido un error en: 'getContentById'")


#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

# Funciones asociadas a la lectura del servidor Outlook

def deleteMailIMAP(servername, mail, password, msgId):
    try:
        imap = imaplib.IMAP4_SSL(servername, 993)
        imap.login(mail, password)
        imap.select()
        typ, data = imap.search(None, 'ALL')
        for num in data[0].split():
            typ, data = imap.fetch(num, '(BODY[HEADER.FIELDS (MESSAGE-ID)])')
            msg_str = email.message_from_string(data[0][1])
            message_id = msg_str.get('Message-ID')
            if message_id == msgId:
                imap.copy(num, "Archivo")
                imap.store(num, '+FLAGS', '\\Deleted')
                imap.expunge()
        imap.close()
        imap.logout()
    except:
        logging.error("Ha sucedido un error en: 'deleteMailIMAP'")


def createPOPConn(servername, email, password):
    try:
        pop_conn = poplib.POP3_SSL(servername)
        pop_conn.user(email)
        pop_conn.pass_(password)
        print(pop_conn)
        return pop_conn
    except:
        logging.error("Ha sucedido un error en: 'createPOPConn'")

def deleteMail(msg, servername, email, password):
    try:
        tries = 3
        for i in range(tries):
            try:
                # Crea la conexión para leer los mails
                pop_conn = createPOPConn(servername, email, password)
            except socket.error as se:
                numTry = (-i) + 3
                print("Hubo un error con la conexión al servidor POP de Outlook. Se realizaran " +
                    str(numTry) + " intentos mas para la cuenta " + email)
                logging.error("Hubo un error con la conexión al servidor POP de Outlook. Se realizaran  " +
                            str(numTry) + " intentos mas para la cuenta " + email, exc_info=True)
                if i < tries - 1:
                    continue
                else:
                    print("No se pudo conectar al servicio POP tras 3 reintentos. Contactarse con un administrador. La cuenta de email es " + email)
                    logging.error(
                        "No se pudo conectar al servicio POP tras 3 reintentos. Contactarse con un administrador. La cuenta de email es " + email + str(se))
            except poplib.error_proto as pe:
                print("Ocurrio un error con la cuenta " +
                    email+" ¿contraseña erronea?")
                print(pe)
                logging.error("Ocurrio un error con la cuenta " +
                            email+" ¿contraseña erronea?", exc_info=True)
            break
        for i in range(1, len(pop_conn.list()[1]) + 1):
            mssg = pop_conn.retr(i)
            mssg = "\n".join(mssg[1])
            mssg = parser.Parser().parsestr(mssg)
            if mssg['Message-ID'] == msg['Message-ID'] and mssg['date'] == msg['date']:
                pop_conn.dele(i)

        pop_conn.quit()
    except:
        logging.error("Ha sucedido un error en: 'deleteMail'")


def fetchMail(pop_conn, delete_after=False):
    try:
        # Get messages from server:
        logging.error("Ingresando a la funcion fetchMail")
        print("Hola")
        # for i in range(1, len(pop_conn.list()[1]) + 1):
        #    print i
        #    mssg = pop_conn.retr(i)
        #    mssg = "\n".join(mssg[1])
        #    mssg = parser.Parser().parsestr(mssg)
        #    print mssg["subject"]
        # Concat message pieces:
        messages = [pop_conn.retr(i)
                    for i in range(1, len(pop_conn.list()[1]) + 1)]
        messages = ["\n".join(mssg[1]) for mssg in messages]
        # Parse message intom an email object:
        messages = [parser.Parser().parsestr(mssg) for mssg in messages]
        # if delete_after == True:
        #    delete_messages = [pop_conn.dele(i) for i in range(1, len(pop_conn.list()[1]) + 1)]
        pop_conn.quit()
        return messages
    except:
        logging.error("Ha sucedido un error en: 'fetchMail'")

#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

# Funciones de lectura y escritura de mails en Archer


def createAttachments(part):
    try:
        value = ' '
        obs = ' '
        charset = part.get_content_charset('iso-8859-1')
        name = part.get_filename()
        if not part.get_filename():
            name = str(random.randint(1, 1000)) + "." + \
                str(part.get_content_type()).split("/")[1]
        name = name.decode(charset, 'replace')
        if '?' in name:
            name = name.split('?')[3].decode('iso-8859-1')
        data = part.get_payload(decode=False)
        attNames.append(name)
        attBase64.append(part.get_payload().replace('\n', ''))
        name = name.replace('\n', '')
        iD = postAttachment(baseurl, sessionToken, name, data)
        logging.error("Adjunto durante la creación" + str(iD))
        try:
            # Cargo el archivo adjunto en Archer
            iD = postAttachment(baseurl, sessionToken, name, data)
            logging.error("Adjunto durante la actualización" + str(iD))
            if iD['IsSuccessful']:
                value = iD['RequestedObject']['Id']
            else:
                obs += iD['ValidationMessages'][0]['ResourcedMessage']
        except err:
            print(err)
            traceback.print_exc()
            logging.error("Hubo un error al cargar el asunto: " +
                        str(err), exc_info=True)
            # print iD
        return value, obs
    except:
        logging.error("Ha sucedido un error en: 'createAttachments'")


def existeTicketAsociadoAlMensaje(msg, email):
    try:
        contentId = None
        ticketId = None
        # El mail posee numero de Ticket
        if 'Ticket#' in decode_header(msg['subject'])[0][0].decode("iso-8859-1"):
            # Valido que el numero de Ticket sea el de Novared
            # En el caso de ANSES tienen un numero de ticket, el primero dentro del Asunto del mail
            if decode_header(msg['subject'])[0][0].decode("iso-8859-1").count("Ticket#") > 1:
                contentId = decode_header(msg['subject'])[0][0].decode(
                    "iso-8859-1").split("Ticket#")[2]
                contentId = contentId[0:7]
            else:
                contentId = decode_header(msg['subject'])[0][0].decode(
                    "iso-8859-1").split("Ticket#")[1]
                # Obtengo el numero de ticket del asunto del mail
                contentId = contentId[0:7]
            # if not getContentById(baseurl, sessionToken, contentId).json()['IsSuccessful']:
            #   contentId = None
            ticketId = existeTicketId(sessionToken, contentId)
        if ticketId is None:
            return None
        else:
            return ticketId
    except:
        logging.error("Ha sucedido un error en: 'existeTicketAsociadoAlMensaje'")


def actualizarTicket(msg, idTicket, email):
    try:
        values = attBase64 = attNames = []
        subformId = obs = valmailcc = toemails = body = obs = ''
        isSuccessful = hasHTML = False
        ddc = ticketsIds["ddc"]
        valcc = tonames = email
        if msg['TO']:
            tonames = ",".join(getEmails(msg['to'])['emails'])
            toemails = ";".join(getEmails(msg['to'])['emails'])
        if msg['CC']:
            valcc = ",".join(getEmails(msg['CC'])['emails'])
            valmailcc = ";".join(getEmails(msg['CC'])['emails'])

        for part in msg.walk():
            if part.get_content_type() in allowed_mimetypes:  # Crea los adjuntos
                value, obs = createAttachments(part)
                values.append(value)

            if part.get_content_type() == "text/plain":
                if part.get_payload().endswith('='):
                    body = part.get_payload().decode('base64')

                if part.get_content_type() == "text/html":
                    hasHTML = True
                    charset = part.get_content_charset('iso-8859-1')
                    body = part.get_payload(decode=True)
                    body = body.decode(charset, 'replace')

                jsonDetalle = createJSONDetalle(body, decode_header(msg['subject'])[0][0].decode(
                    "iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
                jsonDetalleRta = postContent(
                    baseurl, sessionToken, jsonDetalle).json()
                idDetalle = jsonDetalleRta['RequestedObject']['Id']
                jsonTicket = updateJSONTicket(
                    baseurl, idTicket, moduleId, idDetalle)
                req = putContent(baseurl, sessionToken, jsonTicket)
                isSuccessfulUpdate = req.json()['IsSuccessful']

        for name in attNames:
            if name in body:
                i = attNames.index(name)
                cid = body.find('cid:'+name)+len('cid:'+name)
                st = 'cid:'+name+body[cid:cid+18]
                body = body.replace(st, 'http://novatickets/' +
                                    name.split('.')[1]+attBase64[i])

        # print body
        subformJSON = createJSONActualizarAdjunto(idDetalle, values, body)
        req = putContent(baseurl, sessionToken, subformJSON)
        isSuccessful = req.json()['IsSuccessful']
        return isSuccessful
    except:
        logging.error("Ha sucedido un error en: 'actualizarTicket'")


def crearTicket(msg, mail):
    try:
        subforms = []
        values = []
        valasunto = decode_header(msg['subject'])[0][0].decode("iso-8859-1")
        decola = []
        paracola = []
        cccola = []
        valde = mail
        valpara = mail
        valcc = mail
        valmailcc = ''
        if msg['TO']:
            paracola = getEmails(msg['to'])['emails']
            for mail in paracola:
                if '@' not in mail:
                    paracola.remove(mail)
            valpara = ",".join(paracola)
            tonames = valpara
            toemails = ";".join(paracola)
        if msg['CC']:
            cccola = getEmails(msg['cc'])['emails']
            for mail in cccola:
                if '@' not in mail:
                    cccola.remove(mail)
            valcc = ",".join(cccola)
            valmailcc = ";".join(cccola)
        if msg['FROM']:
            decola = getEmails(msg['from'])['emails']
            for mail in decola:
                if '@' not in mail:
                    decola.remove(mail)
            valde = ",".join(decola)
            valmailde = ";".join(decola)
        levelId = ticketsIds['levelId']
        asunto = ticketsIds['asunto']
        cc = ticketsIds['cc']
        de = ticketsIds['de']
        para = ticketsIds['para']
        delacola = ticketsIds['delacola']
        clienteTicket = ticketsIds['clienteTicket']
        arrayClienteAux = valmailde.split("@")
        emailCliente = arrayClienteAux[1].split(".")[0]
        print(emailCliente)
        emailId = [{"ContentID": idCliente(
            sessionToken, clienteTicket, emailCliente)}]

        data = {
            "Content": {
                "LevelId": levelId,
                "FieldContents": {
                    str(asunto): {
                        "Type": 1,
                        "Value": valasunto,
                        "FieldId": str(asunto)

                    },
                    str(cc): {
                        "Type": 1,
                        "Value": valmailcc,
                        "FieldId": str(cc)
                    },
                    str(de): {
                        "Type": 1,
                        "Value": valmailde,
                        "FieldId": str(de)
                    },
                    str(para): {
                        "Type": 1,
                        "Value": toemails,
                        "FieldId": str(para)
                    },
                    str(delacola): {
                        "Type": 4,
                        "Value": {
                            'ValuesListIds': [],
                            'OtherText': None
                        },
                        "FieldId": str(delacola)
                    },
                    str(clienteTicket): {
                        "Type": 9,
                        "Value": emailId,
                        "FieldId": str(clienteTicket)
                    }
                }
            }
        }
        print(data)
        if emailservicios in decola or emailservicios in paracola or emailservicios in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['servicios']]
        if emailsoc in decola or emailsoc in paracola or emailsoc in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['soc']]

        if emailinstalacion in decola or emailinstalacion in paracola or emailinstalacion in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['instalacion']]

        if emailmesadeayuda in decola or emailmesadeayuda in paracola or emailmesadeayuda in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['mesadeayuda']]

        if emailsoporte in decola or emailsoporte in paracola or emailsoporte in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['soporte']]

        if emailsocclaro in decola or emailsocclaro in paracola or emailsocclaro in cccola:
            data['Content']['FieldContents'][str(delacola)]['Value']['ValuesListIds'] = [
                ticketsIds['socclaro']]

        jsonTicket = postContent(baseurl, sessionToken, data).json()
        idTicket = jsonTicket['RequestedObject']['Id']
        isSuccessful = jsonTicket['IsSuccessful']
        if isSuccessful:
            return isSuccessful, idTicket
        else:
            logging.error(
                "Hubo un error al crear el Ticket Padre. Se copia el JSON de Archer a continuación:")
            logging.error(str(jsonTicket))
            return isSuccessful
    except:
        logging.error("Ha sucedido un error en: 'crearTicket'")


def main(servername, email, password):
    #allowed_mimetypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","image/png","image/jpg", "image/jpeg"]
    messages = ""
    tries = 3
    for i in range(tries):
        try:
            # Crea la conexión para leer los mails
            pop_conn = createPOPConn(servername, email, password)
            print(email)
            # Trae los mails
            messages = fetchMail(pop_conn)
        except socket.error as se:
            numTry = (-i) + 3
            print("Hubo un error con la conexión al servidor POP de Outlook. Se realizaran  " +
                  str(numTry) + " intentos mas para la cuenta " + email)
            logging.error("Hubo un error con la conexión al servidor POP de Outlook. Se realizaran  " +
                          str(numTry) + " intentos mas para la cuenta " + email, exc_info=True)
            if i < tries - 1:
                continue
            else:
                print("No se pudo conectar al servicio POP tras 3 reintentos. Contactarse con un administrador. La cuenta de email es " + email)
                logging.error(
                    "No se pudo conectar al servicio POP tras 3 reintentos. Contactarse con un administrador. La cuenta de email es " + email + str(se))
        except poplib.error_proto as pe:
            print("Ocurrio un error con la cuenta " +
                  email+" ¿contraseña erronea?")
            print(pe)
            logging.error("Ocurrio un error con la cuenta " +
                          email+" ¿contraseña erronea?", exc_info=True)
        break

    #sessionToken = apiCall(baseurl+'/api/core/security/login', getHeaders(), data). json()['RequestedObject']['SessionToken']
    #headers = getHeaders(sessionToken)

    for msg in messages:
        # print "ID"
        #msgId = msg['Message-ID']
        # print "ID"
        isSuccessfulTicket = isSuccessfulDetalle = False
        try:
            asuntoLogging = decode_header(msg['subject'])[
                0][0].decode("iso-8859-1")
            logging.error("Asunto del mail: " + asuntoLogging)
            if email == emailsocclaro:

                asunto = decode_header(msg['subject'])[
                    0][0].decode("iso-8859-1")
                asunto2 = asunto.replace("RE:", "")
                asunto3 = asunto2.replace("RV:", "")
                asunto4 = asunto3.replace("FWD:", "")
                asuntof = asunto4.lstrip()
                idTicket = verificarAsunto(sessionToken, asuntof)
                print("Entró")
                if idTicket == None:
                    isSuccessfulTicket, idTicket = crearTicket(msg, email)
                    if isSuccessfulTicket:
                        isSuccessfulDetalle = updateTicket(
                            idTicket, configParser.get('env', 'moduleIdTickets'), msg, email)
                else:
                    isSuccessfulTicket = True
                    isSuccessfulDetalle = updateTicket(
                        idTicket, configParser.get('env', 'moduleIdTickets'), msg, email)

                if isSuccessfulTicket and isSuccessfulDetalle:
                    logging.error(
                        "Se mueve el mail de carpeta y se elimina del")
                    deleteMail(msg, servername, email, password)
                    #deleteMailIMAP(servername, email, password, msgId)
                else:
                    logging.error(
                        "Hubo un error al crear el Ticket:" + str(asuntoLogging))
            if email != emailsocclaro:
                idTicket = existeTicketAsociadoAlMensaje(msg, email)
                if idTicket == None:
                    isSuccessfulTicket, idTicket = crearTicket(msg, email)
                    if isSuccessfulTicket:
                        isSuccessfulDetalle = updateTicket(
                            idTicket, configParser.get('env', 'moduleIdTickets'), msg, email)
                else:
                    isSuccessfulTicket = True
                    isSuccessfulDetalle = updateTicket(
                        idTicket, configParser.get('env', 'moduleIdTickets'), msg, email)

                # Se elimina el mail.
                if isSuccessfulTicket and isSuccessfulDetalle:
                    logging.error(
                        "Se elimina el mail tras una ejecución exitosa")
                    deleteMail(msg, servername, email, password)
                else:
                    logging.error(
                        "Hubo un error al crear el Ticket:" + str(asuntoLogging))
        except err:
            print("Hubo un error:")
            print(err)
            traceback.print_exc()
            logging.error("Hubo un error: " + str(err), exc_info=True)

    # return isSuccessfulTicket, isSuccesfulDetalle
    return True


def updateTicket(contentId, moduleId, msg, mail):
    try:
        values = []
        subformId = ''
        isSuccessful = False
        ddc = ticketsIds["ddc"]
        obs = ''
        valcc = mail
        valmailcc = ''
        tonames = mail
        toemails = ''
        if msg['TO']:
            tonames = ",".join(getEmails(msg['to'])['emails'])
            toemails = ";".join(getEmails(msg['to'])['emails'])
        if msg['CC']:
            valcc = ",".join(getEmails(msg['CC'])['emails'])
            valmailcc = ";".join(getEmails(msg['CC'])['emails'])

        hasHTML = False
        body = ''
        obs = ''
        attNames = []
        attBase64 = []

        for part in msg.walk():
            # print part.get_content_type()
            if part.get_content_type() in allowed_mimetypes:
                # print part
                name = part.get_filename()
                if name == None:
                    extension = part.get_content_type().split('/')[1]
                    name = "Adjunto" + "." + extension
                    # print name
                if '?' in name:
                    name = name.split('?')[3]
                data = part.get_payload(decode=False)
                attNames.append(name)
                attBase64.append(part.get_payload().replace('\n', ''))
                name = name.replace('\n', '')
                try:
                    # Cargo el archivo adjunto en Archer
                    iD = postAttachment(baseurl, sessionToken, name, data)
                    logging.error("Adjunto durante la actualización" + str(iD))
                    if iD['IsSuccessful']:
                        value = iD['RequestedObject']['Id']
                        # En caso que no funcionen los adjuntos, comentar esta línea
                        values.append(value)
                    else:
                        obs += iD['ValidationMessages'][0]['ResourcedMessage']
                except err:
                    print(err)
                    traceback.print_exc()
                    logging.error(
                        "Hubo un error al cargar el asunto: " + str(err), exc_info=True)
                    # print iD

            if part.get_content_type() not in allowed_mimetypes and part.get_content_type() not in ["text/html", "text/plain", "multipart/alternative", "multipart/mixed"]:
                obs += 'Se encontro un adjunto ' + \
                    part.get_content_type() + ' que no se pudo cargar.'

            if part.get_content_type() == "text/plain":
                # print part.get_payload()
                if part.get_payload().endswith('='):
                    texto = part.get_payload().decode('base64')
                    # subject
                    #msg, subject, fromm, to, tomail, cc, ccmail, obs
                    json = createJSONDetalle(texto, decode_header(msg['subject'])[0][0].decode(
                        "iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
                    subform = postContent(baseurl, sessionToken, json).json()
                    subformId = subform['RequestedObject']['Id']
                    json2 = updateJSONTicket(
                        baseurl, contentId, moduleId, subformId)
                    req = putContent(baseurl, sessionToken, json2)
                    isSuccessful = req.json()['IsSuccessful']

            if part.get_content_type() == "text/html":
                hasHTML = True
                charset = part.get_content_charset('iso-8859-1')
                body = part.get_payload(decode=True)
                texto = body.decode(charset, 'replace')
                body = texto
                json = createJSONDetalle(texto, decode_header(msg['subject'])[0][0].decode(
                    "iso-8859-1"), ",".join(getEmails(msg['from'])['emails']), tonames, toemails, valcc, valmailcc, obs)
                subform = postContent(baseurl, sessionToken, json).json()
                subformId = subform['RequestedObject']['Id']
                json2 = updateJSONTicket(baseurl, contentId, moduleId, subformId)
                req = putContent(baseurl, sessionToken, json2)
                isSuccessful = req.json()['IsSuccessful']

        if not hasHTML:
            json = createJSONDetalle(body, decode_header(msg['subject'])[0][0].decode(
                "iso-8859-1"), ";".join(getEmails(msg['from'])['emails']), tonames, toemails, valmailcc, valmailcc, obs)
            subform = postContent(baseurl, sessionToken, json).json()
            subformId = subform['RequestedObject']['Id']
            json2 = updateJSONTicket(baseurl, contentId, moduleId, subformId)
            req = putContent(baseurl, sessionToken, json2)
            isSuccessful = req.json()['IsSuccessful']

        for name in attNames:
            if name in body:
                i = attNames.index(name)
                cid = body.find('cid:'+name)+len('cid:'+name)
                st = 'cid:'+name+body[cid:cid+18]
                body = body.replace(st, 'http://novatickets/' +
                                    name.split('.')[1]+attBase64[i])
        # print body
        subformJSON = createJSONActualizarAdjunto(subformId, values, body)
        req = putContent(baseurl, sessionToken, subformJSON)
        isSuccessful = req.json()['IsSuccessful']
        if isSuccessful:
            return isSuccessful
        else:
            logging.error(
                "Hubo un error al crear el Detalle o al actualizar el ticket. Se copia el JSON de Archer a continuación:")
            logging.error(str(req.json()))
            return isSuccessful
    except:
        logging.error("Ha sucedido un error en: 'updateTicket'")


def getEmails(emailsstring):
    try:
        if '"' in emailsstring:
            emailsstring = ''.join(emailsstring.split('"'))
        if "\n" in emailsstring:
            emailsstring = ''.join(emailsstring.split('\n'))
        if "\t" in emailsstring:
            emailsstring = ''.join(emailsstring.split('\t'))
        emails = {
            "names": [],
            "emails": []
        }

        if ',' in emailsstring:
            emailsstring = emailsstring.split(",")
            for emailstring in emailsstring:
                emailaux = emailstring.split("<")
                emailaux[0] = emailaux[0].strip()
                if len(emailaux) == 1:
                    emails['names'].append(emailaux[0])
                    emails['emails'].append(emailaux[0])
                else:
                    emails['names'].append(emailaux[0])
                    emails['emails'].append(emailaux[1].split(">")[0])
        else:
            emailaux = emailsstring.split("<")
            emailaux[0] = emailaux[0].strip()
            if len(emailaux) == 1:
                emails['names'].append(emailaux[0])
                emails['emails'].append(emailaux[0])
            else:
                emails['names'].append(emailaux[0])
                emails['emails'].append(emailaux[1].split(">")[0])
        return emails
    except:
        logging.error("Ha sucedido un error en: 'getEmails'")


def getHeaders(sessionToken=''):
    try:
        headers = {
            'Accept': 'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Content-Type': 'application/json'
        }

        if sessionToken != '':
            headers['Authorization'] = 'Archer session-id='+sessionToken

        return headers
    except:
        logging.error("Ha sucedido un error en: 'getHeaders'")



def apiCall(url, headers, content):
    try:
        data = content
        headers = headers
        response = requests.post(url, verify=False, headers=headers, json=data)
        return response
    except:
        logging.error("Ha sucedido un error en: 'apiCall'")



#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################

today = date.today()
date = today.strftime("%d%m%Y")
baseurl = configParser.get('env', 'archerurl')
allowed_mimetypes = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword', 'application/vnd.ms-excel',
                     'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/pdf', 'text/csv', 'image/png', 'image/jpeg', 'image/gif', 'application/x-zip-compressed', 'text/plain']
servername = configParser.get('env', 'servername')
servernameSOC = configParser.get('env', 'servernameSOC')
emailservicios = configParser.get('env', 'emailservicios')
emailsoc = configParser.get('env', 'emailsoc')
emailinstalacion = configParser.get('env', 'emailinstalacion')
emailmesadeayuda = configParser.get('env', 'emailmesadeayuda')
emailsoporte = configParser.get('env', 'emailsoporte')
emailsocclaro = configParser.get('env', 'emailsocclaro')
emailpassword = configParser.get('env', 'emailpassword')
passservicios = configParser.get('env', 'passservicios')
passinstalacion = configParser.get('env', 'passinstalacion')
passsoc = configParser.get('env', 'passsoc')
passmesadeayuda = configParser.get('env', 'passmesadeayuda')
passsoporte = configParser.get('env', 'passsoporte')
passsocclaro = configParser.get('env', 'passsocclaro')
fileAddr = configParser.get('env', 'log')
filename = fileAddr + "Log-" + date + ".txt"


data = {
    "InstanceName": configParser.get('env', 'archerinstancename'),
    "Username": configParser.get('env', 'archerusername'),
    "UserDomain": "",
    "Password": configParser.get('env', 'archerpassword')
}


logging.basicConfig(filename=filename, format='%(asctime)s - %(message)s')
logging.error(
    "###############################################################################################################")
logging.error("COMIENZA EJECUCION DEL SCRIPT")
logging.error(
    "###############################################################################################################")

try:
    sessionToken = apiCall(baseurl+'/api/core/security/login',
                        getHeaders(), data). json()['RequestedObject']['SessionToken']
    print(sessionToken)
except:
        logging.error("Ha sucedido un error en: 'Api call sessiontoken'")



headers = getHeaders(sessionToken)


def tareasFinales():
    guidDFContactos = configParser.get('env', 'GuidDFContactos')
    guidDFTicketMerge = configParser.get('env', 'GuidDFTicketMerge')
    dfContent = {"DataFeedGuid": guidDFContactos,
                 "IsReferenceFeedsIncluded": False}
    apiCall(baseurl+'/api/core/datafeed/execution', headers, dfContent)
    dfContent = {"DataFeedGuid": guidDFTicketMerge,
                 "IsReferenceFeedsIncluded": True}
    apiCall(baseurl+'/api/core/datafeed/execution', headers, dfContent)


main(servername, emailsoporte, passsoporte)
main(servername, emailservicios, passservicios)
main(servername, emailsoc, passsoc)
main(servername, emailinstalacion, passinstalacion)
main(servername, emailmesadeayuda, passmesadeayuda)
main(servernameSOC, emailsocclaro, passsocclaro)
tareasFinales()


# Si todo esta OK se ejecutan los datafeeds de mapeo de contactos y merge de ticket.

# print getContentById(baseurl, sessionToken, 240420).json()
#raw_input("Press Enter to continue...")
