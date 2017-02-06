#/usr/bin/env python

#to do 
# add menu for externals, etc
# protect general sheet
# add AOB
# beautify general sheet (merge columns, make it fill the full width, wrapping)

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
import datetime
from  dateutil import relativedelta

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

from subprocess import Popen,PIPE
import json

def callback(request_id, response, exception):
    if exception:
        # Handle error
        print exception
    else:
        print "Permission Id: %s" % response.get('id')


def getReleases():
    jsons=[]
    for page in range(1,100):
        comm='curl -s "https://api.github.com/repos/cms-sw/cmssw/releases?per_page=100&page='+str(page)+'"'
        p=Popen(comm,stdout=PIPE,shell=True)
        pipe=p.stdout.read()

        jsondict=json.loads(pipe)
        print len(jsondict)
        if len(jsondict)>0:
            jsons.append(json.loads(pipe))
            if len(jsondict)!= 100:
                break
        else:
            break

    alljs=[]
    for j in jsons:
        print len(j)
        for ji in j:
            alljs.append(ji)
    return alljs

import datetime
def getNewReleases(updateJson):
    retVal=[]

    if updateJson:
        alljs=getReleases()

        with open('rels.txt','w') as outfile:
            json.dump(alljs,outfile)
    else:
        with open('rels.txt') as infile:
            alljs=json.load(infile)


    for l in open('oldOrp.txt'):
        sheet_id=l.split()[1]
        version=l.split()[0]
        orpDate=l.split()[2]
        break


    year=2000+int(orpDate[3:5])
    month=int(orpDate[5:7])
    da=int(orpDate[7:9])

    lastOrp=datetime.datetime(year,month,da)

    for j in alljs:
        rDate=datetime.datetime.strptime(j['published_at'][0:10],'%Y-%m-%d')
        if rDate>=lastOrp:
            rName='=HYPERLINK(https://github.com/cms-sw/cmssw/releases/'+j['tag_name']+',"'+j['tag_name']+')'
            retVal.append([ rName,str(rDate).split()[0]])

    return retVal

def getIssues(ms):
    jsons=[]
    for page in range(1,100):
        comm='curl -s "https://api.github.com/repos/cms-sw/cmssw/issues?state=open&milestone='+str(ms)+'&per_page=100&page='+str(page)+'"'
        p=Popen(comm,stdout=PIPE,shell=True)
        pipe=p.stdout.read()
        jsondict=json.loads(pipe)
        if len(jsondict)>0:
            jsons.append(json.loads(pipe))
            if len(jsondict)!= 100:
                break
        else:
            break

    alljs=[]
    for j in jsons:
        for ji in j:
            alljs.append(ji)
    return alljs


def getPR(ex):
    return ex['pull_request']['url'].split('/')[-1]

def getPendingSigs(ex):
    pendingSigs=[]
    for i in ex['labels']:
        if 'pending' in i['name'] and 'orp-pending' not in i['name'] and 'pending-sig' not in i['name'] and 'tests-pending' not in i['name'] and 'comparison-pending' not in i['name']:
            pendingSigs.append(i['name'].split('-')[0])
        if 'hold' in i['name']:
            pendingSigs.append('hold')
    return pendingSigs

def getApprovedSigs(ex):
    approvedSigs=[]
    for i in ex['labels']:
        if 'approved' in i['name'] and 'tests-approved' not in i['name']:
            approvedSigs.append(i['name'].split('-')[0])
    return approvedSigs

def getTestsPassed(ex):
    for i in ex['labels']:
        if 'tests-approved' in i['name']: return 'yes'
    return 'no'

def getCreation(ex):
    return ex['created_at'][0:10]

def getTitle(ex):
    return ex['title']



# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Other client 1 quick start'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials

def getOldOrp(service, milestones):
    for l in open('oldOrp.txt'):
        sheet_id=l.split()[1]
        version=l.split()[0]
        orpDate=l.split()[2]
        break

    print version, sheet_id, orpDate
    
    prCol=0
    commentsCol=6
    if int(version[1:])<=0: commentsCol=5    
    print commentsCol
    
    retVal={}
    
    for m in milestones:
        retVal[m]={}
        rangeName=m
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=rangeName).execute()
        values=result.get('values',[])
        for row in values:
            if len(row)>commentsCol:
                comm=row[commentsCol]
                
                if ':' not in comm or "ORP" not in comm.split(':')[0][0:3]:
                    retVal[m][row[prCol]]=orpDate+': '+comm
                else:
                    retVal[m][row[prCol]]=comm
    return retVal



class SheetFormatter:
    def __init__(self,sheetId,genSheet):
        self.lines=[]
        self.formats=[]
        self.sheetId=sheetId
        self.genSheet=genSheet

    def addBold(self,row,fontsize):
        self.formats.append({ "repeatCell" : { "range" : { "sheetId": self.genSheet, "startRowIndex":row, "endRowIndex":row+1 },
                                               "cell" : {"userEnteredFormat" : { "textFormat" : {"fontSize": fontsize, "bold": "true" } } },
                                               "fields": "userEnteredFormat(textFormat)" }
                              })

    def addColor(self,row,colorTriplet):
        self.formats.append({ "repeatCell" : { "range" : { "sheetId": self.genSheet, "startRowIndex":row, "endRowIndex":row+1 },
                                               "cell" : {"userEnteredFormat" : { "backgroundColor" : {"red": colorTriplet[0],"green":colorTriplet[1], "blue":colorTriplet[2] } } },
                                               "fields": "userEnteredFormat(backgroundColor)" }
                              })

    def addMerge(self,row,firstCellToMerge):
        self.formats.append({ "mergeCells" : { "range" : { "sheetId": self.genSheet, "startRowIndex":row, "endRowIndex":row+1, "startColumnIndex":firstCellToMerge-1,"endColumnIndex":4 },
                                               "mergeType" : "MERGE_ALL" }
                              })

    def addLine(self,line,bold=False,fontSize='d'):
        self.lines.append(line)
        if bold or fontSize!='d':
            self.addBold(len(self.lines)-1,fontSize)
        if len(line)!=4: #merge last N
            self.addMerge(len(self.lines)-1,len(line))

    def addColoredLine(self,colorTriplet):
        self.lines.append([''])
        self.addColor(len(self.lines)-1,colorTriplet)

def  getGeneralInputs(milestones,sheetId,genSheet,updateJson):
    genSheet=SheetFormatter(sheetId,genSheet)

    genSheet.addLine(['Welcome to the CMSSW release meeting'],bold=True,fontSize=20)
    genSheet.addLine( [''])
    genSheet.addLine( [''])
    genSheet.addLine(['Connection information'],bold=True,fontSize=14)
    genSheet.addLine(['=HYPERLINK("http://vidyoportal.cern.ch/flex.html?roomdirect.html&key=uFVZiCsN6DIb","(Vidyo: Weekly_Offline_Meetings room, Extension: 9226777)")'],bold=False,fontSize=10)
    genSheet.addLine( [''])
    genSheet.addLine( [''])

    genSheet.addLine( ['Releases made since the last ORP'],bold=True,fontSize=14 )
    genSheet.addLine( ['Release','Date','Purpose'],bold=True,fontSize=12 )
    
    rels=getNewReleases(updateJson)
    for r in rels:
        genSheet.lines.append([r[0],r[1],'To be filled in'])
    genSheet.addLine( [''])

    for r in milestones:
        genSheet.addColoredLine([0.,0.,1.])
        genSheet.addLine( [''])
        genSheet.addLine([r],bold=True,fontSize=14 )
        genSheet.addLine(['Pending issues'],bold=True,fontSize=14 )
        genSheet.addLine( [''])
        genSheet.addLine(['External requests'],bold=True,fontSize=14 )
        genSheet.addLine( [''])
        genSheet.addLine( ['Issues to raise'],bold=True,fontSize=14 )
        genSheet.addLine( [''])
        genSheet.addLine(['Use this pulldown for your requests'],bold=True,fontSize=14 )
        genSheet.addLine( [''])

    genSheet.addColoredLine([0.,0.,1.])
    genSheet.AddLine(['AOB'],bold=True,fontSize=14)
    genSheet.AddLine([''])

    return genSheet

def deleteDefaultSheet(service,sheet_id):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheetsT = sheet_metadata.get('sheets', '')
    title = sheetsT[0].get("properties", {}).get("title", "Sheet1")
    tmp_sheet_id = sheetsT[0].get("properties", {}).get("sheetId", 0)
    reqArr=[ { "deleteSheet" : { "sheetId" : tmp_sheet_id } } ]
    data={'requests': reqArr}
    result = service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body=data).execute()

def defineGeneralSheet(service,sheet_id,sheetGeneral,milestoneList,updateJson):
    genSheet=getGeneralInputs(milestoneList,sheet_id,sheetGeneral,updateJson)
    data={
        "range":"General",
        "majorDimension":"ROWS",
        "values":genSheet.lines
        }

    result = service.spreadsheets().values().update(spreadsheetId=sheet_id,range="General",valueInputOption="USER_ENTERED",body=data).execute()
    data={'requests': genSheet.formats}
    result = service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body=data).execute()

def beautifyMilestoneSheets(service,sheet_id,sheets,sheetGeneral):
    reqArr=[]
    ps=[50,75,200,100,100,50,400]
    psGen=[150,150,150]

    for i in sheets:
        for j in range(0,len(ps)):
            reqArr.append( { "updateDimensionProperties": {"range" : {"sheetId":sheets[i],"dimension":"COLUMNS","startIndex":j,
                                                                      "endIndex":j+1},
                                                           "properties" : { "pixelSize" : ps[j] },
                                                           "fields" : "pixelSize"
                                                           }
                             }
                           )

    for j in range(0,len(psGen)):
        reqArr.append( { "updateDimensionProperties": {"range" : {"sheetId":sheetGeneral,"dimension":"COLUMNS","startIndex":j,
                                                                  "endIndex":j+1},
                                                       "properties" : { "pixelSize" : psGen[j] },
                                                       "fields" : "pixelSize"
                                                       }
                         }
                       )



#    print reqArr
    data={'requests': reqArr}
    result = service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body=data).execute()

def initGoogle():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    serviceDrive= discovery.build('drive','v3',http=http)
    return service,serviceDrive

def getAllIssues(milestoneList,milestones,updateJson):
    alljs={}
    for ms in milestoneList:

        if updateJson:
            jsms=getIssues(milestones[ms])

            with open('data_'+ms+'.txt','w') as outfile:
                json.dump(jsms,outfile)
        else:
            with open('data_'+ms+'.txt') as infile:
                jsms=json.load(infile)
        alljs[ms]=jsms

    for ms in alljs:
        jsms=alljs[ms]
        print 'Total number of open issues in',ms,':',len(jsms)
    return alljs

def makeORPDoc(service,ext):
    nextORP=datetime.date.today() + relativedelta.relativedelta(weeks=0, weekday=1)
    orpStr="ORP%02d%02d%02d" %(nextORP.year-2000,nextORP.month,nextORP.day)
    orpStr=orpStr+ext
    data={ 'properties': { 'title': orpStr}
         }
    result = service.spreadsheets().create(body=data).execute()
#    print result
    return result['spreadsheetId'],orpStr

def makeSheets(service,sheet_id,milestoneList,alljs):
    reqArr=[]
    #add a general tab
    reqArr.append( {"addSheet": {"properties": {"title":"General", "gridProperties": { "rowCount":50,"columnCount":4} } } } )

    #add a tab per release
    for ms in milestoneList:
        nRow=len(alljs[ms])+1
        reqArr.append( { "addSheet": { "properties": {"title":ms, "gridProperties": { "rowCount":nRow,"columnCount":7} } } } )
    data={'requests': reqArr}
    result = service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body=data).execute()

    sheets={}
    for i,r in enumerate(result['replies']):
        if i==0: 
            sheetGeneral=r['addSheet']['properties']['sheetId']
        else:
            ms=milestoneList[i-1]
            sheets[ms]=r['addSheet']['properties']['sheetId']
    return sheets,sheetGeneral


def fillMilestones(service,sheet_id,milestoneList,alljs,oldORPInfo):

    #fill data for the week!
    for ms in milestoneList:
        inputs=[]
        inputs.append( ['PR#','ReqDate','Title','Approved Sigs','Pending Sigs','Tests ok?','Requests/Comments'] )

        for ex in alljs[ms]:
            prnum=getPR(ex)
            tLink='=HYPERLINK("http://www.github.com/cms-sw/cmssw/pull/'+prnum+'","'+prnum+'")'
            comments=''
            if prnum in oldORPInfo[ms]: comments=oldORPInfo[ms][prnum]
            inputs.append([tLink,getCreation(ex),getTitle(ex),'.'.join(getApprovedSigs(ex)),','.join(getPendingSigs(ex)),getTestsPassed(ex),comments])
            
        data={
            "range":ms,
            "majorDimension":"ROWS",
            "values":inputs
        }
        result = service.spreadsheets().values().update(spreadsheetId=sheet_id,range=ms,valueInputOption="USER_ENTERED",body=data).execute()

def protectSheets(service,sheet_id,milestoneList,sheets,alljs):
    #now protect
    reqArr=[]
    for i,ms in enumerate(milestoneList):
        nRow=len(alljs[ms])+1
        reqArr.append( { "addProtectedRange": { "protectedRange": {  "range" : {"sheetId":sheets[ms]},
                                                                     "unprotectedRanges" : [ { "sheetId":sheets[ms],
                                                                                                "startColumnIndex": 6,
                                                                                                "endColumnIndex": 7 } ],
                                                                     "description" : "do not edit",
                                                                     "requestingUserCanEdit" : False,
                                                                     "warningOnly" : False,
                                                                     "editors" : { "users" : [], "groups": [], "domainUsersCanEdit" : False }
                                                } } } )


        for j in range(0,nRow):
            reqArr.append( { "updateCells" : { #'start' : { 'sheetId' : sheets[ms] },
                                            'range' : { 'sheetId' : sheets[ms],
                                                        'startRowIndex' : j,
                                                        'startColumnIndex' : 0,
                                                        'endRowIndex' : j+1,
                                                        'endColumnIndex' : 7
                                                        },
                                            'rows' : [ { 'values' : 7*[ { 'userEnteredFormat' : { 'wrapStrategy' : 'WRAP' 
                                                                                                }
                                                                        } ]
                                                            } ],
                                            'fields' : 'userEnteredFormat.wrapStrategy'
                                            }
                                            }
                                            )
                                                      
    data={'requests': reqArr}
    result = service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body=data).execute()

def printResults(sheet_id,milestoneList,sheets):    
    print 'https://docs.google.com/spreadsheets/d/'+sheet_id
    for ms in milestoneList:
        print "   * Set SHEETSLINK_"+ms,"= [[https://docs.google.com/spreadsheets/d/"+sheet_id+"#gid="+str(sheets[ms])+"]["+ms+" PRs]]"


def updateOrpFile(sheet_id,orpStr):
    f=open('oldOrp.txt','w')
    f.write('v1 '+str(sheet_id)+' '+orpStr+'\n')
    f.close()

def setPermissions(serviceDrive,sheet_id):
    user_permission = { 'role' : 'writer',
                        'type' : 'anyone' }
    batch= serviceDrive.new_batch_http_request(callback=callback)
    batch.add(serviceDrive.permissions().create(fileId=sheet_id, body=user_permission,fields='id'))
    batch.execute()


def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """

    milestoneList=['CMSSW_5_3_X',
                'CMSSW_7_1_X',
                'CMSSW_7_5_X',
                'CMSSW_7_6_X',
                'CMSSW_8_0_X',
                'CMSSW_8_1_X',
                'CMSSW_9_0_X']

    milestones={'CMSSW_5_3_X':20,
                'CMSSW_7_1_X':47,
                'CMSSW_7_5_X':51,
                'CMSSW_7_6_X':55,
                'CMSSW_8_0_X':57,
                'CMSSW_8_1_X':59,
                'CMSSW_9_0_X':64}

    isTest=True
    updateJson=(not isTest) and True
    extStr=''
    if isTest: extStr='dev'

    service,serviceDrive=initGoogle()
    oldORPInfo=getOldOrp(service,milestones)
    alljs=getAllIssues(milestoneList,milestones,updateJson)
    sheet_id,orpStr=makeORPDoc(service,extStr)
    sheets,sheetGeneral=makeSheets(service,sheet_id,milestoneList,alljs)
    beautifyMilestoneSheets(service,sheet_id,sheets,sheetGeneral)
    deleteDefaultSheet(service,sheet_id)
    defineGeneralSheet(service,sheet_id,sheetGeneral,milestoneList,updateJson)
    fillMilestones(service,sheet_id,milestoneList,alljs,oldORPInfo)
    protectSheets(service,sheet_id,milestoneList,sheets,alljs)
    printResults(sheet_id,milestoneList,sheets)
    if not isTest: updateOrpFile(sheet_id,orpStr)
    setPermissions(serviceDrive,sheet_id)

if __name__ == '__main__':
    main()
#!/usr/bin/env python


