import xlrd
import os
from os import path
import json
import csv
from csv import reader
import random
import requests
from datetime import datetime

PARTNERS = []
FIXES = {}
QUIZ_ATTEMPTS = 0
class Config:
    Name = ""
    Url = ""
    ApiUrl = ""
    ApiKey = ""
    PartnerProfileTypeId = ""
    QuizData = {}
    TrainingDataFile = ""
    FixesFile = ""
    QuizDirectory = ""
    OutputDirectory = ""
 
CONFIG = Config()

def welcome():
    print("\n\n")
    print("##################################################################")

def writeLine(line, new = False):
    global CONFIG
    scriptDir = os.path.dirname(__file__) #<-- absolute dir the script is in
    filename = CONFIG.TrainingDataFile
    filePath = CONFIG.OutputDirectory + "/" + filename
    absFilePath = os.path.join(scriptDir, filePath)
    if os.path.exists(absFilePath):
        append_write = 'a' # append if already exists
    else:
        append_write = 'w' # make a new file if not


    if new:
        append_write = 'w'
        combinedFile = open(absFilePath,append_write)
        combinedFile.truncate(0)
        combinedFile.write('attempt_id,quiz_name,name,email,org,start_date,complete_date,total_score\n')
        combinedFile.close()
    else:
        combinedFile = open(absFilePath, append_write)
        combinedFile.write(line + "\n")

def createTempFileName(suffix):
    tempDataFile = "temp-"
    for i in range(0,10):
        tempDataFile += str(random.randint(0,9))
    tempDataFile += suffix
    return tempDataFile

def displayText(text, header = False):
    if header:
        print(str(text))
    else:
        print("    > " + str(text))

def grabPartnerData():
    global PARTNERS, CONFIG
    displayText("Updating Partner Data. Calling out to "+CONFIG.Url+"....", True)
    tempDataFile = createTempFileName(".json")
    done = False #used to jump out of loop
    loops = 0
    limit = 100
    offset = 0
    temp = open(tempDataFile, "w")
    temp.write("{ \"partners\": [")
    numRows = 0

    scriptDir = os.path.dirname(__file__) #<-- absolute dir the script is in
    fileName = "partners.csv"
    filePath = CONFIG.OutputDirectory + "/" + fileName
    absFilePath = os.path.join(scriptDir, filePath)
    
    while(done == False):
        loops += 1
        response = sendAPIRequest(CONFIG, "GET", CONFIG.ApiUrl + "/profiles?profile_type_id="+CONFIG.PartnerProfileTypeId+"&query[limit]="+str(limit)+"&query[offset]="+str(offset),{})
        responseJSON = response.json()
        if "error" in responseJSON:
            displayText("Finished Fetching Profiles. A total of " + str(numRows) + " were found.")
            done = True
            break
        for p in responseJSON['profiles']:
            numRows += 1
            if numRows != 1:
                temp.write(",")
            json.dump(p,temp)
        
        displayText("Fetching more profiles. " + str(numRows) + " profiles fetched.")
        offset += limit
    temp.write("]}")
    temp.close()
    #Finished grabbing ALL partners
    with open(tempDataFile) as f:
        data = json.load(f)
        f = open(absFilePath, "w")
        #write the header first
        attributes = [
            "corporate_email",
            "id", 
            "name", 
            "status", 
            "updated_at", 
            "created_at",
            "sharepoint_invitation_date_ne_attribute",
            "sharepoint_invite_sent_by_ne_attribute",
            "personal_first_name",
            "personal_last_name",
            "partner_type_ne_attribute",
            "partner_organization_partners_ne_attribute"
        ]
        line = ""
        for att in attributes:
            line += att + ","
        line = line[:-1]
        f.write(line + "\n")
        line = ""
        for p in data['partners']:
            for att in attributes:
                if(att in p):
                    line = line + "\"" + p[att] + "\","
                elif(att in p["attributes"]):
                    line = line + "\"" + p["attributes"][att] + "\","
                else:
                    line = line + "\"\","
            line = line[:-1]
            f.write(line + "\n")
            line = ""
            PARTNERS.append(p)
            
    f.close()
    os.remove(tempDataFile)
    displayText("Success! Exported "+ str(numRows) +" profiles to " + str(fileName))

def processQuizData():
    global FIXES, CONFIG, QUIZ_ATTEMPTS
    loadFixesFile()
    displayText("Processing Quiz Data...", True)
    
    partner_by_email = {}
    partner_by_name = {}
    for partner in PARTNERS:
        partner_by_email[partner["attributes"]["corporate_email"].lower()] = partner
    
    # Write the initial row (header row) 
    writeLine('',True)
    # Loop thru quiz map and write to the file
    issues = 0
    scriptDir = os.path.dirname(__file__) #<-- absolute dir the script is in
    absPath = os.path.join(scriptDir, CONFIG.QuizDirectory)
    for quizId in CONFIG.QuizData:
        foundQuiz = False
        for fileName in sorted(os.listdir(absPath)):
            if fileName.endswith('.xlsx') and not(fileName.startswith("~")):
                if CONFIG.QuizData[quizId]["file_name"] in fileName:
                    foundQuiz = True
                    displayText("Found " + quizId + ". Processing...")
                    wb = xlrd.open_workbook(absPath + "/" + fileName)
                    sheet = wb.sheet_by_index(0)
                    attempts = 0
                    found = 0
                    skipped = 0
                    notFound = 0
                    for i in range(sheet.nrows):
                        if i == 0:
                            continue # header
                        attempts += 1
                        
                        id = str(sheet.cell_value(i,0)).strip()
                        if "." in id : #for some reason the ID's are being pulled in as decimals 62.0
                            id = id[:-2]
                        attemptID = str(quizId) + "-" + id
                        
                        email = ""
                        if attemptID in FIXES:
                            email = FIXES[attemptID]
                            if email == 'skip':
                                skipped+=1
                                continue
                        else:
                            email = sheet.cell_value(i,CONFIG.QuizData[quizId]["email_col"])
                        
                        if email.lower() in partner_by_email:
                            found+=1
                            partner = partner_by_email[email.lower()]
                            addQuizAttempt(partner, attemptID, quizId, CONFIG,sheet,i)
                        else:
                            issues += 1
                            notFound += 1
                            displayText(str(email) + " was not found! " + attemptID)
                            return
                    displayText("Finished Processing. Processed " + str(found) + " out of " + str(attempts) + ". Skipped:" + str(skipped))
        if foundQuiz == False:
            exit("ERROR - File Not Found: " + CONFIG.QuizData[quizId]["display_name"] + " has not been processed. Is the Excel file downloaded?")       
    displayText("Analyzing Quiz Results", True)
    displayText(str(QUIZ_ATTEMPTS) + " quiz attempts were processed.")
    displayText("All Quiz Data was written to " + str(CONFIG.OutputDirectory) + "/" str(CONFIG.TrainingDataFile))

def loadConfig():
    displayText("Loading Configuration...", True)
    f = open('config.json',)
    data = json.load(f) 
    CONFIG.Name = data["seczetta_environment"]["name"]
    CONFIG.Url = data["seczetta_environment"]["url"]
    CONFIG.ApiUrl = data["seczetta_environment"]["api_url"]
    CONFIG.ApiKey = data["seczetta_environment"]["api_key"]
    CONFIG.PartnerProfileTypeId = data["seczetta_environment"]["partner_profile_type_id"]
    CONFIG.QuizData = data["quiz_data"]
    CONFIG.TrainingDataFile = data["training_data_file_name"]
    CONFIG.FixesFile = data["fixes_file_name"]
    CONFIG.QuizDirectory = data["quiz_directory"]
    CONFIG.OutputDirectory = data["output_directory"]
    displayText("...Done")
    f.close() 

def sendAPIRequest(config, verb, url, data):

    verb = verb.upper()

    headers = {
        'Content-Type': 'application/json',
        'Authorization': 'Token token='+config.ApiKey,
        'Accept': 'application/json'
    }

    response = ""
    if(verb == "GET"):
        response = requests.get(url, headers=headers, json =data)

    if(verb == "POST"):
        response = requests.post(url, headers=headers, json = data)

    if(verb == "PATCH"):
        response = requests.patch(url, headers=headers, json = data)

    return response

def loadFixesFile():
    global FIXES,CONFIG
    displayText("Loading Manual Fixes...", True)
    with open(CONFIG.FixesFile, 'r') as read_obj:
        # pass the file object to reader() to get the reader object
        csv_reader = reader(read_obj)
        # Iterate over each row in the csv using reader object
        rowNum = 0
        for row in csv_reader:
            rowNum +=1
            if len(row) <= 0:
                continue #blank row?
            if row[0].startswith("#"):
                continue #comment
            if len(row) == 3:
                FIXES[row[0].strip()+"-"+row[1].strip()] = row[2].strip()
            else:
                raise ValueError('FixesFileError: Invalid Format on row number: ' + str(rowNum))
    displayText(str(len(FIXES)) + " manual fixes were loaded!")

def addQuizAttempt(partner, attempt, quiz, config, ws, row_index):
    global QUIZ_ATTEMPTS

    totalPoints = ws.cell_value(row_index, config.QuizData[quiz]["total_points_col"])
    pointsPossible = config.QuizData[quiz]["total_points_possible"]
    passingPercentage = config.QuizData[quiz]["passing_percent"]

    quizPercent = totalPoints / pointsPossible
    passPercent = quizPercent > passingPercentage
    status = "Passed"
    if passPercent == False:
        status = "Failed"


    QUIZ_ATTEMPTS += 1
    quizAttemptStr = ""
    quizAttemptStr += attempt + ","
    quizAttemptStr += config.QuizData[quiz]["display_name"] + ","
    quizAttemptStr += partner["name"] + ","
    quizAttemptStr += partner["attributes"]["corporate_email"] + ","
    quizAttemptStr += partner["attributes"]["partner_organization_partners_ne_attribute"] + ","
    quizAttemptStr += str(parseDate(ws.cell_value(row_index, config.QuizData[quiz]["start_time_col"]))) + ","
    quizAttemptStr += str(parseDate(ws.cell_value(row_index, config.QuizData[quiz]["complete_time_col"]))) + ","
    quizAttemptStr += str(totalPoints)
    writeLine(quizAttemptStr)

    print(partner["attributes"]["corporate_email"] + " has " + status + " " + quiz)




def parseDate(dateStr):
    dt = xlrd.xldate_as_datetime(dateStr,0)
    return dt.strftime("%m/%d/%Y %H:%M:%S")

welcome()
loadConfig()      
grabPartnerData() ## remember to uncomment the API Call
processQuizData()
