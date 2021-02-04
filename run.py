import xlrd
import os
from os import path
import json
import csv
from csv import reader
import random
import requests
from datetime import datetime
from enum import Enum

#### CLASSES



class TrainingAttributes():
    TRAINING_STATUS_101 = "101_training_status"
    TRAINING_STATUS_201 = "201_training_status"
    TRAINING_STATUS_301 = "301_training_status"
    TRAINING_STATUS_LESSON00 =   "lesson_00_training_status"
    TRAINING_STATUS_LESSON01 =   "lesson_01_training_status"
    TRAINING_STATUS_LESSON02 =   "lesson_02_training_status"
    TRAINING_STATUS_LESSON03 =   "lesson_03_training_status"
    TRAINING_STATUS_LESSON04 =   "lesson_04_training_status"
    TRAINING_STATUS_LESSON05 =   "lesson_05_training_status"
    TRAINING_STATUS_LESSON06 =   "lesson_06_training_status"
    TRAINING_STATUS_LESSON0708 = "lesson_0708_training_status"
    TRAINING_STATUS_LESSON09 =   "lesson_09_training_status"
    TRAINING_STATUS_LESSON1011 = "lesson_1011_training_status"
    TRAINING_STATUS_LESSON12 =   "lesson_12_training_status"
    TRAINING_STATUS_LESSON13 =   "lesson_13_training_status"
    TRAINING_STATUS_LESSON14 =   "lesson_14_training_status"
    TRAINING_STATUS_LESSON16 =   "lesson_16_training_status"
    TRAINING_COMPLETE_DATE_101 = "101_training_completed_date"
    TRAINING_COMPLETE_DATE_201 = "201_training_completed_date"
    TRAINING_COMPLETE_DATE_301 = "301_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON00 = "lesson_00_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON01 = "lesson_01_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON02 = "lesson_02_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON03 = "lesson_03_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON04 = "lesson_04_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON05 = "lesson_05_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON06 = "lesson_06_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON0708 = "lesson_0708_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON09 = "lesson_09_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON1011 = "lesson_1011_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON12 = "lesson_12_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON13 = "lesson_13_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON14 = "lesson_14_training_completed_date"
    TRAINING_COMPLETE_DATE_LESSON16 = "lesson_16_training_completed_date"

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

class TrainingDetails:
    status_101  = "Not Completed"
    status_201 = "Not Completed"
    status_301 = "Not Completed"
    status_lesson00 = "Not Completed"
    status_lesson01 = "Not Completed"
    status_lesson02 = "Not Completed"
    status_lesson03 = "Not Completed"
    status_lesson04 = "Not Completed"
    status_lesson05 = "Not Completed"
    status_lesson06 = "Not Completed"
    status_lesson0708 = "Not Completed"
    status_lesson09 = "Not Completed"
    status_lesson1011 = "Not Completed"
    status_lesson12 = "Not Completed"
    status_lesson13 = "Not Completed"
    status_lesson14 = "Not Completed"
    status_lesson16 = "Not Completed"
    date_101 = None
    date_201 = None
    date_301 = None
    date_00 = None
    date_01 = None
    date_02 = None
    date_03 = None
    date_04 = None
    date_05 = None
    date_06 = None
    date_0708 = None
    date_09 = None
    date_1011 = None
    date_12 = None
    date_13 = None
    date_14 = None
    date_16 = None

    def convertToEnum(self, attName):
        if "101" in attName and "lesson" not in attName.lower():
            if "date" not in attName:
                self.status_101 = TrainingAttributes.TRAINING_STATUS_101
            else:
                self.status_101 = TrainingAttributes.TRAINING_COMPLETE_DATE_101
        elif "201" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_201
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_201
        elif "301" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_301
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_301
        elif "00" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON00
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON00
        elif "01" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON01
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON01
        elif "02" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON02
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON02
        elif "03" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON03
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON03
        elif "04" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON04
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON04
        elif "05" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON05
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON05
        elif "06" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON06
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON06
        elif "0708" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON0708
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON0708
        elif "09" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON09
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON09
        elif "1011" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON1011
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON1011
        elif "12" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON12
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON12
        elif "13" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON13
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON13
        elif "14" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON14
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON14
        elif "16" in attName:
            if "date" not in attName:
                return TrainingAttributes.TRAINING_STATUS_LESSON16
            else:
                return TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON16
        return None

    def toString(self):
        return str(self.status_101) + " :: " + str(self.status_201) + " :: " + str(self.date_101)

    def tryToGetAttribute(self, obj,attName, exceptionValue):
        try:
            return obj[attName]
        except:
            return exceptionValue

    def loopAndSetAttributes(self, obj):
        self.status_101 = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_101,self.status_101)
        self.status_201 = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_201,self.status_201)
        self.status_301 = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_301,self.status_301)
        self.status_lesson00   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON00,self.status_lesson00)
        self.status_lesson01   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON01,self.status_lesson01)
        self.status_lesson02   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON02,self.status_lesson02)
        self.status_lesson03   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON03,self.status_lesson03)
        self.status_lesson04   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON04,self.status_lesson04)
        self.status_lesson05   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON05,self.status_lesson05)
        self.status_lesson06   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON06,self.status_lesson06)
        self.status_lesson0708 = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON0708,self.status_lesson0708)
        self.status_lesson09   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON09,self.status_lesson09)
        self.status_lesson1011 = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON1011,self.status_lesson1011)
        self.status_lesson12   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON12,self.status_lesson12)
        self.status_lesson13   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON13,self.status_lesson13)
        self.status_lesson14   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON14,self.status_lesson14)
        self.status_lesson16   = self.tryToGetAttribute(obj,TrainingAttributes.TRAINING_STATUS_LESSON16,self.status_lesson16)
        self.date_101  = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_101,self.date_101)
        self.date_201  = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_201,self.date_201)
        self.date_301  = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_301,self.date_301)
        self.date_00   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON00,self.date_00)
        self.date_01   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON01,self.date_01)
        self.date_02   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON02,self.date_02)
        self.date_03   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON03,self.date_03)
        self.date_04   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON04,self.date_04)
        self.date_05   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON05,self.date_05)
        self.date_06   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON06,self.date_06)
        self.date_0708 = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON0708,self.date_0708)
        self.date_09   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON09,self.date_09)
        self.date_1011 = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON1011,self.date_1011)
        self.date_12   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON12,self.date_12)
        self.date_13   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON13,self.date_13)
        self.date_14   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON14,self.date_14)
        self.date_16   = self.tryToGetAttribute(obj, TrainingAttributes.TRAINING_COMPLETE_DATE_LESSON16,self.date_16)


    def __init__(self, obj):
        if "attributes" in obj:
            self.loopAndSetAttributes(obj["attributes"])
        else:
            self.loopAndSetAttributes(obj)


### GLOBALS 
PARTNERS = []
FIXES = {}
QUIZ_ATTEMPTS = 0
PARTNER_ATTEMPTS = {} # "email" "[ "attempt1", "attempt2", etc.. ]"
PARTNER_TRAINING_STATUS = {} #"id": { "101" : { "status" : "Pass", "completeDate" : "01/01/1900"}}
FINAL_RESULTS = {}
CONFIG = Config()


#### FUNCTIONS
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
        combinedFile.write('email,id,101_training_status,101_training_completed_date,201_training_status,201_training_completed_date,301_training_status,301_training_completed_date,lesson_00_training_status,lesson_00_training_completed_date,lesson_01_training_status,lesson_01_training_completed_date,lesson_02_training_status,lesson_02_training_completed_date,lesson_03_training_status,lesson_03_training_completed_date,lesson_04_training_status,lesson_04_training_completed_date,lesson_05_training_status,lesson_05_training_completed_date,lesson_06_training_status,lesson_06_training_completed_date,lesson_0708_training_status,lesson_0708_training_completed_date,lesson_09_training_status,lesson_09_training_completed_date,lesson_1011_training_status,lesson_1011_training_completed_date,lesson_12_training_status,lesson_12_training_completed_date,lesson_13_training_status,lesson_13_training_completed_date,lesson_14_training_status,lesson_14_training_completed_date,lesson_16_training_status,lesson_16_training_completed_date\n')
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
    pullNewData = True
    hardcodedTempFile = "temp-4321741582.json"
    global PARTNERS, CONFIG
    displayText("Updating Partner Data. Calling out to "+CONFIG.Url+"....", True)
    done = False #used to jump out of loop
    loops = 0
    limit = 40
    offset = 0
    if pullNewData: 
        tempDataFile = createTempFileName(".json")
        temp = open(tempDataFile, "w")
        temp.write("{ \"partners\": [")
    numRows = 0
    scriptDir = os.path.dirname(__file__) #<-- absolute dir the script is in
    fileName = "partners.csv"
    filePath = CONFIG.OutputDirectory + "/" + fileName
    absFilePath = os.path.join(scriptDir, filePath)
    
    if pullNewData:
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
    if pullNewData == False:
        tempDataFile = hardcodedTempFile

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
    if pullNewData:
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
                        
                        if "@seczetta.com" in email:
                            skipped+=1
                            continue

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
    displayText("All Quiz Data was written to " + str(CONFIG.OutputDirectory) + "/" + str(CONFIG.TrainingDataFile))

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
    global QUIZ_ATTEMPTS, PARTNER_ATTEMPTS

    email = partner["attributes"]["corporate_email"].lower()
    totalPoints = ws.cell_value(row_index, config.QuizData[quiz]["total_points_col"])
    pointsPossible = config.QuizData[quiz]["total_points_possible"]
    passingPercentage = config.QuizData[quiz]["passing_percent"]

    quizPercent = totalPoints / pointsPossible
    passPercent = quizPercent > passingPercentage
    result = "Pass"

    if passPercent == False:
        result = "Fail"

    QUIZ_ATTEMPTS += 1
    quizAttempt = {}
    quizAttempt["attempt_id"] = attempt
    quizAttempt["quiz_id"] = quiz
    quizAttempt["quiz_display_name"] = CONFIG.QuizData[quiz]["display_name"]
    quizAttempt["partner_name"] = partner["name"]
    quizAttempt["partner_org"] = ["partner_organization_partners_ne_attribute"]
    quizAttempt["start_time"] = str(parseDate(ws.cell_value(row_index, config.QuizData[quiz]["start_time_col"])))
    quizAttempt["complete_time"] = str(parseDate(ws.cell_value(row_index, config.QuizData[quiz]["complete_time_col"])))
    quizAttempt["total_points"] = totalPoints
    quizAttempt["total_points_possible"] = pointsPossible
    quizAttempt["result"] = result

    if email in PARTNER_ATTEMPTS:
        PARTNER_ATTEMPTS[email].append(quizAttempt) 
    else:
        PARTNER_ATTEMPTS[email] = [quizAttempt]


    # quizAttemptStr = ""
    # quizAttemptStr += attempt + ","
    # quizAttemptStr += config.QuizData[quiz]["display_name"] + ","
    # quizAttemptStr += partner["name"] + ","
    # quizAttemptStr += partner["attributes"]["corporate_email"] + ","
    # quizAttemptStr += partner["attributes"]["partner_organization_partners_ne_attribute"] + ","
    # quizAttemptStr += str(parseDate(ws.cell_value(row_index, config.QuizData[quiz]["start_time_col"]))) + ","
    # quizAttemptStr += str(parseDate(ws.cell_value(row_index, config.QuizData[quiz]["complete_time_col"]))) + ","
    # quizAttemptStr += str(totalPoints)
    # writeLine(quizAttemptStr)

    

    #print(partner["attributes"]["corporate_email"] + " has " + result + " " + quiz)

def finalizeTrainingResults():
    global PARTNERS, PARTNER_ATTEMPTS, PARTNER_TRAINING_STATUS, FINAL_RESULTS
    displayText("Finalizing Results...", True)
    finalResults = {}
    for p in PARTNERS:
        finalResults[p["attributes"]["corporate_email"].lower()] = {
            "id": p["id"],
            "101": "Not Completed",
            "201": "Not Completed",
            "301": "Not Completed",
            "Lesson 00": "Not Completed",
            "Lesson 01": "Not Completed",
            "Lesson 02": "Not Completed",
            "Lesson 03": "Not Completed",
            "Lesson 04": "Not Completed",
            "Lesson 05": "Not Completed",
            "Lesson 06": "Not Completed",
            "Lesson 0708": "Not Completed",
            "Lesson 09": "Not Completed",
            "Lesson 1011": "Not Completed",
            "Lesson 12": "Not Completed",
            "Lesson 13": "Not Completed",
            "Lesson 14": "Not Completed",
            "Lesson 16": "Not Completed"
        }
        PARTNER_TRAINING_STATUS[p["attributes"]["corporate_email"].lower()] = { 
            "id": p["id"]
        }

    i = 0
    j = 0
    for email in PARTNER_ATTEMPTS:    
        i+=1
        for attempt in PARTNER_ATTEMPTS[email]:
            j+=1
            email = email.lower()
            if attempt["quiz_id"] == "101":
                if attempt["result"] == "Pass":
                    finalResults[email]["101"] = "Pass"
                    finalResults[email]["101_training_status"] = "Pass"
                    finalResults[email]["101_training_completed_date"] = attempt["complete_time"]
                else:
                    if finalResults[email]["101"] == "Not Completed":
                        finalResults[email]["101"] = "Fail"  
                        finalResults[email]["101_training_status"] = "Fail"    

            elif attempt["quiz_id"] == "201":
                if attempt["result"] == "Pass":
                    finalResults[email]["201"] = "Pass"
                    finalResults[email]["201_training_status"] = "Pass"
                    finalResults[email]["201_training_completed_date"] = attempt["complete_time"]
                else:
                    if finalResults[email]["201"] == "Not Completed":
                        finalResults[email]["201"] = "Fail"
                        finalResults[email]["201_training_status"] = "Fail"
            
            else: 
                #should be lessons now
                quizId = attempt["quiz_id"]
                quizIdNoSpaces = attempt["quiz_id"].replace(" ","_").lower()
                if attempt["result"] == "Pass":
                    finalResults[email][quizId] = "Pass"
                    finalResults[email][quizIdNoSpaces + "_training_status"] = "Pass"
                    finalResults[email][quizIdNoSpaces + "_training_completed_date"] = attempt["complete_time"]
                    if checkIfPartnerPassed301(finalResults[email]):
                        finalResults[email]["301"] = "Pass"
                        finalResults[email]["301_training_status"] = "Pass"
                        finalResults[email]["301_training_completed_date"] = attempt["complete_time"]
                else:
                    if finalResults[email][quizId] == "Not Completed":
                        finalResults[email][quizId] = "Fail"
                        finalResults[email][quizIdNoSpaces + "_training_status"] = "Fail"
            #print(email + " attempted quiz " + attempt["quiz_display_name"] + " and has " + attempt["result"])
        writeLine('',True)
    FINAL_RESULTS = finalResults
    displayText("Done! " + str(len(PARTNER_ATTEMPTS)) + " unique partners have attempted at least one quiz.")

def buildTrainingCSV():
    global FINAL_RESULTS
    displayText("Writing Training Results to csv...", True)
    for email in FINAL_RESULTS:
        line = ""
        line += email + ","
        line += FINAL_RESULTS[email]["id"] + ","
        #101 Training Status
        line += getTrainingStatus(email, FINAL_RESULTS[email], "101")
        line += getTrainingStatus(email, FINAL_RESULTS[email], "201")
        line += getTrainingStatus(email, FINAL_RESULTS[email], "301")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 00")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 01")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 02")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 03")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 04")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 05")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 06")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 0708")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 09")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 1011")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 12")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 13")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 14")
        line += getTrainingStatus(email, FINAL_RESULTS[email],"Lesson 16")
        line = line[:-1]

        writeLine(line)
    displayText("..Done")

def syncTrainingToSecZetta():
    global FINAL_RESULTS, PARTNERS
    displayText("Syncing Training Results to IRS....", True)
    syncActions = {} #id: { "101_training_status": { "new_value": "Pass", "old_value": "Not Completed" } }
    updatedResults = {}
    currentResults = {} 
    syncItems = {}
    attributesByEmail = {}
    for p in PARTNERS:
        sync = {}
        id = p["id"]
        email = p["attributes"]["corporate_email"].lower()
        result = FINAL_RESULTS[email]
        updatedResults[email] = TrainingDetails(result)
        currentResults[email] = TrainingDetails(p["attributes"])
        sync = compareTrainingResults(TrainingDetails(result), TrainingDetails(p["attributes"]))
        attributesByEmail[email] = { "id" : id, "profile_type_id": p["profile_type_id"], "status": p["status"]}
        if len(sync) > 0:
           syncItems[email] = sync
    processSyncItems(attributesByEmail,syncItems)
    displayText("Done!")

def compareTrainingResults(new, old):
    sync = {}
    if new.status_101 != old.status_101:
        sync[TrainingAttributes.TRAINING_STATUS_101] = { "new": new.status_101, "old": old.status_101}
    if new.status_201 != old.status_201:
        sync[TrainingAttributes.TRAINING_STATUS_201] = { "new": new.status_201, "old": old.status_201}
    if new.status_301 != old.status_301:
        sync[TrainingAttributes.TRAINING_STATUS_301] = { "new": new.status_301, "old": old.status_301}
    if new.status_lesson00 != old.status_lesson00:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON00] = { "new": new.status_lesson00, "old": old.status_lesson00}
    if new.status_lesson01 != old.status_lesson01:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON01] = { "new": new.status_lesson01, "old": old.status_lesson01}
    if new.status_lesson02 != old.status_lesson02:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON02] = { "new": new.status_lesson02, "old": old.status_lesson02}
    if new.status_lesson03 != old.status_lesson03:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON03] = { "new": new.status_lesson03, "old": old.status_lesson03}
    if new.status_lesson04 != old.status_lesson04:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON04] = { "new": new.status_lesson04, "old": old.status_lesson04}
    if new.status_lesson05 != old.status_lesson05:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON05] = { "new": new.status_lesson05, "old": old.status_lesson05}
    if new.status_lesson06 != old.status_lesson06:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON06] = { "new": new.status_lesson06, "old": old.status_lesson06}
    if new.status_lesson0708 != old.status_lesson0708:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON0708] = { "new": new.status_lesson0708, "old": old.status_lesson0708}
    if new.status_lesson09 != old.status_lesson09:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON09] = { "new": new.status_lesson09, "old": old.status_lesson09}
    if new.status_lesson1011 != old.status_lesson1011:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON1011] = { "new": new.status_lesson1011, "old": old.status_lesson1011}
    if new.status_lesson12 != old.status_lesson12:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON12] = { "new": new.status_lesson12, "old": old.status_lesson12}
    if new.status_lesson13 != old.status_lesson13:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON13] = { "new": new.status_lesson13, "old": old.status_lesson13}
    if new.status_lesson14 != old.status_lesson14:
        print("NEW: " + str(new.status_lesson14) + " OLD: " + str(old.status_lesson14))
        sync[TrainingAttributes.TRAINING_STATUS_LESSON14] = { "new": new.status_lesson14, "old": old.status_lesson14}
    if new.status_lesson16 != old.status_lesson16:
        sync[TrainingAttributes.TRAINING_STATUS_LESSON16] = { "new": new.status_lesson16, "old": old.status_lesson16}
    
    return sync

def processSyncItems(attributesByEmail, syncItems):
    global CONFIG
    displayText("There are " + str(len(syncItems)) +  " items that need synced")
    inputStr = ""
    for email in syncItems:
        items = syncItems[email]
        inputStr = email + " has " + str(len(items)) + " attribute(s) out of sync"
        attributes = {}
        for item in items:
            #inputStr += "\n > Attribute: " + item + " is currently '" + items[item]["old"] + "' and it should be '" + items[item]["new"] + "'"
            attributes[item] = items[item]["new"]
        displayText(inputStr)
        choice = input("\n    > Would you like to update " + email + "'s training attributes? (Y):")
        if choice.lower() == 'y' or choice == "":
            id = attributesByEmail[email]["id"]
            body = { 
                "profile": {
                    "profile_type_id": attributesByEmail[email]["profile_type_id"],
                    "status": attributesByEmail[email]["status"],
                    "attributes": attributes
                }
            }
            response = sendAPIRequest(CONFIG,"PATCH",CONFIG.ApiUrl + "/profiles/" + id, body)
            if response.status_code == 200:
                displayText("200 OK - Successfully updated " + email)
            else:
                displayText(str(response.status_code) + " - Something went wrong\n\n")
                displayText(response.json())
                print("\n\n")
        else:
            displayText("\nNot Updating " + email + ".")
        
def checkForSyncAction(szAttributes, newAttributes, szAttName, newAttName):
    if szAttName not in szAttributes: 
        return True
    szVal = szAttributes[szAttName]
    newVal = newAttributes[newAttName]
    if szVal != newVal:
        return True
    return False

def getTrainingStatus(email, results, quizId):
    quizIdNoSpaces = quizId.replace(" ","_").lower()
    line2 = ""
    if results[quizId] == "Pass":
        line2 += results[quizIdNoSpaces + "_training_status"] + ","
        line2 += results[quizIdNoSpaces + "_training_completed_date"] + ","
    else:
        line2 += results[quizId] + ","
        line2 += ","
        

    return line2

def checkIfPartnerPassed301(results):
    if results["Lesson 00"] != "Pass":
        return False
    if results["Lesson 01"] != "Pass":
        return False
    if results["Lesson 02"] != "Pass":
        return False
    if results["Lesson 03"] != "Pass":
        return False
    if results["Lesson 04"] != "Pass":
        return False
    if results["Lesson 05"] != "Pass":
        return False
    if results["Lesson 06"] != "Pass":
        return False
    if results["Lesson 0708"] != "Pass":
        return False
    if results["Lesson 09"] != "Pass":
        return False
    if results["Lesson 1011"] != "Pass":
        return False
    if results["Lesson 12"] != "Pass":
        return False
    if results["Lesson 13"] != "Pass":
        return False
    if results["Lesson 14"] != "Pass":
        return False
    if results["Lesson 16"] != "Pass":
        return False
    return True

def parseDate(dateStr):
    dt = xlrd.xldate_as_datetime(dateStr,0)
    return dt.strftime("%m/%d/%Y")

welcome()
loadConfig()      
grabPartnerData() ## remember to uncomment the API Call
processQuizData()
finalizeTrainingResults()
buildTrainingCSV()
syncTrainingToSecZetta()