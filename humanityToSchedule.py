# Humanity to Team IS Schedule

from __future__ import print_function
import pickle   #pickle rick
import os.path
# pip install --upgrade google-api-python-client
# pip install google-auth-oauthlib
# https://console.cloud.google.com/apis/credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)


    ssId = input("Input Speadsheet ID (The string after URL): ")

    semester = input("Input Semester Name: ")

    print("")
    print('Duplicate the sheet called "Default" and rename it to "' + semester + '"')
    print("")

    toSchedule(service, semester, ssId)
    toHeatmap(service, semester, ssId)

    print("Complete.")

def toSchedule(service, semester, ssId):
    shouldSchedule = input("Humanity Report to Schedule? (Y or N)").lower()
    if (shouldSchedule[0] == "n"):
        return
    elif (shouldSchedule[0] != "y"):
        return toSchedule(service, semester, ssId)

    print("")
    print("")
    print("Navigate to Humanity > Reports > Custom Reports")
    print("Report Type: Shifts Scheduled, This Week - Apply")
    print("Select Employee, Start Day, Start Time, End Time - Apply")
    print("Save as CSV")
    print('Create a new sheet under "Help Desk Schedules" and paste the values from the Humanity Report into it')
    print("")
    print("")
    input("Press Enter when the above steps are complete.")
    print("")

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ssId, range="A:D").execute()
    values = result.get('values', [])

    mostRecentDay = values[1][1]
    employees = {
        # "Seth Williamson" : [[Monday], [Tuesday], [Wednesday], [Thursday], [Friday]]
    }

    # Create dictionary with names for all employees
    # Set default values for each time entry (in and out 3 times per day for each day of week)
    for i in range(1, len(values)):
        if values[i][0] not in employees:
            employees[values[i][0]] = [["","","","","",""], ["","","","","",""], ["","","","","",""], ["","","","","",""], ["","","","","",""]]

    day = 0

    for i in range(1, len(values)):
        # Changes day for array when day changes
        if values[i][1] != mostRecentDay:
            mostRecentDay = values[i][1]
            day += 1

        # Append values to dictionary
        for j in range(0, 6):
            if not employees[values[i][0]][day][j]:
                employees[values[i][0]][day][j] = values[i][2]
                employees[values[i][0]][day][j+1] = values[i][3]
                break

    # Since sheets needs an x*y list / array, this for loop appends all the values created in the dicitonary to list^2
    employeesArr = []

    for i in range(0, len(employees)):
        currentEmployee = employees.popitem()
        employeesArr.append([currentEmployee[0]])
        for j in range(0, 5):
            for k in range(0, 6):
                employeesArr[i].append(currentEmployee[1][j][k])

    # Update Google sheet
    service.spreadsheets().values().update(spreadsheetId=ssId, range= semester + "!A3:AE", valueInputOption="USER_ENTERED", body={"values" : employeesArr}).execute()
    print("You may delete the sheet with Humanity Report info")
    print("")

def toHeatmap(service, semester, ssId):

    shouldHeatmap = input("Humanity Report to Heatmap? (Y or N)").lower()
    if (shouldHeatmap[0] == "n"):
        return
    elif (shouldHeatmap[0] != "y"):
        return toHeatmap(service, semester, ssId)

    print("")
    print("")
    print("Navigate to Humanity > Reports > Daily Peak Hours")
    print("Select This Week, 30 minutes, Location: LSU Res Life")
    print("Save as CSV")
    print('Create a new sheet under "Help Desk Schedules" and paste the values from the Humanity Report into it')
    print("")
    print("")
    input("Press Enter when the above steps are complete.")
    print("")

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ssId, range="U2:AO27").execute()
    values = result.get('values', [])
    
    # Converting all values from string to floats
    for inner in values:
        for index, string in enumerate(inner):
            if not string:
                string = 0
            inner[index] = float(string)

        # The matrix must be nxn
        # 8:00 to 6:30 has 21 30 minute intervals
        while len(inner) < 21:
            inner.append(0)

    print(len(values))

    # Only get rows (weekdays) corresponding to correct employee group
    techsListOnCampus = [values[0]] + [values[5]] + [values[10]] + [values[15]] + [values[20]]
    leadsListOnCampus = [values[1]] + [values[6]] + [values[11]] + [values[16]] + [values[21]]
    leadsListWFH = [values[2]] + [values[7]] + [values[12]] + [values[17]] + [values[22]]
    techsListWFH = [values[3]] + [values[8]] + [values[13]] + [values[18]] + [values[23]]

    # Transpose lists
    techsListOnCampus = transpose(techsListOnCampus)
    leadsListOnCampus = transpose(leadsListOnCampus)
    leadsListWFH = transpose(leadsListWFH)
    techsListWFH = transpose(techsListWFH)

    techsList = addMatrices(techsListOnCampus, techsListWFH)
    leadsList = addMatrices(leadsListOnCampus, leadsListWFH)

    WFHList = addMatrices(leadsListWFH, techsListWFH)
    onCampusList = addMatrices(techsListOnCampus, leadsListOnCampus)
    
    fullList = addMatrices(leadsList, techsList)

    print(techsList)
    print(fullList)

    # Update Google sheet
    service.spreadsheets().values().update(spreadsheetId=ssId, range= semester + "!C33:G54", valueInputOption="USER_ENTERED", body={"values" : fullList}).execute()
    service.spreadsheets().values().update(spreadsheetId=ssId, range= semester + "!Q33:U54", valueInputOption="USER_ENTERED", body={"values" : techsList}).execute()
    service.spreadsheets().values().update(spreadsheetId=ssId, range= semester + "!J33:N54", valueInputOption="USER_ENTERED", body={"values" : leadsList}).execute()

    service.spreadsheets().values().update(spreadsheetId=ssId, range= semester + "!C56:G76", valueInputOption="USER_ENTERED", body={"values" : onCampusList}).execute()
    service.spreadsheets().values().update(spreadsheetId=ssId, range= semester + "!J56:N76", valueInputOption="USER_ENTERED", body={"values" : WFHList}).execute()

    print("You may delete the sheet with Humanity Report info")
    print("")

# https://www.geeksforgeeks.org/python-transpose-elements-of-two-dimensional-list/
# Python program to get transpose 
# elements of two dimension list 
def transpose(l1): 
    l2 = []
    # iterate over list l1 to the length of an item  
    for i in range(len(l1[0])): 
        row =[] 
        for item in l1: 
            # appending to new list with values and index positions 
            # i contains index position and item contains values 
            row.append(item[i]) 
        l2.append(row) 
    return l2 

def addMatrices(m1, m2):
    result = []

    for i in range(len(m1)):
        result.append([])
        for j in range(len(m1[i])):
            result[i].append(m1[i][j] + m2[i][j])
        
    return result


if __name__ == '__main__':
    main()

