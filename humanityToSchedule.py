# Humanity to Team IS Schedule

from __future__ import print_function
import pickle   #pickle rick
import os.path
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


    print("")
    print("")
    print("Put custom Humanity report into")
    print("") # Google Sheets URL
    print("");
    print("Humanity > Reports > Custom Reports")
    print("Report Type: Shifts Scheduled, This Week")
    print("Select Employee, Start Day, Start Time, End Time")
    print("Save as CSV and replace values in Google Sheets Report")
    print("")
    print("")
    print("Put custom Humanity report into")
    print("") # Google Sheets URL
    print("")
    print("Humanity > Reports > Daily Peak Hours")
    print("Select This Week, 30 minutes, Location: LSU Res Life")
    print("Save as CSV and replace values in Google Sheets Report")
    print("")
    print('In Help Desk Schedules Document (), duplicate "Default" and rename it to current Semester') # Google Sheets URL
    print("")
    semester = input("Input Semester Name: ")

    #toSchedule(service, semester)
    toHeatmap(service, semester)

def toSchedule(service, semester):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId="", range="A:D").execute()
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
    service.spreadsheets().values().update(spreadsheetId="", range= semester + "!A3:AE", valueInputOption="USER_ENTERED", body={"values" : employeesArr}).execute()

def toHeatmap(service, semester):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId="", range="U3:AO16").execute()
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

    # Only get rows corresponding to correct employee group
    leadsList = [values[0]] + [values[3]] + [values[6]] + [values[9]] + [values[12]]
    techsList = [values[1]] + [values[4]] + [values[7]] + [values[10]] + [values[13]]

    # Transpose lists
    leadsList = transpose(leadsList)
    techsList = transpose(techsList)

    # Update Google sheet
    service.spreadsheets().values().update(spreadsheetId="", range= semester + "!C33:G54", valueInputOption="USER_ENTERED", body={"values" : techsList}).execute()
    service.spreadsheets().values().update(spreadsheetId="", range= semester + "!J33:N54", valueInputOption="USER_ENTERED", body={"values" : leadsList}).execute()

# https://www.geeksforgeeks.org/python-transpose-elements-of-two-dimensional-list/
# Python program to get transpose 
# elements of two dimension list 
def transpose(l1): 
    l2 = []
    # iterate over list l1 to the length of an item  
    for i in range(len(l1[0])): 
        # print(i) 
        row =[] 
        for item in l1: 
            # appending to new list with values and index positions 
            # i contains index position and item contains values 
            row.append(item[i]) 
        l2.append(row) 
    return l2 

if __name__ == '__main__':
    main()