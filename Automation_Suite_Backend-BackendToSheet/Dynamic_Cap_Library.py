import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from constants import *
import numpy as np
from sklearn.linear_model import LinearRegression
import math



"""
FOR POPULATING CAP SHEET TEMPLATE ON DRIVE 

"""
def validate_credentials():
    """
    Establishes connection between files on google cloud that have enabled access to Service Account, as specified in JSON file

    :return: client with access privileges
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('ConnectPybackendWithModel-647da95d34c7.json', scope)
    client = gspread.authorize(creds)

    return client





"""

CALCULATING CAP VALUES ON DATASET  

"""


"""
Generating Cap Values from Transformed DataFrames

"""


def LinearRegression_on_data(request_attendance: pd.DataFrame):
    """
    Filters data by interquartile range and applies linear regression to get cap values

    :param request_attendance: pd.DataFrame
    :return:
    """
    #Perform Linear Regression on dataframe

    Q1 = request_attendance['request/attendance'].quantile(0.25)
    Q3 = request_attendance['request/attendance'].quantile(0.75)


    interQuartileRange = Q3 - Q1

    mask = (request_attendance['request/attendance'] > (Q1 - 1.5 * interQuartileRange)) & (request_attendance['request/attendance'] < (Q3 + 1.5 * interQuartileRange))
    filtered_data = request_attendance[mask]

    X = filtered_data.iloc[:, 1].values.reshape(-1, 1)
    Y = filtered_data.iloc[:, 2].values.reshape(-1, 1)

    Q1 = request_attendance['Expected Attendance'].quantile(0.25)
    Q3 = request_attendance['Expected Attendance'].quantile(0.75)


    interQuartileRange = Q3 - Q1

    mask = (request_attendance['Expected Attendance'] > (Q1 - 1.5 * interQuartileRange)) & (request_attendance['Expected Attendance'] < (Q3 + 1.5 * interQuartileRange))
    filtered_data = request_attendance[mask]


    linear_regressor = LinearRegression()
    linear_regressor.fit(X, Y)

    minimum = int(filtered_data['Expected Attendance'].min())
    maximum = int(filtered_data['Expected Attendance'].max())

    step = math.floor((maximum - minimum) / 5)

    attendance_caps = np.asarray([val for val in range(minimum + step, maximum, step)])
    unitrates = linear_regressor.predict(attendance_caps.reshape(-1, 1))


    return attendance_caps, unitrates



"""
Transforming base dataframe to desired structure

"""


def filter(row: pd.Series):
    try:
        request = float(row.loc['Request'])
        expected_attendance = float(row.loc['Expected Attendance'])

        return request / expected_attendance
    except (ValueError, ZeroDivisionError):
        return np.nan


def Transform_GetRequestToAttendance(data: pd.DataFrame):
    # Calculate Unit Rate
    request_attendance = data.loc[:, ['Request', 'Expected Attendance']]

    request_attendance['request/attendance'] = request_attendance.apply(filter, axis=1)
    request_attendance.dropna(inplace=True)

    request_attendance['Expected Attendance'] = request_attendance['Expected Attendance'].apply(lambda row: float(row))

    # print(request_attendance.head(5))

    return request_attendance


def Standaloneprogramming_controller(data: pd.DataFrame, indices: tuple, type: str):
    """
           returns list in form [XS, S, M, L, XL] where each idx is an int value

           :param data: pd.DataFrame
           :return: list
       """
    df = application_data.iloc[:, indices[0]:indices[1]]
    df.columns = ['SABO', 'Request', 'Expected Attendance', 'Allocation']

    attendance_caps, unitrates = LinearRegression_on_data(Transform_GetRequestToAttendance(df))

    cap_values = [round(a * b, -1) for a, b in zip(attendance_caps, unitrates.flatten())]

    with open("caps.txt", "a") as caps:
        print(f"caps for {type} - programming stand alone: {cap_values}")
        caps.write(f"caps for {type} - programming stand alone: {cap_values}\n")


if __name__ == "__main__":


    application_data = pd.read_excel(io="Fall_2020.xlsm", skiprows=1)


    #Grab Stand Alone associated with room rental
    Standaloneprogramming_controller(data=application_data, indices=(0, 4), type="Room Rental")
    Standaloneprogramming_controller(data=application_data, indices=(4, 8), type="Advertising")
    Standaloneprogramming_controller(data=application_data, indices=(8, 12), type="Food")
    Standaloneprogramming_controller(data=application_data, indices=(12, 16), type="Supplies/Decorations")



# if __name__ == "__main__":
#     client = validate_credentials()
#
#     sheet = client.open(DYNAMIC_CAP_FILE)
#
#     sheet_instance = sheet.get_worksheet(0)
#
#     records_data = sheet_instance.get_all_records()
#
#     print(records_data)

