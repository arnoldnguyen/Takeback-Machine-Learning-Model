import pymongo
import pandas as pd
import numpy as np


def connect_to_db(db: str):
    client = pymongo.MongoClient(f"mongodb+srv://sameet:Maryland1234@mflix.vytuz.mongodb.net/{db}?retryWrites=true&w=majority")
    return client


"""
DATA RESTRUCTURING
"""

def restructuring_dfToDict(data: pd.DataFrame):
    """
    Restructure each row (corresponding to a club application) into a dictionary and then appends to a list

    :param data: pd.DataFrame

    :return list[dict]
    """
    template = {
        "Organization Name": ...,
        "Contact Person": ...,
        "SABO Number": ...,
        "Organization Advisor": ...,
        "Application": []
    }


    restructered_apps = []

    for index, row in data.iterrows():
        template_copy = template.copy()

        template_copy["Organization Name"] = row.filter(regex="Organization Name", axis=0).values[0]
        template_copy["Contact Person"] = row.filter(regex="Name of Contact Person", axis=0).values[0]
        template_copy["SABO Number"] = row.filter(regex="SABO", axis=0).values[0]
        template_copy["Organization Advisor"] = row.filter(regex="Organization Advisor", axis=0).values[0]
        template_copy["Application"].append(row["Do You Have Approved Storage Space on Campus":].to_dict())

        restructered_apps.append(template_copy)

    return restructered_apps



def restructuring_listOfDictToDf(data: list):
    """
    Restructures a list of dictionaries to a dataframe

    :param data: list
    :return: pd.Dataframe
    """

    return pd.DataFrame(data)



def restructuring_grabApplicationDataFromDf(data: pd.DataFrame):
    """
    Restructures all application data from club data into a dataframe

    :param data: pd.DataFrame
    :return: dataframe with all application data of the dataframe
    """
    application_data = data['Application']
    application_data_dict = {}

    count = 0

    for index, club in application_data.iteritems():
        for application in club:
            application_data_dict[count] = application
            count += 1

    return pd.DataFrame.from_dict(application_data_dict, "index")


def restructuring_grabSpecificDataFromApplicationData(applicationData: pd.DataFrame, *dtype):
    """
    From Application data, we aggregate columns along with attendance column

    ** MAKE SURE THAT COLUMNS SPECIFIED IN AGGREGATION HAVE FLOAT VALUES **

    :param applicationData: pd.DataFrame
    :param dtype: pass a variable number of column names to aggregate into a dataframe
    :return: pd.DataFrame
    """

    print(application_data.filter(regex="Attendance"))



if __name__ == "__main__":
    # client = connect_to_db(db="test")
    #
    # db = client["test"]
    # collection = db["myCollection"]
    #
    # test_post = {"_id": 0, "name": "test", "value": 5}
    #
    # collection.insert_one(test_post)

    data = pd.read_csv("FormSubmissions.csv")

    restructured_apps = restructuring_dfToDict(data)

    df_restructured_apps = restructuring_listOfDictToDf(restructured_apps)
    application_data = restructuring_grabApplicationDataFromDf(df_restructured_apps)

    restructuring_grabSpecificDataFromApplicationData(application_data, "Advertising")


    restructured_apps = restructuring_dfToDict(data)

    dataframe_apps = restructuring_listOfDictToDf(restructured_apps)


    print(restructuring_grabApplicationDataFromDf(dataframe_apps))
    print(dataframe_apps)









