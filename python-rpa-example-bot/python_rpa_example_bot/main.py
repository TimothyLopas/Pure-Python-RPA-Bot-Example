import base64
import boto3
import openpyxl
import pymsteams
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from simple_salesforce import Salesforce
from typing import List
import time
import os


def send_message_to_sfdc_messages_channel(message: str) -> None:
        """Sends a message to the SFDC Messages channel on MS Teams

        :param message: the text message you want to sent teams
        :return: None
        """
        uri = os.environ.get("TEAMSURI")
        my_teams_message = pymsteams.connectorcard(uri)
        my_teams_message.text(message)
        my_teams_message.send()

def main() -> None:
    # Let's use Amazon S3
    s3 = boto3.resource('s3')
    s3.meta.client.download_file("robocorp-test", "New/Review Cases.xlsx", "output/Review Cases.xlsx")

    wb = openpyxl.load_workbook(filename="output/Review Cases.xlsx")
    ws = wb["Sheet1"]
    cell_value_list = []
    for row in ws.iter_rows(min_row=2, max_col=1, max_row=100):
        for cell in row:
            if cell.value != None:
                cell_value_list.append(cell.value)

    # There are several ways to secure secrets in Python projects
    # One of the common ways is environment variables, which is what I have done here
    # One of the benefits of the Robocorp RPA solution is that secret management is a native part of the solution 
    sf_username = os.environ.get("SFUSERNAME")
    sf_password = os.environ.get("SFPASSWORD")
    sf_token = os.environ.get("SFTOKEN")
    sf = Salesforce(username=f'{sf_username}', password=f'{sf_password}', security_token=f'{sf_token}')

    
    address_list = []
    case_dict_master = []
    i = 0
    for item in cell_value_list:
        # Query Salesforce for the mailing address for each individual associated with a particular case
        case_query = sf.query(f"SELECT Id, ContactId FROM Case WHERE CaseNumber = '{item}'")
        case_id = case_query["records"][0]["Id"]
        contact_id = case_query["records"][0]["ContactId"]
        mailing_address_dict = sf.query(f"SELECT Id, MailingAddress FROM Contact WHERE Id = '{contact_id}'")
        mailing_address = f'{mailing_address_dict["records"][0]["MailingAddress"]["street"]} {mailing_address_dict["records"][0]["MailingAddress"]["city"]}, {mailing_address_dict["records"][0]["MailingAddress"]["state"]} {mailing_address_dict["records"][0]["MailingAddress"]["postalCode"]}'
        address_list.append(mailing_address)
        case_dict = {"case_id": case_id, "contact_id": contact_id, "mailing_address": mailing_address, "case": item}
        case_dict_master.append(case_dict)
        i += 1

    
    # Creating a webdriver instance
    driver = webdriver.Chrome()
    driver.get("https://www.google.com/maps")
    assert "Google Maps" in driver.title

    for address in address_list:
        elem = driver.find_element_by_id("searchboxinput")
        elem.clear()
        elem.send_keys(address)
        elem.send_keys(Keys.RETURN)
        assert "No results found." not in driver.page_source
        # Sleep is necessary here because the bot is faster than the page load
        time.sleep(5)
        driver.save_screenshot(f"output/{address}.png")

    driver.close()
    
    for x in range(i):
        caseFeed_data = {
            "ParentId" : case_dict_master[x]['case_id'],
            "Body" : case_dict_master[x]['mailing_address'],
            "Title" : case_dict_master[x]['mailing_address']
        }
        caseFeed = sf.FeedItem.create(caseFeed_data)

        binary_file_data = open(f"output/{case_dict_master[x]['mailing_address']}.png", 'rb').read()
        base64_encoded_data = base64.encodebytes(binary_file_data).decode('utf-8')
        contentVersion_data = {
            "VersionData" : f"{base64_encoded_data}",
            "PathOnClient" : f"output/{case_dict_master[x]['mailing_address']}.png",
            "FirstPublishLocationId" : "0058c000009XrFN",
            "Origin" : "H",
            "ContentLocation" : "S"
        }
        contentVersion = sf.ContentVersion.create(contentVersion_data)
        
        caseAttachment_data = {
            "Type" : "Content",
            "FeedEntityID" : caseFeed["id"],
            "RecordId" : contentVersion["id"]
        }
        sf.FeedAttachment.create(caseAttachment_data)

        base_url = os.environ.get("SFBASEURL")
        message = f"Case {case_dict_master[x]['case']} has been updated on SalesForce: https://{base_url}.lightning.force.com/lightning/r/Case/{case_dict_master[x]['case_id']}/view"
        send_message_to_sfdc_messages_channel(message)


    # delete all attachments
    # for x in range(i):
    #     attachment = sf.query(f"SELECT Id FROM Attachment WHERE ParentId = '{case_dict_master[x]['case_id']}'")
    #     print("-------------")
    #     print(attachment)
    #     print("------*******-----")
    #     print("---- RECORDS START ----")
    #     print(attachment["records"])
    #     print("---- RECORDS END ----")
    #     for item in range(attachment["totalSize"]):
    #         sf.Attachment.delete(attachment["records"][item]["Id"])


    # delete all ContentDocuments
    # document = sf.query("SELECT Id FROM ContentDocument WHERE FileType = 'png'")
    # print("-------------")
    # print(document)
    # print("------*******-----")
    # for item in range(document["totalSize"]):
    #      sf.ContentDocument.delete(document["records"][item]["Id"])


def list_bucket_names() -> List:
    # Let's use Amazon S3
    s3 = boto3.resource('s3')

    # Create empty List
    bucket_names = []

    # Print out bucket names
    for bucket in s3.buckets.all():
        print(bucket.name)
        bucket_names.append(bucket.name)
    
    return bucket_names