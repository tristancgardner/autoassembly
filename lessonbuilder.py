# 7/2 NEED TO TEST 3 MOGRTs with one each, and then with all

# Get all lesson titles via headers and send to csv
# Add mogrt at time code
# Add mogrt and adjust length

from __future__ import print_function

import os.path
import csv

import pandas as pd
import streamlit as st

from fuzzywuzzy import fuzz

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/drive.file',
          'https://www.googleapis.com/auth/drive.metadata',
          'https://www.googleapis.com/auth/documents']

creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())
        
file_id = '1qJXKLs_7bb6ybcEIMIxs6TADHQAwGgry89hyTvnh1Qg'
document_id = file_id


def getTitle():

    try:
        service = build('docs', 'v1', credentials=creds)
        # Gets all headers for lesson titles
        document = service.documents().get(documentId=document_id).execute()
        title = document['title']
        print()
        print(f"Document Title: {title}")
        print()

        if not title:
            print('No files found.')
            return

    except HttpError as error:
            print('An error occurred: %s' % error)

def parseComments():

    # Get Comments in Doc
    try:
        
        service = build('drive', 'v3', credentials=creds)
        comments = service.comments().list(fileId=file_id,fields='comments').execute()
        items = comments.get('comments', [])
        
        if not items:
            print('No files found.')
            return
        
        titleCounter = 0
        defCounter = 0
        calloutCounter = 0

        titles = ['Title']
        callouts = ['Callouts']
        def_definitions = ['Definitions']
        def_titles = ['Definition Title']

        
        for item in items:
            # the for loop makes 'item' the iterable for each item in the array 'items', it's more of a for loop syntax thing than it is an actual variable, just like you don't need to declare 'i' when writing a for loop in java

            # this just makes it so I can type less when referencing the content of each item iterated through in the array 'items'
            content = item['content']

            target_Title = 'EXO_A_Title'
            target_Callout = 'EXO_E_Callout'
            target_Definition = 'EXO_D_Definition'
            
            
            # DEFINITION MOGRT
            if 'Tit' in content:
                    
                # Calculate the similarity score between the content and target_substring
                similarity_score = fuzz.ratio(content, target_Title)
                # Define a threshold for considering a match
                threshold = 30
    
                if similarity_score >= threshold:
                    titleCounter = titleCounter + 1
                    substring = 'Title:'
                    index = content.find(substring) + len(substring) + 1
                    print(content)
                    print('')
                    print('Info going to CSV --> ', content[index:])
                    print('------------------------------------------')
                    #appends retreived title_mogrt comments into an array for temp storage, we'll use this array to print to a csv in a certain way in contrast to that of other mogrts
                    titles.append(content[index:])
                    
            # DEFINITION MOGRT
            elif 'Def' in content:
                similarity_score = fuzz.ratio(content, target_Definition)
                threshold = 30
    
                if similarity_score >= threshold:
                    defCounter = defCounter + 1
                    print(content)
                    substring = 'Title:'
                    index1 = content.find(substring) + len(substring) + 1
                    substring2 = 'Body:'
                    index2 = content.find(substring2) + len(substring2) + 1
                    print('')
                    print('Info going to CSV --> ')
                    print(content[index1:], content[index2:])
                    print('------------------------------------------')
                    def_definitions.append(content[index1:])
                    def_titles.append(content[index2:])
                    
            # CALLOUT MOGRT
            elif 'Cal' in content:
                similarity_score = fuzz.ratio(content, target_Callout)
                threshold = 30
    
                if similarity_score >= threshold:
                    calloutCounter = calloutCounter + 1
                    print(content)
                    substring = 'Phrase:'
                    index = content.find(substring) + len(substring) + 1
                    print('')
                    print('Info going to CSV --> ', content[index:])
                    print('------------------------------------------')
                    callouts.append(content[index:])

            #all other MOGRTs are accounted for here with additional "elif" blocks
              
            #saving this as a reminder of the other fields I can call of the comment object in the JSON comment files for gDocs
            """
            content = item['content']
            author_name = item['author']['displayName']
            created_time = item['createdTime']
    
            print('Comment:', content)
            print('Author:', author_name)
            print('Created Time:', created_time)
            print('---')
            """

        print('')
        print('                      --- Methods QA ---')   
        print('# of Title MOGRTs: ', titleCounter)   
        print('# of Definition MOGRTs: ', defCounter)
        print('')


        print('Last Item in Retrieved Titles:\n', titles[0])
        print('')

        # Write CSV _____________________________________________________________________________________________________________________________
        with open('output.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)

            # Writing titles and definitions as rows in the CSV file
            for Title, D_Title, D_Definition in zip(titles, def_definitions, def_titles):
                writer.writerow([Title, D_Title, D_Definition])

            # Handling missing definitions
            remaining_titles = titles[len(def_definitions):]
            for title in remaining_titles:
                writer.writerow([title, ""])

            print("CSV file written successfully.")



    except HttpError as error:
        print('An error occurred: %s' % error)

def getLessons():
    try:
        service = build('docs', 'v1', credentials=creds)
        # Gets all headers for lesson titles
        document = service.documents().get(documentId=document_id).execute()
        
        # Extract the header elements from the document
        lesson_titles = []
        for element in document['body']['content']:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                if 'paragraphStyle' in paragraph and 'namedStyleType' in paragraph['paragraphStyle']:
                    named_style_type = paragraph['paragraphStyle']['namedStyleType']
                    if named_style_type == 'HEADING_2':
                        lesson_titles.append(paragraph['elements'][0]['textRun']['content'])

        print('\n', lesson_titles, '\n') 
        
        # Write CSV _____________________________________________________________________________________________________________________________
        with open('lessons.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            
            for lesson in lesson_titles:
                writer.writerow([lesson])

            print("CSV file written successfully.")

    except HttpError as error:
            print('An error occurred: %s' % error)  



# Run Functions _____________________________________________________________________________________________________________________________
#if __name__ == '__main__':
    #getTitle()
    #parseComments()
    #getLessons()
    
    
# Send Spreadsheet data to streamlit_____________________________________________________________________________________________________________________________
run_button = st.button("Run Script")
def run_script():
    # Code to run your Python script
    getLessons()
    df = pd.read_csv('lessons.csv')
    st.dataframe(df)
    st.write("Script executed successfully!")


if run_button:
    run_script()




