# Google APIs
from googleapiclient.discovery import build

# Working with Google Sheets and Slides
from google.oauth2 import service_account
import gspread

# Pandas and numpy to work with df
import pandas as pd
import numpy as np

# Open AI API
import openai

# Streamlit for desktop App
import streamlit as st


def extract_data_from_gs(filename, service_account_file, scopes):

    try:

        creds = None  # Clears credentials in case it catches an exception
        # Loads service account credentials
        creds = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)

        # Authorizing Google client, opening the file and retrieving the data from it
        client = gspread.authorize(creds)
        google_sheet = client.open(filename).sheet1
        result = google_sheet.get_all_records()

        # Cleaning missing and faulty entries in the data
        df = pd.DataFrame.from_records(result, columns=result[0])
        header = str(df.columns.tolist())  # Converting the header to string
        df.replace('', np.nan, inplace=True)

        deleted_rows = df[df.isnull().any(axis=1)]  # Stores the deleted rows
        df.dropna(inplace=True)

        # Checks all columns with numbers and converts it to int so filtering operations can be conducted at later stages
        df = df.apply(pd.to_numeric, errors='ignore')
        numeric_columns = df.select_dtypes(include='number').columns
        df[numeric_columns] = df[numeric_columns].astype(int)

        if len(deleted_rows) > 0:
            st.write(f"The following rows with missing or wrong data have been excluded ❌")
            st.write(deleted_rows) # Displays deleted rows

        return df, header

    except gspread.SpreadsheetNotFound as e:
        raise e


def chatgpt_query(prompt, header):
    # What is sent onto the AI model as a request
    initial_prompt = 'I am going to pass a query in order to find all employees who meet certain requirements inside of a pandas dataframe. Translate those requirements in terms of the header which is ' + header + ' \n\nGive the commands in relationship with the structure of the given header. The answer needs to follow the following df queries in the next examples:\nExample1\nQuery: I want all employees whos name is Jon and live in Spain\n\ndf.loc[(df["Name"] == "Jon") & (df["Country"] == "Spain")]\n\nExample 2\nQuery: I want all employees that live in the UK and are Data scientists\ndf.loc[(df["Country"] == "UK") & (df["Occupation"] == "Data scientist")]\n\nYour answer must be precise and short, no explanation JUST the code. Write it in plane text so it is easier to read. You are only allowed to use one single line of code. You get bonus points for shorter answer and for the least amount of code lines. You are NOT allowed to write 2 lines of code or you will lose points. If you give the answer in 2 separate lines of code, you will be deducted points from the total score'
    full_prompt = initial_prompt + '\nNow my query is:\n' + prompt + '\n Only provide your answer with the code'

    # Tuning of the AI model parameters allowing to further customize the response
    openai.api_key = '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' # This is my personal OpenAI API key
    model_engine = 'text-davinci-003'  # The most advanced AI model currently offered by the API

    completion = openai.Completion.create(engine=model_engine,
                                          prompt=full_prompt,
                                          temperature=0.6,
                                          max_tokens=1000,
                                          top_p=1,
                                          frequency_penalty=0.2,
                                          presence_penalty=0)

    # Answer from the request
    text = completion.choices[0].text
    text = text.replace("\n", "")
    text = text.lstrip('.')

    return text


def is_email(input):  # Checks if input is an email
    return '@' in input and '.' in input


def share_slide_copies(dataframe, share_email):

    if is_email(share_email) is True:
        # Defining the Placeholders on the Google Slides
        placeholders_list = ["**Employee ID**",
                             "**Employee Name**",
                             "**Occupation**",
                             "**Country**",
                             "**Age**"]
        st.write("Sharing, please hold tight")

        # Empty list to store the copies IDs
        copies = []

        # Iterates over the filtered df results
        for i in range(len(dataframe)):
            # Set the path to your service account credentials JSON file
            credentials_path = 'C:/Users/sergi/Downloads/roche2023.json'

            # Necessary authorizations to access the Google Slides template file, create copies and share them
            scopes = ['https://www.googleapis.com/auth/drive',
                      'https://www.googleapis.com/auth/drive.file',
                      'https://www.googleapis.com/auth/presentations']

            source_slide_id = '190x1G-7DH6zaWEJjTI49sdoJ9ZVr_tyxW05QB_2W-SY'  # ID of the Google Slides template

            # Build the credentials and Sheets service
            credentials = service_account.Credentials.from_service_account_file(credentials_path, scopes=scopes)
            drive_service = build('drive', 'v3', credentials=credentials)  # Google Drive service
            slides_service = build('slides', 'v1', credentials=credentials)  # Google Slides service

            # Extracts info of the employee from the filtered df
            name = dataframe.iloc[i]['Name']
            employer_id = str(dataframe.iloc[i]['ID'])

            # Sets the name of the copied file to be saved with the employee Name and ID
            copy_title = name + ' - ' + employer_id
            copy_body = {'name': copy_title}

            # Makes a copy of the template file and stores the response
            copy_response = drive_service.files().copy(fileId=source_slide_id, body=copy_body).execute()
            copies.append(copy_response['id'])

            # Store the batch updates requests
            requests = []

            list_of_info = dataframe.iloc[i].tolist()
            list_of_info = [str(item) for item in list_of_info]

            for placeholder, info in zip(placeholders_list, list_of_info):
                # Replaces placeholders info with employee info
                requests.append({"replaceAllText": {"containsText": {"text": placeholder, "matchCase": False}, "replaceText": info}})
                slides_service.presentations().batchUpdate(presentationId=copies[i], body={'requests': requests}).execute()

            # Share the copied Google Sheets files with the specified email
            drive_service.permissions().create(
                fileId=copies[i],
                body={'type': 'user', 'role': 'writer', 'emailAddress': share_email},
                fields='id',
                sendNotificationEmail=False).execute()  # Important to set it to False to avoid exceeding usage limit
        st.write(f"{len(dataframe)} copies of the Google Sheets file created and shared with {share_email}.")

    else:
        st.write("Please type in a valid email address")


def main():
    st.title("Data Filtering Desktop App")

    # Input field for Google Sheets filename
    filename = st.text_input("Enter the Google Sheets filename:")
    csv_file_button = st.file_uploader("Upload a CSV file")
    if csv_file_button is not None:
        try:
            df = pd.read_csv(csv_file_button)
            st.dataframe(df)
        except pd.errors.EmptyDataError:
            st.write("The CSV file is empty")

    service_account_file = 'C:/Users/sergi/Downloads/roche2023.json'
    scopes = ['https://www.googleapis.com/auth/spreadsheets',
              'https://www.googleapis.com/auth/drive',
              'https://www.googleapis.com/auth/presentations']

    if filename:
        try:
            # Read Google Sheets data
            df, header = extract_data_from_gs(filename, service_account_file, scopes)
            # Display the dataframe
            st.write("Cleaned Google Sheets file loaded successfully ✔")
            st.dataframe(df)
            st.divider()

            # Selectbox for choosing operation
            st.write("Data Filter Options")
            st.markdown("* By Name (select only with employee data)")
            st.markdown("* By AI (select with any file that you want) ")
            st.markdown("* By Other Parameters (select only with employee data)")
            option = st.selectbox('Choose:', ('Filter by Name', 'Filter with AI', 'Filter by Other Parameters'))

            if option == 'Filter by Name':
                # Filter employees based on name
                filter_name = st.text_input("Enter a name to filter employees:")
                if filter_name:
                    filtered_df = df[df['Name'] == filter_name]
                    st.write("Filtered Data:")
                    st.dataframe(filtered_df)
                    st.divider()
                    st.write("Do you wish to create a Google Slides file for each employee info")
                    share_email = st.text_input("Type in the email address to share the Google Slides with and click SHARE")

                    if st.button("SHARE"):
                        share_slide_copies(filtered_df, share_email)

            elif option == 'Filter with AI':
                filter_ai = st.text_input("Write any query as a command")
                st.write("Not sure what to write? Try one of these:")
                st.markdown("* Example 1: I want to find all employees older than 30 that live in Spain or the UK")
                st.markdown("* Example 2: Give me all employees that are a Data Analyst and live in Switzerland ")
                st.markdown("* Example 3: I want to know which cars were fabricated between 1999 and 2008 and also run on diesel")

                # Perform some AI operation on the data
                if filter_ai:
                    answer = chatgpt_query(filter_ai, header)

                    # Check if the loaded file has the employee structure by comparing the df header
                    employee_header = str(['ID', 'Name', 'Occupation', 'Country', 'Age'])
                    if header == employee_header:
                        try:
                            filtered_df = eval(answer)
                            st.write("Filtered Data:")
                            st.dataframe(filtered_df)
                            st.divider()
                            st.write("Do you want to create a Google Slides file for each employee info?")
                            share_email = st.text_input("Type in the email address to share the Google Slides with and click SHARE")

                            if st.button('SHARE'):
                                share_slide_copies(filtered_df, share_email)

                        except SyntaxError:
                            st.write('Apologies, I am an Artificial Intelligence but sometimes I still mess up 🤖 Try again now')
                    else:
                        filtered_df = eval(answer)
                        st.write("Filtered Data:")
                        st.dataframe(filtered_df)
                        st.write('At the moment this App only supports creating and sharing Google Slides for employee data')

            # Perform filter by Age, Occupation and Country
            elif option == 'Filter by Other Parameters':
                try:
                    # Age slider
                    values = st.slider('Select a range of values', 18, 100, (25, 65))
                    # Occupation selectbox
                    occupation_options = st.multiselect(
                        'Select profession/s',
                        ['Data Analyst', 'Data Scientist', 'Software Engineer', 'Developer', 'Accountant', 'Executive', 'Intern'],
                        ['Data Analyst', 'Software Engineer'])
                    # Country selectbox
                    country_options = st.multiselect(
                        'Select country',
                        ['France', 'Spain', 'Switzerland', 'USA', 'UK', 'Germany', 'Italy','India', 'Netherlands', 'Austria'],
                        ['UK', 'Germany', 'Italy'])
                    # Perform df query based on all previous specifications
                    filtered_df = df[(df['Age'] >= values[0]) & (df['Age'] <= values[1])
                                     & df['Occupation'].isin(occupation_options)
                                     & df['Country'].isin(country_options)]
                    st.divider()
                    st.write('Filtered Data:')
                    st.dataframe(filtered_df)
                    st.write("Do you wish to create a Google Slides file for each employee info")
                    share_email = st.text_input(
                        "Type in the email address to share the Google Slides with and click SHARE")

                    if st.button("SHARE"):
                        share_slide_copies(filtered_df, share_email)

                except KeyError:  # Catches exception if the file does not have the employee structure
                    st.write('Please load a file with the employee structure')

        except gspread.SpreadsheetNotFound:  # Exception handling of not the input filename
            st.write(f"Sorry, could not find file: {filename} :mag:")
            st.write("Try the following:")
            st.write("1. Make sure that you typed in the filename correctly")
            st.write("2. Make sure that your Google Sheets file is located on your Google Drive")
            st.write("3. Make sure that it has been shared with the service account")


if __name__ == '__main__':
    main()
