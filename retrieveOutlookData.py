import numpy as np
import pandas as pd
from pathlib import Path
import datetime
import re
import uuid
import summarizer
import streamlit as st
# import win32com.client  # pip install pywin32



def extract(count):
    """Get emails from outlook."""
    items = []
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox
    messages = inbox.Items
    message = messages.GetFirst()
    i = 0
    while message:
        try:
            msg = dict()
            msg["Subject"] = getattr(message, "Subject", "<UNKNOWN>")
            msg["SentOn"] = getattr(message, "SentOn", "<UNKNOWN>")
            msg["EntryID"] = getattr(message, "EntryID", "<UNKNOWN>")
            msg["Sender"] = getattr(message, "Sender", "<UNKNOWN>")
            msg["Size"] = getattr(message, "Size", "<UNKNOWN>")
            msg["Body"] = getattr(message, "Body", "<UNKNOWN>")
            items.append(msg)
        except Exception as ex:
            print("Error processing mail", ex)
        i += 1
        if i < count:
            message = messages.GetNext()
        else:
            return items

    return items


def show_message(items):
    """Show the messages."""
    items.sort(key=lambda tup: tup["SentOn"])
    for i in items:
        print(i["SentOn"], i["Subject"], i["Sender"])

def save_csv(items):
    """Save the messages to a csv file."""
    import csv

    # Create csv file
    with open('emails.csv', 'w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["SentOn", "Subject", "Sender", "Size", "Body"])
        for i in items:
            writer.writerow([i["SentOn"], i["Subject"], i["Sender"], i["Size"], i["Body"]])

    print("CSV file saved.")

def preprocess(text):
    if len(text) > 1000:
        text = text[:900]
    text = truncate(text,100)
    return text

def truncate(text, length):
    words = text.split()
    for i in range(len(words)):
        if len(words[i]) > length:
            words[i] = words[i][:length]

    return ' '.join(words)

def predisplay(df,i):
    datetime_object = datetime.datetime.strptime(df['SentOn'][i], '%Y-%m-%d %H:%M:%S+00:00')
    st.markdown(
        """
        <style>
        .custom-header {
            font-size:25px;
            padding-bottom: 0px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    st.markdown(f"<h2 class='custom-header'>{df['Subject'][i]}<h2>", unsafe_allow_html=True)
    st.write(datetime_object.strftime('%A, %d %B %Y %I:%M %p')," | ", df['Sender'][i])

def streamlit_helper():
    st.set_page_config(layout="wide")
    # Main content
    st.title("Outlook Mail Summarizer :sunglasses:")
    


def main():
    items = extract(100)
    save_csv(items)
    streamlit_helper()

    df = pd.read_csv('emails.csv')
    body = df['Body']
    # display data in columns of 3 for each loop
    
    # call the summarizer function in summarizer.py file in same directory
    i = 0
    while i<len(df):
        col1, col2, col3 = st.columns(3)
        with col1:
            predisplay(df,i)
            # Create a column to store the summary
            df['Summary'] = np.array(['nan' for i in range(len(df))])
            with st.spinner('Summarizing...'):
                if body[i] == 'nan':
                    continue
                df.loc[i, 'Body'] = preprocess(df.loc[i, 'Body'])

                # if the summary is not already present summarize the email
                if (df['Summary'][i] == 'nan'):
                    summary = summarizer.summarizer(body[i])
                else:
                    summary = df['Summary'][i]

                st.info(summary)
            # Add a column to the dataframe
            df.loc[i, 'Summary'] = summary
            i = i+1
        
        if i <len(df):
            with col2:
                predisplay(df,i)
                # Create a column to store the summary
                df['Summary'] = np.array(['nan' for i in range(len(df))])
                with st.spinner('Summarizing...'):
                    if body[i] == 'nan':
                        continue
                    df.loc[i, 'Body'] = preprocess(df.loc[i, 'Body'])

                    # if the summary is not already present summarize the email
                    if (df['Summary'][i] == 'nan'):
                        summary = summarizer.summarizer(body[i])
                    else:
                        summary = df['Summary'][i]

                    st.info(summary)
                # Add a column to the dataframe
                df.loc[i, 'Summary'] = summary
                i = i+1
        if i <len(df):
            with col3:
                predisplay(df,i)
                # Create a column to store the summary
                df['Summary'] = np.array(['nan' for i in range(len(df))])
                with st.spinner('Summarizing...'):
                    if body[i] == 'nan':
                        continue
                    df.loc[i, 'Body'] = preprocess(df.loc[i, 'Body'])

                    # if the summary is not already present summarize the email
                    if (df['Summary'][i] == 'nan'):
                        summary = summarizer.summarizer(body[i])
                    else:
                        summary = df['Summary'][i]

                    st.info(summary)
                # Add a column to the dataframe
                df.loc[i, 'Summary'] = summary
                i = i+1
        df.to_csv('emails.csv', index=False)
        

if __name__ == "__main__":
    main()
