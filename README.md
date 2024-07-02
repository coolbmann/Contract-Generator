https://github.com/coolbmann/Contract-Generator/assets/117170435/3272f99f-7a86-4b55-8a7f-f1d60822de5d

A script for creating PDFs with dynamic information.

## 📄 Overview 
This project was created out of a real-world business need.

In my first few months at MILKRUN, part of my responsibility having owned the rider supply and onboarding function was to create employment contracts for ~50 new hires per week.

This script allowed me to save 3.5 hours per week, which would otherwise have been used manually clicking and copy/pasting field data such as names and addresses into a template document before sharing via DocuSign.

As part of the project, I wrote a comprehensive how-to guide which allowed me to pass on this tool to the next owner as I shifted my area of responsibility away from onboarding.

## 💡 Technologies Used 
JavaScript was used to write the automation, run on the Google Apps Script platform.

This allowed me to leverage Google Sheets as the UI and provided seamless connection with Google Drive as the document storage solution.
## 🔥 Features 
- Auto-fill existing placeholder text with real data;
- Select destination folder in Google drive;
- Generate an unlimited amount of documents in one run;
- Automatically convert finished document into a downloadable PDF.

## ⚡ See it in Action!
View a working copy of the sheet [here](https://docs.google.com/spreadsheets/d/1Ed2Xymlpe5NcSQnswnuqDThG6TL_uQRcjjUwlS0zsis/edit?gid=0#gid=0).

View the user guide [here](https://docs.google.com/document/d/1wKKooBaRSzwePu-RoUYEMHRTRGERWN5HNzFlIfRi7L8/edit).


## ⚙️ Setup Guide

Code in main.js contains some methods which are native to the Google Apps Script platform.

To run your own copy of the code, create a new Apps Script file and paste the code in the editor.

Follow the prompts to allow Google Apps Script to have access to read and write into your Google Account.

You can edit the location variable in the start() function to specify the destination folder for newly created documents.

You can also interact with this code live via this demo [Google Sheet](https://docs.google.com/spreadsheets/d/1Ed2Xymlpe5NcSQnswnuqDThG6TL_uQRcjjUwlS0zsis/edit?gid=0#gid=0).
