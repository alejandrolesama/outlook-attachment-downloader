# Outlook Attachment Downloader

Python tool to download Outlook email attachments using Microsoft Graph API.

## Description

This project allows users to connect to an Outlook or Microsoft email account, filter emails by date range, detect emails with attachments, and automatically download those attachments into local folders.

The authentication process is handled securely through Microsoft Graph and does not require storing the email password inside the code.

## Features

- Secure authentication with Microsoft Graph API.
- Email filtering by date range.
- Detection of emails with attachments.
- Automatic attachment download.
- Folder organization by email date and subject.
- Windows-compatible file and folder name cleaning.

## Technologies Used

- Python
- Microsoft Graph API
- MSAL
- Requests
- Visual Studio Code
- Git and GitHub

## Project Structure

```text
outlook-attachment-downloader/
│
├── main.py
├── requirements.txt
├── config.example.py
├── README.md
└── .gitignore