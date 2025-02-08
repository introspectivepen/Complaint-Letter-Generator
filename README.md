Complaint Letter Generator and Sender

Overview

This Flask-based web application allows users to generate complaint letters for financial and non-financial fraud cases. The letters are automatically created in a .docx format, saved to a specified directory, and sent via email to the recipient.

Features

Users can submit a complaint via a web form.

Complaints can be of type financial or non-financial.

Generates a .docx complaint letter with a logo and relevant details.

Saves the generated letter to a specified directory.

Sends the complaint letter as an email attachment to the recipient.

Technologies Used

Python (Flask, re, datetime, os, smtplib)

Flask for handling HTTP requests and rendering templates

Python-docx for generating .docx files

smtplib for sending emails

Installation

Prerequisites

Ensure you have Python installed along with the required libraries:

pip install flask python-docx

Clone the Repository

git clone https://github.com/your-repository.git
cd your-repository

Configuration

Update Email Credentials

Modify the following section with your sender email and app password:

sender_email = "your-email@gmail.com"
sender_password = "your-app-password"

Note: Ensure that you have enabled "Less Secure Apps" or used an App Password if using Gmail.

Update Logo and Save Path

Replace the logo path and save directory with appropriate values:

logo_path = r'C:\path\to\logo_crime.jpeg'
save_path = r'C:\path\to\letters'

Running the Application

Start the Flask server:

python app.py

The application will be accessible at http://127.0.0.1:5000/.

Usage

Open the web application in a browser.

Fill in the required complaint details.

Click submit to generate and send the letter.

Check the specified directory for saved letters.

The recipient will receive an email with the letter as an attachment.

Error Handling

Invalid emails return an error message.

Missing form data results in an "Incomplete form data" message.

SMTP errors are caught and displayed.
