# budgetcalculator
Budget Calculator

First creates an output folder for the email attachment to be saved in.

Using the Outlook API we search through the emails to search for a specific title.

We retrieve the latest email with that heading and download its attachments to the output folder (along with the body of the email saved in a text file - this isn't useful for this project)

We read through the PDF (note: this only works for specific banks due to the formatting) extracting information such as the date, description, transaction amount and current balance

This is all put into an SQL database.

From there it is put into an Excel sheet where the program recognises certain words in the description and categorises them for further analysis.
