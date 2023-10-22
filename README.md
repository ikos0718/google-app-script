# google-app-script
Google App Script to retrieve the contacts from Gmail

The script performs the following tasks
**1.	Contact Retrieval:**
Extracts contacts' information, including their name, email, and company link, from both the inbox and the sent box in Gmail.

**2.	Menu Addition:**
Adds a 'Contacts' menu to the Google Sheet interface, facilitating the retrieval of contacts.

**3.	Label Creation:**
Creates a label named 'Contact_Saved' and applies it to emails that have already been processed and their contacts saved. It will save the script's execution time because you won't need to retrieve the contacts that have already been saved.

**4.	Sheet Renaming:**
Renames the default sheet to 'Contacts' sheet, where the retrieved contact information will be stored.

**5.	Exclusion of Special Emails:**
Filters out specific email addresses from the contact list to be excluded from the retrieval process.

**6.	Exclusion of Special Domains:**
Filters out specific domains from the company list to be excluded from the retrieval process.

**7.	Filter Creation:**
Sets up a filter for the columns 'No,' 'Name,' 'Email,' and 'Company Link' in the 'Contacts' sheet to enable efficient data organization and retrieval.
