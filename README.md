### automating_custom_batch_word_document_email_delivery

- the purpose of this project is to help automate the customization od word documents (letters)
- and delivering them to their respective emails.


##### Purpose

Had 1000's of letters to send and decide to write a script to edit the sample letter and then send emails to the letter recievers.

##### How it Works

- It works by having 2 files. 1 word document and 1 csv. 
- The csv holds the user information and email details which will be used to replace the placeholders in the document.
- The word document contains the sample letter or text with placeholders indecating what needs to be changed in the word document.

- We first read from the csv and then go on ahead to make edits in the word document. there are someplaces that needs to be changed for use case to work.


##### RUN

- Clone the repository.
- Put the document and the excel file to be used for processing within the repository folder.
- Edit the python file and replace the documet in the python file with the name of your word document and do same with the excel document.
- Run the python file like any other python file.
