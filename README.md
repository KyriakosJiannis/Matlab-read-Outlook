# Matlab function which reads Outlook emails
Author :  Ioannis Kyriakos

Matlab function which imports the 'readed' or 'unreaded' Outlook emails from inbox, folders and sub-folders. 
<br />Extracts their subjects, bodies and can save their attachments.

% Reads all emails from inbox
<br />mails = ReadOutlook;

% Reads all Unread emails from inbox
<br />mails = ReadOutlook('Read', 1);

% Reads all Unread emails from inbox and marks them as read
<br />mails = ReadOutlook('Read', 1, 'Mark', 1);

% Reads all emails from folder and save their attachments
<br />mails = ReadOutlook('Folder', 'Groups', 'Savepath', 'C:\matlab\data\test');
