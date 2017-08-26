# Matlab function which reads Outlook emails
Author :  Ioannis Kyriakos

Matlab function which imports the 'readed' or 'unreaded' outlook emails from inbox and their folders - subfolders. 
Extracts their subjects, bodies and can save their attachments.


% Reads all emails from inbox
mails = ReadOutlook;

% Reads all Unread emails from inbox
mails = ReadOutlook('Read', 1);

% Reads all Unread emails from  inbox and mark them as read
mails = ReadOutlook('Read', 1, 'Mark', 1);

% Reads all emails from a email folders and save their attachments
mails = ReadOutlook('Folder', 'Groups', 'Savepath', 'C:\matlab\data\test');
