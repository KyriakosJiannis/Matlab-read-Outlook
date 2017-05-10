function [Email]= ReadOutlook2007_v2(varargin)
%  Scraping emails from Microsoft Outlook 2007
%  Functionality which imports read or unread emails from inbox or 
%  or its subfolders
%  Extracts subjects, bodies and can save their attachements
%  
% It is able read emails 
%           from Inbox
%                Inbox/Folder
%                Inbox/Folder/Subfolder
%
% Inputs
%    Basic import functionality
%    Varargin: SQL style
%    ------------------------------
%       Folder    : outlook folder name
%       Subfolder = outlook subfolder
%       Savepath  = path to save the attachments
%       Read      = 1,  reads only the UnRead emails, 
%                   else empty ''
%       Mark      = 1,  marks UnRead emails as read, 
%              else empty ''
%  ReadOutlook2007(Folder,Subfolder,Savepath,Read,Mark)
%--------------------------------------------------------------------------
% Examples:
% %     Reads all emails from a subfolder and save the attchments
%         mails = ReadOutlook2007(...
%             'Folder', 'Groups',...
%             'Subfolder', 'ce-ig',...
%             'Savepath', 'C:\matlab\data\test')
% %     Reads all unread emails from a subfolder        
%          mails = ReadOutlook2007(...
%             'Folder', 'Groups',...
%             'Subfolder', 'ce-ig',...
%             'Read',1);        
% %     Reads all unread emails from a subfolder and mark tham as read       
%          mails = ReadOutlook2007(...
%             'Folder', 'Groups',...
%             'Subfolder', 'ce-ig',...
%             'Read',1 , ...
%             'Mark',1);                                  
% %     Reads all emails from your inbox
%          mails = ReadOutlook2007
% %     Reads all Unread emails from your inbox
%          mails = ReadOutlook2007(...
%              'Read',1)
%--------------------------------------------------------------------------

%% Function imports
vargs = varargin;
f = Function_Varargin(vargs);
clearvars varargin vargs

%% Connects to Outlook2007
outlook = actxserver('Outlook.Application');
mapi = outlook.GetNamespace('mapi');
INBOX = mapi.GetDefaultFolder(6);

%% Retrieving UnRead or read emails / save or not save attachments
if isempty(f.Folder) && isempty(f.Subfolder)
    % reads Inbox only
    count = INBOX.Item.Count;
    Email = cell(count,2);
elseif ~isempty(f.Folder)
    % reads Inbox folder
    folder_numbers = INBOX.Folders.Count;
    % find folder / subfolder's outlookindex
    for i = 1:folder_numbers
        name = INBOX.Folders(1).Item(i).Name;
        if strcmp(name,f.Folder)
            n = i;
        end
    end    
    switch f.Subfolder
        % working for folder emails
        case ''
            % number of emails
            count = INBOX.Folders(1).Item(n).Items.Count;
            % cell for emailbody
            Email = cell(count,2);
        otherwise
            % Search for nth Inbox folder and count sub-folders
            folder_numbers = INBOX.Folders(1).Item(n).Folders(1).Count;
            % find Outlook Subfolder Index
            for i=1:folder_numbers
                name = INBOX.Folders(1).Item(n).Folders(1).Item(i).Name;
                if strcmp(name,f.Subfolder)
                    s= i;
                end
            end
            % number of emails
            count = INBOX.Folders(1).Item(n).Folders(1).Item(s).Items.Count;
            % cell for emailbody
            Email = cell(count,2);
    end
end

%% download & read emails
for i = 1:count
    if f.Read == 1 %****only unreads emails
        %inbox
        if isempty(f.Folder) && isempty(f.Subfolder)
            UnRead = INBOX.Items.Item(count+1-i).UnRead;
        %folder
        elseif ~isempty(f.Folder) && isempty(f.Subfolder)
            UnRead = INBOX.Folders(1).Item(n).Items(1).Item(count+1-i).UnRead;
        %subfolder
        elseif ~isempty(f.Folder) && ~isempty(f.Subfolder)
            UnRead = INBOX.Folders(1).Item(n).Folders(1).Item(s).Item(1).Item(count+1-i).UnRead;
        end
        
        if UnRead
            %inbox
            if isempty(f.Folder) && isempty(f.Subfolder)
                if f.Mark == 1
                INBOX.Items.Item(count+1-i).UnRead=0;
                end
                email=INBOX.Items.Item(count+1-i);
                %folder
            elseif   ~isempty(f.Folder) && isempty(f.Subfolder)
                if Mark==1
                INBOX.Folders(1).Item(n).Items(1).Item(count+1-i).UnRead=0;
                end
                email=INBOX.Folders(1).Item(n).Items(1).Item(count+1-i);
                %subfolder
            elseif ~isempty(f.Folder) && ~isempty(f.Subfolder)
                if f.Mark==1
                INBOX.Folders(1).Item(n).Folders(1).Item(s).Item(1).Item(count+1-i).UnRead=0;
                end
                email=INBOX.Folders(1).Item(n).Folders(1).Item(s).Items.Item(count+1-i);
            end
        end
    else   %****import all
        %inbox
        if isempty(f.Folder) && isempty(f.Subfolder)
            email=INBOX.Items.Item(count+1-i);
            %folder
        elseif   ~isempty(f.Folder) && isempty(f.Subfolder)
            email=INBOX.Folders(1).Item(n).Items(1).Item(count+1-i);
            %subfolder
        elseif ~isempty(f.Folder) && ~isempty(f.Subfolder)
            email=INBOX.Folders(1).Item(n).Folders(1).Item(s).Items.Item(count+1-i);
        end
        UnRead=1; %pseudo for next step
    end
    if UnRead
        % read and save body
        subject = email.get('Subject');
        body = email.get('Body');
        Email{i,1}=subject;
        Email{i,2}=body;
        if ~isempty(f.Savepath)
            attachments = email.get('Attachments');
            if attachments.Count >=1
                fname = attachments.Item(1).Filename;
                full = [f.Savepath,'\',fname];
                attachments.Item(1).SaveAsFile(full)
            end
        end
    end
end
Email(all(cellfun('isempty', Email),2),:)=[];
end

%% functions 
function f = Function_Varargin(vargs)
% varargin as structure
n = length(vargs);
if n>0
    names = vargs(1:2:n);
    values = vargs(2:2:n);
    for ix=1:(n/2)
        f.(names{ix}) = values{ix};
    end
    
    if ~isfield(f, 'Folder')
        f.Folder = '';
    end
    
    if ~isfield(f, 'Subfolder')
        f.Subfolder = '';
    end
    
    if ~isfield(f, 'Savepath')
        f.Savepath = '';
    end
    
    if ~isfield(f, 'Read')
        f.Read = '';
    end
    
    if ~isfield(f, 'Mark')
        f.Mark = '';
    end
else
    f.Folder = '';
    f.Subfolder = '';
    f.Savepath = '';
    f.Read = '';
    f.Mark = '';
end
end