# RPA-Project-Email-Automation
Filter emails from inbox based on sender's email addresses or email subject and store emails and attachments in seperate folders according to sender.

### Table of Content
  * [Overview](#overview)
  * [Motivation](#motivation)
  * [Steps in the project execution](#steps-in-the-project-execution)
  * [Technical Aspect](#technical-aspect)
  * [Installation](#installation)

### Overview 

### Steps in the project execution

1. Create a Sequence and add 'Get IMAP Mail Messages'. Then create email, password and mailMessages variables. Set your email and password for accessing your inbox. Set mailMessages type as List<System.Net.Mail.MailMessage>.

2. Now in Properties, set the Port as 993 and Server as "imap.gmail.com". Also enter email and password into Logon section.

3. Set the count of mails you want to filter in Top property. Uncheck OnlyUnreadMessages option to filter emails from already read mails.

4. Now in sequence, drag ForEach activity and iterate over mailMessages list. Inside foreach, add If activity to check each mail's sender email address using mail.From.Address.Contains(""). If the condition holds true, do whatever task you wish to do. For false condition, also define a task.

### Technical Aspect
### Email Automation with RPA Using UiPath
Among numerous activities that UiPath can automate, Email automation is one of the most popular requirements for many employees and organizations across the globe. Using Uipath we can automate sending emails as well as receiving them. In this tutorial, we will explain all the activities and packages that UiPath offers for Email automation.

<img width="300" alt="1" src="https://user-images.githubusercontent.com/122998236/229130999-fde7de28-265c-4db3-9690-005e32f7bc54.png">

Simply searching for mail in the Activities Panel gives out all the required activities and opportunities that UiPath provides.

SMTP – It is short for Simple Mail Transfer Protocol and is a basic protocol that only allows sending messages
POP3 – It stands for Post Office Protocol and is an older version of a protocol for reading messages. Although this protocol is almost obsolete, most mail servers supports it.
IMAP - Internet Message Access Protocol only helps to receive messages and also provides additional features such as “Read” and “Move” emails between folders.
Exchange - This is a Microsoft Enterprise email solution which is integrated within UiPath flawlessly and allows the user to perform multiple actions such as Send, Delete, Move, and Get Email messages.
Outlook – This activity uses the API of the desktop application and hence does not require to set up users, servers, or other advanced details. This activity integrates Outlook API such that UiPath can use the already logged-in Outlook account.
IBM Notes – This uses the IBM notes desktop API to interact with the application and perform send, receive, move, and reading activities. Similar to Outlook, this does not require the user to set up servers, users, and works with the logged-in account.
Save Mail Message – Saves Emails on local disk
Save Attachments – Downloads attachments and saves them to local disk

### Checking Emails
There are four activities in UiPath that allow checking received Emails which include Get IMAP mail messages, Get POP3 Mail Messages, Get Outlook Mail Messages, And Get Exchange Mail Messages, and Finally Get IBM Notes Mail Messages.

<img width="300" alt="2" src="https://user-images.githubusercontent.com/122998236/229132848-7b4d887b-bac1-4bbe-aee4-5b18f4771b35.png">

All of them provide the same functionality of receiving or providing email messages but each one works for their own environment or application. Naturally, they share the same Properties as well.
<img width="300" alt="3" src="https://user-images.githubusercontent.com/122998236/229132938-94c9cfc2-717b-45e0-b083-a3c3bb2e92ce.png">


For example, the MailFolder from where the Emails will be fetched. By default, it is set to Inbox for all Get Mail activities.
<img width="300" alt="4" src="https://user-images.githubusercontent.com/122998236/229132969-5f3cc86a-8217-4f8f-b6e7-36f39ce06c21.png">


Helps to fetch only Unread Messages and works perfectly with MarkAsRead, which eliminates reprocessing and getting the same Emails every time.
Next, we have the Top property which limits the number of incoming Emails and has a value of 30 by default.

<img width="300" alt="5" src="https://user-images.githubusercontent.com/122998236/229133193-ea6e71eb-bb4e-4ea6-8d07-62be4689af1d.png">

This property helps to set up a Server connection and requires the input of the server address and port along with the User name and Password.
 
<img width="300" alt="6" src="https://user-images.githubusercontent.com/122998236/229133276-141695b4-eaba-41dc-a2e9-59c417dcc853.png">
<img width="300" alt="7" src="https://user-images.githubusercontent.com/122998236/229133500-e6346501-d51e-4259-b462-bc62f6dc89a6.png">

It is important to note that this Server, User name, and Password settings are only necessary for IMAP and POP3 and hence they show validation errors when these fields are empty.
Hence, automating Outlook, IBM Notes or Exchange requires much less effort and also has an additional feature of EmailAutoDiscover that searches for required parameters without any user input.<br>
<img width="300" alt="8" src="https://user-images.githubusercontent.com/122998236/229133534-8fc92e1a-954f-424e-af9f-b48f3055f7f7.png">

### Using IMAP to Read Emails
To read messages using IMAP we need to set up a server and port and also input Username and Password.<br>
<img width="300" alt="9" src="https://user-images.githubusercontent.com/122998236/229133720-7cc90d7c-fe83-44ad-aaf5-dcf2c2d3a604.png">

In this tutorial, we will simply use a Gmail account which has a server of "imap.google.com" with Port 993.
Inside Email, we have supplied a variable containing an email address, and similarly, for the password, a variable has been supplied.
This information can be assigned directly to a variable using Assign or Get Password or the user may also Get Credentials to collect them from Windows.<br>
<img width="300" alt="10" src="https://user-images.githubusercontent.com/122998236/229133844-a3dbd839-c217-4d56-b2e6-eef7164d5078.png">

We only want to read the first 4 Emails that are marked Unread. The final output is saved inside a Mail List variable called message.<br>

<img width="300" alt="11" src="https://user-images.githubusercontent.com/122998236/229133935-232a00d5-836a-4018-8537-40b8f9ab78f5.png">

Now we will use the For Each activity to create a loop for all the Emails that are read.<br>

<img width="300" alt="12" src="https://user-images.githubusercontent.com/122998236/229134024-d88bf01f-85ce-4b61-8b96-a6acb578fe45.png">

We also need to the TypeArgument to System.Net.Mail.MailMessage type variable using the Browse option manually. Also, instead of For Each item, we will use For Each mail, which is necessary to identify all Emails that were read.<br>

<img width="300" alt="13" src="https://user-images.githubusercontent.com/122998236/229134068-5332f96a-8bfb-4756-aa26-4653b2849e16.png">

Next, we will use a Message Box to display the Subject of all the first 4 unread Emails using the syntax mail.Subject.<br>
<img width="300" alt="14" src="https://user-images.githubusercontent.com/122998236/229134081-8b8b1a86-962e-4484-b914-6db8577df850.png">


It is important to note that, putting a dot after mail allows you to select a wide range of Email properties as per requirements.
Executing this workflow will read the Subject of the first 4 unread Emails one by one.

### Filtering Emails and saving attachments
Suppose a user needs the UiPath Robot to check unread messages and download attachments only from the ones that contain the subject line "Employee Details".<br>

<img width="301" alt="15" src="https://user-images.githubusercontent.com/122998236/229134229-72ef8cae-8df2-41a0-ada5-b117b8e92f75.png">

To do this we will again use the same For Each mail loop, but include an If as well as Save Attachments activity as shown in the image.<br>

### Sending Emails
<img width="300" alt="16" src="https://user-images.githubusercontent.com/122998236/229134254-336a48c6-91dd-48b8-ad9f-026b5837796a.png">


For this activity, we will the Send SMTP Mail Message activity and much similar to IMAP, this requires setting up Port, Server, Username, and Password.<br>
<img width="245" alt="17" src="https://user-images.githubusercontent.com/122998236/229134272-a5e478dc-c37d-471b-8fc9-3e2ec15599c0.png">


Next, we need to fill in the Subject, Body, and recipient's address as shown in the image. The user is allowed to send files as well using the Attach Files option.<br>
<img width="221" alt="18" src="https://user-images.githubusercontent.com/122998236/229134292-d5b6247d-74ec-40e7-b231-c7f1ff133c40.png">
