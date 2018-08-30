# PYTHON-Resume-Management-System
This project can potentially be used as a resume management tool. It can maintain records of the resumes, applicants email id, etc.

It takes the company's email id and password as input and searches for all the unseen emails with subject "resume"(any keyword can be used. for ex: job,cv. It can also take date or search before kind of parameters.). Next it asks for keywords to search in the resume(skillset).

This program can run as a simple python script `main.py`.Before running the program go to 
gmail->my account->sign in and scurity->connected apps and site. Set allow less secure apps to ON.

It will download the attachments for emails with subject resume in a specified folder. The attachment can be a .pdf or .docx.The resumes will be saved with file name as email id to avoid overwriting common names.

The resumes will then be scanned to look for the specified skillset.(it can be modified for different uses.)How to evaluate resume varies for different organizations.

The system will put email id , name , phone number, date  in an excel sheet with a decision 'yes' or 'no' against them.
As for the decision column, the program asks for the keywords to look for in the resume and if all the keywords are present the decision will be yes else no. 
The candidates will be sent an automated reply based on the decision in excel sheet.

Every time the program is run a new directory is created named as current date-time in the directory resumes-and-candidate-data and all the resumes as well as the excel sheet are put in that directory.This is done because each time a resume is downloaded from email its status is changed from seen to unseen and all this resumes are put in the specific date-time folder and this avoids duplication.



You will manually have to create a directory 'resumes-and-candidate-data' in the directory where the program is saved.
Some snapshots:

1) Every time the program is ran ,a new folder is created in the format shown below and each folder will have corresponding unseen resumes and the excel sheet with the candidate data and decisions.
 
![Alt text](https://github.com/kunalchhabria/PYTHON-Automated-recruitment-system/blob/master/python%20auto%20recruitment%20pics/1.png "1")  

2) When the program is run it initially asks for the input and then this kind of display is shown.The location of the the corresponding folder is shown, The resumes, number of words in resumes(just for debugging), and the enails to which replies are sent.


![Alt text](https://github.com/kunalchhabria/PYTHON-Automated-recruitment-system/blob/master/python%20auto%20recruitment%20pics/2.png "2")
 
3) The next photo shows the contents of folder created above.
 
![Alt text](https://github.com/kunalchhabria/PYTHON-Automated-recruitment-system/blob/master/python%20auto%20recruitment%20pics/3.png "3") 

4) This is a sample of the excel sheet that the system will create.
 
![Alt text](https://github.com/kunalchhabria/PYTHON-Automated-recruitment-system/blob/master/python%20auto%20recruitment%20pics/4.png "4")

5) This is an example of autmomated reply. The name of candidate will be taken from thte excel sheet.
 

![Alt text](https://github.com/kunalchhabria/PYTHON-Automated-recruitment-system/blob/master/python%20auto%20recruitment%20pics/5.png "5")
