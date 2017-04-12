# Automatic-recruitment-system
This system is aimed for big organizations that get tons of job applications daily and dont even meet the job description.

This is a complete recruitment system. It takes your email id and password as input and searches for all the unseen emails with subject "resume"(any keyword can be used. for ex: job,cv. It can also take date or search before kind of parameters.). Next it asks for keywords to search in the resume(skillset).

It will download the attachments for emails with subject resume in a specified folder. The attachment can be a pdf or docx.The resumes will be saved with file name as email id to avoid overwriting common names.

The resumes will then be scanned to look for the specified skillset.(it can be modified for different uses.)How to evaluate resume varies for different organizations.

The email id , name , phone number, date will be put in an excel sheet with a decision 'yes' or 'no' against them.

The candidates will be sent an automated reply based on the decision in excel sheet.
