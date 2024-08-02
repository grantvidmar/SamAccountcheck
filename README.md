# SamAccountcheck
#This project was created to connect to Microsoft.Graph and get some infromation that we needed for creating new users in a program called Everyfile. We needed the below information to help create the user account.

#Display name
#SAM Account
#ObjectID
#Department
#UPN
#Email

#To preform this search we already had the UPN for the users we jsut needed the remaining information from AzureAD. So the script looks at an already generated CSV file and searched the individual UPN names and fills in the remaining information. This came in handy when we are searching for over 200 users information. V1 of this script can be used for individual users and will not output to a CSV file. Both scripts could be updated to add what ever information you may need to search for, you would just need to update the columns to reflect the new data requested. 
