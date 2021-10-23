
# Bank Management System

The Bank Management System is an application for maintaining a person's account in a bank. In this project I tried to show the working of a banking account system and cover the basic functionality of a Bank Account Management System.


## Authors

- [@SatyamAaditya](https://www.github.com/SatyamAaditya)

  
## Features

- Login Panel
- Admin Dashboard
- Customer Dashboard
### Banking Feature like:

  - Home Page
- Open Account Page
- Customer Page
- Withdraw Page
- Deposit Page
- Statement Page
- Transaction Page
- Employee Page
- Logout Page

  
## Screenshots

![Login Window](https://github.com/SatyamAaditya/Bank-Management-System/blob/Projects/Images/Screenshot%20(456).png)
![Admin Panel](https://github.com/SatyamAaditya/Bank-Management-System/blob/Projects/Images/Screenshot%20(458).jpg)
![Admin Panel](https://github.com/SatyamAaditya/Bank-Management-System/blob/Projects/Images/Screenshot%20(467).jpg)

  
## Appendix

Any additional information goes here

  
## Installation

Install my-project with npm

```bash
Step: A  Download Oracle 10g/11g 32bit/64bit and install

Step: B  Open Run SQL Command Line create a user with Username: bank  and Password:SABank
	Command to create user: 	Step:1 connect SYSTEM/password; 	(It will Connect with SYSTEM user. )
			         	Step:2 CREATE USER bank IDENTIFIED BY SABank;
				Step:3 GRANT CONNECT,RESOURCE,DBA TO bank;
				Step:4 GRANT CREATE SESSION, GRANT ANY PRIVILEGE TO bank;
				Step 5: GRANT UNLIMITED TABLESPACE TO bank;
				
Step: C  Create all the tables and its column.  [ Check SABankScript.sql  for all fields ]

Step: D GRANT SELECT, INSERT, UPDATE, DELETE on <Table Name> TO bank;  	(Provide Access To Tables)   ----[ OPTIONAL ]

Step: E Open ' Get Started With Oracle Database 11g Express Edition ' App 
			OR
 Open URL: http://127.0.0.1:8080/apex/f?p=4950

Step: F Go To Application Express

Step: G Login Username: bank Password: SABank

Step: H Now Create Application Express Workspace
		
		Database User: Use Existing
		Database Username: bank
		Application Express UserName: bank
		Password: SABank
		Confirm Password: SABank

Step: I After creating workspace Login with same creditionals

Step: J Go To ' SQL Workspace '

Step: K Go to ' SQL Script'

Step: L Go to 'Upload'

Step: M Choose File from SA Bank Daabase Folder, Filename: SABankScript.sql
	Script Name: SABankScript
	Click Upload

Step: N Now Run the Script by clicking (:) and then Click On Run Now
				
			Congratulation Your All Tables and its Data Types got Uploaded Successfully in bank user.

Step : M How To Load Admin_Login Data?
		Step: 1 After Login Application Express Workspace GoTo SQL Workshop > Utilities > Data Workshop > Data Load (XML Data) > Schema (Bank)  NEXT > Table (Admin_Login) NEXT > File (Choose File) Admin_Login.xml {Choose from SABank Database Folder) > Load Data 
			
		Congratulation Your All Data of Admin_Login got Uploaded Successfully in Admin_Login Table.

Ready to go...............
```
    
