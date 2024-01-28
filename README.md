# Walmart_Tracking
Finds Walmart trackings from email

Make sure to put all these files into the same folder

You can change the credentials for the Gmail in credentials.txt

These credentials are from Gmail app passwords. You can get it from here: https://myaccount.google.com/security?hl=en

You need to enable 2FA, then go all the way to App Passwords and generate the password for the app, the name can be anything

Put your email and app password in credentials.txt

You can change the date it searches from by changing the values on the main file

It can be changed in this line: start_date = datetime.date(2024, 1, 1)  # Year, Month, Day

The folder where you save this script to will contain the .csv file that contains all the trackings

If you get an error as 'e' then that means it did not find any trackings and/or there is something wrong
