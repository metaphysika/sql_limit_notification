# sql_limit_notification

This program takes 2 arguments, one for an exam type and another for a ctdivol threshold value that will trigger an alert.

The progam first connects to a database and creates a pandas dataframe from one of the tables.

## Finding values above set limit

It then finds the exam type specified and compares the actual ctdivol with the specified limit.  If the exam is above the limit, the program will open a file containing a record of all the previous limits.  If the record is already there, it moves on.  If the record is not there, the program then gets some information about that exam from several different tables in the database.

## Sending the alert

Once the program has all the necessary data, it will send an email and append the data to the file containing a record of all the the previous alerts.
