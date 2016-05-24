
# Cannot update. Database or object is read-only. (Error 3027)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You tried to save changes in a database that was opened for read-only access.

The database is read-only for one of these reasons:


- You used the  **OpenDatabase** method and opened the database for read-only access.
    
- In Microsoft Visual Basic, you are using the  **Data** control, and you set the **ReadOnly** property to **True**.
    
- The database file is defined as read-only in the operating system or by your network.
    
- The database file is stored on read-only media.
    
- In a network environment, you do not have write privileges for the database file.
    
- When working with a secured database, the database or one of its objects (such as a field or table) may be set to read-only. You may not have permission to access this data with your user name and password.
    

Close the database, resolve the read-only condition, and then reopen the file for read/write access.
