
# ODBCConnectStr Property

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **ODBCConnectStr** property in an SQL pass-through query to specify the Open Database Connectivity (ODBC) connection string for the query.


 **Note**  The  **ODBCConnectStr** property applies only to pass-through queries.


## Setting

Enter the ODBC connection string that defines the connection to the SQL database you want to use.

You can set this property by using the query's property sheet or Visual Basic .

You can also use the ODBC Connection String Builder to create the ODBC connection string for this property. This builder establishes a connection to the SQL database server and then ends the connection after the ODBC connection string is created.


## Remarks

The default setting for this property is the string "ODBC;", which Microsoft Access restores if you delete an existing setting. When the  **ODBCConnectStr** property is set to "ODBC;", Microsoft Access will prompt you for a connection string whenever the query is used but won't store the connect string. You must enter a connection string in the **ODBCConnectStr** property box if you want Microsoft Access to store the connection string.


 **Tip**  If you know the full connection string for the SQL database, enter it in the  **ODBCConnectStr** property box. This way you will avoid the need to enter the connection string in the ODBC connection dialog box each time you use the query.

The connection string is different for different types of ODBC data sources. For example, to connect to the Human Resources data source on the HRSRVR server (a Microsoft SQL Server) by using the logon identification (ID) Smith and the password Sesame, you can use the following connection string.




```
ODBC;DSN=Human Resources;SERVER=HRSRVR;UID=Smith;PWD=Sesame;
```

