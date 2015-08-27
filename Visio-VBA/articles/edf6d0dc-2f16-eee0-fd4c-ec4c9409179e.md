
# DataRecordsets Object (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

The collection of  **DataRecordset** objects associated with a **Document** object.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **DataRecordsets** collection is ** [Item](8a289fb1-8cc5-eb76-efb1-c01f73c6340a.md)**.

Every Visio  **Document** object has a **DataRecordsets** collection, which is empty until you import data into Visio. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document.

To add a  **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the type of data source you want to connect to (OLEDB/ODBC or XML) and how you want to pass connection string and query command strings to Visio. By using the




-  ** [DataRecordsets.Add](9eb136ce-d543-75c3-3a72-cb23dfc8df78.md)** method, you can connect to an OLEDB or ODBC data source and pass connection and query command string information to Visio directly as method parameters.
    
-  ** [DataRecordsets.AddFromConnectionFile](7118bd4d-484b-dc22-e6f8-925376a5a67a.md)** method, you can connect to an OLEBD or ODBC data source by passing the method an Office Data Connection (ODC) file that contains the connection and query command string information you want to supply to Visio.
    
-  ** [DataRecordsets.AddFromXML](b75d7ecc-98d2-ae9b-608f-a9ec2b736ea6.md)** method, you pass the method an ADO classic XML string that contains all the data that you want to include in the data recordset.
    


Once you have created a data recordset, the connection string and query command string associated with the data recordset are represented by the  ** [DataConnection.ConnectionString](a1a6105f-64ee-1e0c-3b54-9831aec06bf4.md)** and ** [DataRecordset.CommandString](7d9151b0-db8c-a8ce-edea-7ef25d241e98.md)** properties respectively.

