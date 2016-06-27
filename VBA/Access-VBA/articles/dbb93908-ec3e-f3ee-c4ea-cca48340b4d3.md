
# Workspace.OpenDatabase Method (DAO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


Opens a specified database in a  **[Workspace](bf3ab863-5e9a-4842-1f82-2ccf958d9779.md)** object and returns a reference to the **[Database](6cf2ddf8-3957-a15e-5eeb-85f81c1e415e.md)** object that represents it.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **OpenDatabase**( ** _Name_**, ** _Options_**, ** _ReadOnly_**, ** _Connect_** )

 _expression_ A variable that represents a **Workspace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|the name of an existing Microsoft Access database engine database file, or the data source name (DSN) of an ODBC data source. See the  **[Name](5f4a95cd-63a3-aedf-df64-793158b2283d.md)** property for more information about setting this value.|
| _Options_|Optional|**Variant**|Sets various options for the database, as specified in Remarks.|
| _ReadOnly_|Optional|**Variant**|**True** if you want to open the database with read-only access, or **False** (default) if you want to open the database with read/write access.|
| _Connect_|Optional|**Variant**|Specifies various connection information, including passwords.|

### Return Value

Database


## Remarks
<a name="sectionSection1"> </a>

You can use the following values for the  _options_ argument.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|Opens the database in exclusive mode.|
|**False**|(Default) Opens the database in shared mode.|
When you open a database, it is automatically added to the  **Databases** collection.

Some considerations apply when you use  _dbname_:




- If it refers to a database that is already open for access by another user, an error occurs.
    
- If it doesn't refer to an existing database or valid ODBC data source name, an error occurs.
    
- If it's a zero-length string ("") and  _connect_ is `"ODBC;"`, a dialog box listing all registered ODBC data source names is displayed so the user can select a database.
    


To close a database, and thus remove the  **Database** object from the **Databases** collection, use the **[Close](9b1a77cb-da12-24d6-892f-a56be103d51d.md)** method on the object.




 **Note**  When you access a Microsoft access database engine-connected ODBC data source, you can improve your application's performance by opening a  **Database** object connected to the ODBC data source, rather than by linking individual **[TableDef](715146b6-c62a-abff-28ee-e6bbe3c08adf.md)** objects to specific tables in the ODBC data source.


## Example
<a name="sectionSection2"> </a>

This example uses the  **OpenDatabase** method to open one Microsoft Access database and two Microsoft Access database engine-connected ODBC databases.


```vb
Sub OpenDatabaseX() 
 
 Dim wrkAcc As Workspace 
 Dim dbsNorthwind As Database 
 Dim dbsPubs As Database 
 Dim dbsPubs2 As Database 
 Dim dbsLoop As Database 
 Dim prpLoop As Property 
 
 ' Create Microsoft Access Workspace object. 
 Set wrkAcc = CreateWorkspace("", "admin", "", dbUseJet) 
 
 ' Open Database object from saved Microsoft Access database 
 ' for exclusive use. 
 MsgBox "Opening Northwind..." 
 Set dbsNorthwind = wrkAcc.OpenDatabase("Northwind.mdb", _ 
 True) 
 
 ' Open read-only Database object based on information in 
 ' the connect string. 
 MsgBox "Opening pubs..." 
 
 ' Note: The DSN referenced below must be set to 
 ' use Microsoft Windows NT Authentication Mode to 
 ' authorize user access to the Microsoft SQL Server. 
 Set dbsPubs = wrkAcc.OpenDatabase("Publishers", _ 
 dbDriverNoPrompt, True, _ 
 "ODBC;DATABASE=pubs;DSN=Publishers") 
 
 ' Open read-only Database object by entering only the 
 ' missing information in the ODBC Driver Manager dialog 
 ' box. 
 MsgBox "Opening second copy of pubs..." 
 Set dbsPubs2 = wrkAcc.OpenDatabase("Publishers", _ 
 dbDriverCompleteRequired, True, _ 
 "ODBC;DATABASE=pubs;DSN=Publishers;") 
 
 ' Enumerate the Databases collection. 
 For Each dbsLoop In wrkAcc.Databases 
 Debug.Print "Database properties for " &; _ 
 dbsLoop.Name &; ":" 
 
 On Error Resume Next 
 ' Enumerate the Properties collection of each Database 
 ' object. 
 For Each prpLoop In dbsLoop.Properties 
 If prpLoop.Name = "Connection" Then 
 ' Property actually returns a Connection object. 
 Debug.Print " Connection[.Name] = " &; _ 
 dbsLoop.Connection.Name 
 Else 
 Debug.Print " " &; prpLoop.Name &; " = " &; _ 
 prpLoop 
 End If 
 Next prpLoop 
 On Error GoTo 0 
 
 Next dbsLoop 
 
 dbsNorthwind.Close 
 dbsPubs.Close 
 dbsPubs2.Close 
 wrkAcc.Close 
 
End Sub
```

