
# TableDef Object (DAO)

 **Last modified:** July 01, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Remarks](#sectionSection0)
[Example](#sectionSection1)
[About the Contributors](#AboutContributors)


A  **TableDef** object represents the stored definition of a base table or a linked table (Microsoft Access workspaces only).

## Remarks
<a name="sectionSection0"> </a>

You manipulate a table definition using a  **TableDef** object and its methods and properties. For example, you can:




- Examine the field and index structure of any local, linked, or external table in a database.
    
- Use the  **Connect** and **SourceTableName** properties to set or return information about linked tables, and use the **RefreshLink** method to update connections to linked tables.
    
- Use the  **ValidationRule** and **ValidationText** properties to set or return validation conditions.
    
- Use the  **OpenRecordset** method to create a table-, dynaset-, dynamic-, snapshot-, or forward-only-type **Recordset** object, based on the table definition.
    


For base tables, the  **RecordCount** property contains the number of records in the specified database table. For linked tables, the **RecordCount** property setting is always -1.

To create a new  **TableDef** object, use the **[CreateTableDef](d919b44e-ffae-dc4a-f1cc-d01df49987a3.md)** method.


### To add a field to a table




1. Make sure any  **[Recordset](9774232c-e6da-175b-fc7f-ed2ab7908fa0.md)** objects based on the table are all closed.
    
2. Use the  **CreateField** method to create a **Field** object variable and set its properties.
    
3. Use the  **Append** method to add the **Field** object to the **Fields** collection of the **TableDef** object.
    
You can delete a  **Field** object from a **TableDefs** collection if it doesn't have any indexes assigned to it, but you will lose the field's data.


### To create a table that is ready for new records in a database




1. Use the  **CreateTableDef** method to create a **TableDef** object.
    
2. Set its properties.
    
3. For each field in the table, use the  **CreateField** method to create a **Field** object variable and set its properties.
    
4. Use the  **Append** method to add the fields to the **Fields** collection of the **TableDef** object.
    
5. Use the  **Append** method to add the new **TableDef** object to the **TableDefs** collection of the **Database** object.
    
A linked table is connected to the database by the  **SourceTableName** and **Connect** properties of the **TableDef** object.


### To link a table to a database




1. Use the  **CreateTableDef** method to create a **TableDef** object.
    
2. Set its  **Connect** and **SourceTableName** properties (and optionally, its **Attributes** property).
    
3. Use the  **Append** method to add it to the **TableDefs** collection of a **Database**.
    
To refer to a  **TableDef** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:

 **TableDefs** (0)

 **TableDefs** (" _name_")

 **TableDefs** ![ _name_]


## Example
<a name="sectionSection1"> </a>

This example creates a new  **TableDef** object and appends it to the **TableDefs** collection of the Northwind Database object. It then enumerates the **TableDefs** collection and the **Properties** collection of the new **TableDef**.


```vb
Sub TableDefX() 
 
   Dim dbsNorthwind As Database 
   Dim tdfNew As TableDef 
   Dim tdfLoop As TableDef 
   Dim prpLoop As Property 
 
   Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
   ' Create new TableDef object, append Field objects  
   ' to its Fields collection, and append TableDef  
   ' object to the TableDefs collection of the  
   ' Database object. 
   Set tdfNew = dbsNorthwind.CreateTableDef("NewTableDef") 
   tdfNew.Fields.Append tdfNew.CreateField("Date", dbDate) 
   dbsNorthwind.TableDefs.Append tdfNew 
 
   With dbsNorthwind 
      Debug.Print .TableDefs.Count &; _ 
         " TableDefs in " &; .Name 
 
      ' Enumerate TableDefs collection. 
      For Each tdfLoop In .TableDefs 
         Debug.Print "  " &; tdfLoop.Name 
      Next tdfLoop 
 
      With tdfNew 
         Debug.Print "Properties of " &; .Name 
 
         ' Enumerate Properties collection of new 
         ' TableDef object, only printing properties 
         ' with non-empty values. 
         For Each prpLoop In .Properties 
            Debug.Print "  " &; prpLoop.Name &; " - " &; _ 
               IIf(prpLoop = "", "[empty]", prpLoop) 
         Next prpLoop 
 
      End With 
 
      ' Delete new TableDef since this is a  
      ' demonstration. 
      .TableDefs.Delete tdfNew.Name 
      .Close 
   End With 
 
End Sub
```

This example creates a new  **TableDef** object in the Northwind database.




```vb
Sub CreateTableDefX() 
 
   Dim dbsNorthwind As Database 
   Dim tdfNew As TableDef 
   Dim prpLoop As Property 
 
   Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
   ' Create a new TableDef object. 
   Set tdfNew = dbsNorthwind.CreateTableDef("Contacts") 
 
   With tdfNew 
      ' Create fields and append them to the new TableDef  
      ' object. This must be done before appending the  
      ' TableDef object to the TableDefs collection of the  
      ' Northwind database. 
      .Fields.Append .CreateField("FirstName", dbText) 
      .Fields.Append .CreateField("LastName", dbText) 
      .Fields.Append .CreateField("Phone", dbText) 
      .Fields.Append .CreateField("Notes", dbMemo) 
 
      Debug.Print "Properties of new TableDef object " &; _ 
         "before appending to collection:" 
 
      ' Enumerate Properties collection of new TableDef  
      ' object. 
      For Each prpLoop In .Properties 
         On Error Resume Next 
         If prpLoop <> "" Then Debug.Print "  " &; _ 
           prpLoop.Name &; " = " &; prpLoop 
         On Error GoTo 0 
      Next prpLoop 
 
      ' Append the new TableDef object to the Northwind  
      ' database. 
      dbsNorthwind.TableDefs.Append tdfNew 
 
      Debug.Print "Properties of new TableDef object " &; _ 
         "after appending to collection:" 
 
      ' Enumerate Properties collection of new TableDef  
      ' object. 
      For Each prpLoop In .Properties 
         On Error Resume Next 
         If prpLoop <> "" Then Debug.Print "  " &; _ 
           prpLoop.Name &; " = " &; prpLoop 
         On Error GoTo 0 
      Next prpLoop 
 
   End With 
 
   ' Delete new TableDef object since this is a  
   ' demonstration. 
   dbsNorthwind.TableDefs.Delete "Contacts" 
 
   dbsNorthwind.Close 
 

```

The following example shows how to create a calculated field. The  **CreateField** method creates a field named **FullName**. The  **Expression** property is then set to the expression that calculates the value of the field.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.mdl) |[About the Contributors](#AboutContributors)




```vb
Sub CreateCalculatedField()
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field2
    
    ' get the database
    Set dbs = CurrentDb()
    
    ' create the table
    Set tdf = dbs.CreateTableDef("tblContactsCalcField")
    
    ' create the fields: first name, last name
    tdf.Fields.Append tdf.CreateField("FirstName", dbText, 20)
    tdf.Fields.Append tdf.CreateField("LastName", dbText, 20)
    
    ' create the calculated field: full name
    Set fld = tdf.CreateField("FullName", dbText, 50)
    fld.Expression = "[FirstName] &; "" "" &; [LastName]"
    tdf.Fields.Append fld
    
    ' append the table and cleanup
    dbs.TableDefs.Append tdf
    
Cleanup:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 

