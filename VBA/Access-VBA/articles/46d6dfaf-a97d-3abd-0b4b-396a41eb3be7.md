
# Relation Object (DAO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

A  **Relation** object represents a relationship between fields in tables or queries (Microsoft Access database engine databases only).


## Remarks

You can use the  **Relation** object to create new relationships and examine existing relationships in your database.

Using a  **Relation** object and its properties, you can:




- Specify an enforced relationship between fields in base tables (but not a relationship that involves a query or a linked table).
    
- Establish unenforced relationships between any type of table or query— native or linked.
    
- Use the  **Name** property to refer to the relationship between the fields in the referenced primary table and the referencing foreign table.
    
- Use the  **Attributes** property to determine whether the relationship between fields in the table is one-to-one or one-to-many and how to enforce referential integrity.
    
- Use the  **Attributes** property to determine whether the Microsoft Access database engine can perform cascading update and cascading delete operations on primary and foreign tables.
    
- Use the  **Attributes** property to determine whether the relationship between fields in the table is left join or right join.
    
- Use the  **Name** property of all **Field** objects in the **Fields** collection of a **Relation** object to set or return the names of the fields in the primary key of the referenced table, or the **ForeignName** property settings of the **Field** objects to set or return the names of the fields in the foreign key of the referencing table.
    


If you make changes that violate the relationships established for the database, a trappable error occurs. If you request cascading update or cascading delete operations, the Microsoft Access database engine also modifies the primary key or foreign key tables to enforce the relationships you establish.

For example, the Northwind database contains a relationship between an Orders table and a Customers table. The CustomerID field of the Customers table is the primary key, and the CustomerID field of the Orders table is the foreign key. For the Microsoft Access database engine to accept a new record in the Orders table, it searches the Customers table for a match on the CustomerID field of the Orders table. If the Microsoft Access database engine doesn't find a match, it doesn't accept the new record, and a trappable error occurs.

When you enforce referential integrity, a unique index must already exist for the key field of the referenced table. The Microsoft Access database engine automatically creates an index with the  **Foreign** property set to act as the foreign key in the referencing table.

To create a new  **Relation** object, use the **CreateRelation** method. To refer to a **Relation** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:

 **Relations** (0)

 **Relations** (" _name_")

 **Relations** ![ _name_]


## Example

This example shows how an existing  **Relation** object can control data entry. The procedure attempts to add a record with a deliberately incorrect CategoryID; this triggers the error-handling routine.


```vb
Sub RelationX() 
 
 Dim dbsNorthwind As Database 
 Dim rstProducts As Recordset 
 Dim prpLoop As Property 
 Dim fldLoop As Field 
 Dim errLoop As Error 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 Set rstProducts = dbsNorthwind.OpenRecordset("Products") 
 
 ' Print a report showing all the different parts of 
 ' the relation and where each part is stored. 
 With dbsNorthwind.Relations!CategoriesProducts 
 Debug.Print "Properties of " &; .Name &; " Relation" 
 Debug.Print " Table = " &; .Table 
 Debug.Print " ForeignTable = " &; .ForeignTable 
 Debug.Print "Fields of " &; .Name &; " Relation" 
 With .Fields!CategoryID 
 Debug.Print " " &; .Name 
 Debug.Print " Name = " &; .Name 
 Debug.Print " ForeignName = " &; .ForeignName 
 End With 
 End With 
 
 ' Attempt to add a record that violates the relation. 
 With rstProducts 
 .AddNew 
 !ProductName = "Trygve's Lutefisk" 
 !CategoryID = 10 
 On Error GoTo Err_Relation 
 .Update 
 On Error GoTo 0 
 .Close 
 End With 
 
 dbsNorthwind.Close 
 
 Exit Sub 
 
Err_Relation: 
 
 ' Notify user of any errors that result from 
 ' the invalid data. 
 If DBEngine.Errors.Count > 0 Then 
 For Each errLoop In DBEngine.Errors 
 MsgBox "Error number: " &; errLoop.Number &; _ 
 vbCr &; errLoop.Description 
 Next errLoop 
 End If 
 
 Resume Next 
 
End Sub
```

This example uses the  **CreateRelation** method to create a **Relation** between the Employees **TableDef** and a new **TableDef** called Departments. This example also demonstrates how creating a new **Relation** will also create any necessary **Indexes** in the foreign table (the DepartmentsEmployees Index in the Employees table).




```vb
Sub CreateRelationX() 
 
 Dim dbsNorthwind As Database 
 Dim tdfEmployees As TableDef 
 Dim tdfNew As TableDef 
 Dim idxNew As Index 
 Dim relNew As Relation 
 Dim idxLoop As Index 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
 With dbsNorthwind 
 ' Add new field to Employees table. 
 Set tdfEmployees = .TableDefs!Employees 
 tdfEmployees.Fields.Append _ 
 tdfEmployees.CreateField("DeptID", dbInteger, 2) 
 
 ' Create new Departments table. 
 Set tdfNew = .CreateTableDef("Departments") 
 
 With tdfNew 
 ' Create and append Field objects to Fields 
 ' collection of the new TableDef object. 
 .Fields.Append .CreateField("DeptID", dbInteger, 2) 
 .Fields.Append .CreateField("DeptName", dbText, 20) 
 
 ' Create Index object for Departments table. 
 Set idxNew = .CreateIndex("DeptIDIndex") 
 ' Create and append Field object to Fields 
 ' collection of the new Index object. 
 idxNew.Fields.Append idxNew.CreateField("DeptID") 
 ' The index in the primary table must be Unique in 
 ' order to be part of a Relation. 
 idxNew.Unique = True 
 .Indexes.Append idxNew 
 End With 
 
 .TableDefs.Append tdfNew 
 
 ' Create EmployeesDepartments Relation object, using 
 ' the names of the two tables in the relation. 
 Set relNew = .CreateRelation("EmployeesDepartments", _ 
 tdfNew.Name, tdfEmployees.Name, _ 
 dbRelationUpdateCascade) 
 
 ' Create Field object for the Fields collection of the 
 ' new Relation object. Set the Name and ForeignName 
 ' properties based on the fields to be used for the 
 ' relation. 
 relNew.Fields.Append relNew.CreateField("DeptID") 
 relNew.Fields!DeptID.ForeignName = "DeptID" 
 .Relations.Append relNew 
 
 ' Print report. 
 Debug.Print "Properties of " &; relNew.Name &; _ 
 " Relation" 
 Debug.Print " Table = " &; relNew.Table 
 Debug.Print " ForeignTable = " &; _ 
 relNew.ForeignTable 
 Debug.Print "Fields of " &; relNew.Name &; " Relation" 
 
 With relNew.Fields!DeptID 
 Debug.Print " " &; .Name 
 Debug.Print " Name = " &; .Name 
 Debug.Print " ForeignName = " &; .ForeignName 
 End With 
 
 Debug.Print "Indexes in " &; tdfEmployees.Name &; _ 
 " TableDef" 
 For Each idxLoop In tdfEmployees.Indexes 
 Debug.Print " " &; idxLoop.Name &; _ 
 ", Foreign = " &; idxLoop.Foreign 
 Next idxLoop 
 
 ' Delete new objects because this is a demonstration. 
 .Relations.Delete relNew.Name 
 .TableDefs.Delete tdfNew.Name 
 tdfEmployees.Fields.Delete "DeptID" 
 .Close 
 End With 
 
End Sub
```

