
# Field.Required Property (DAO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


Sets or returns a value that indicates whether a  **[Field](47282ce2-9b49-ccf9-ad37-c4bb25cfd037.md)** object requires a non-Null value.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Required**

 _expression_ A variable that represents a **Field** object.


## Remarks
<a name="sectionSection1"> </a>

For a  **Field** not yet appended to the **Fields** collection, this property is read/write.

The availability of the  **Required** property depends on the object that contains the[Fields](4be3ba07-20c1-d958-c1b8-7dd8b4731f60.md) collection, as shown in the following table.



|**If the Fields collection belongs to a**|**Then Required is**|
|:-----|:-----|
|**Index** object|Not supported|
|**QueryDef** object|Read-only|
|**Recordset** object|Read-only|
|**Relation** object|Not supported|
|**TableDef** object|Read/write|
You can use the  **Required** property along with the **[AllowZeroLength](5103a905-9258-e088-0210-857372f41c3c.md)**, **[ValidateOnSet](00245a8a-a78f-b0a8-3eb3-11dd27873984.md)**, or **[ValidationRule](b07e644d-54d3-7199-6f99-178774e54398.md)** property to determine the validity of the **[Value](6c0f9a8d-f51a-b8cf-8830-f8d960a1d08c.md)** property setting for that **Field** object. If the **Required** property is set to **False**, the field can contain **null** values as well as values that meet the conditions specified by the **AllowZeroLength** and **ValidationRule** property settings.




 **Note**  When you can set this property for either an  **Index** object or a **Field** object, set it for the **Field** object. The validity of the property setting for a **Field** object is checked before that of an **Index** object.


## Example
<a name="sectionSection2"> </a>

This example uses the  **Required** property to report which fields in three different tables must contain data in order for a new record to be added. The RequiredOutput procedure is required for this procedure to run.


```vb
Sub RequiredX() 
 
 Dim dbsNorthwind As Database 
 Dim tdfloop As TableDef 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
 With dbsNorthwind 
 ' Show which fields are required in the Fields 
 ' collections of three different TableDef objects. 
 RequiredOutput .TableDefs("Categories") 
 RequiredOutput .TableDefs("Customers") 
 RequiredOutput .TableDefs("Employees") 
 .Close 
 End With 
 
End Sub 
 
Sub RequiredOutput(tdfTemp As TableDef) 
 
 Dim fldLoop As Field 
 
 ' Enumerate Fields collection of the specified TableDef 
 ' and show the Required property. 
 Debug.Print "Fields in " &; tdfTemp.Name &; ":" 
 For Each fldLoop In tdfTemp.Fields 
 Debug.Print , fldLoop.Name &; ", Required = " &; _ 
 fldLoop.Required 
 Next fldLoop 
 
End Sub
```

