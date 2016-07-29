
# Field2.ForeignName Property (DAO)

 **Last modified:** June 30, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


Sets or returns a value that specifies the name of the  **Field2** object in a foreign table that corresponds to a field in a primary table for a relationship (Microsoft Access workspaces only).

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ForeignName**

 _expression_ A variable that represents a **Field2** object.


## Remarks
<a name="sectionSection1"> </a>

If the  **[Relation](46d6dfaf-a97d-3abd-0b4b-396a41eb3be7.md)** object isn't appended to the **[Database](6cf2ddf8-3957-a15e-5eeb-85f81c1e415e.md)**, but the **Field2** is appended to the **Relation** object, the **ForeignName** property is read/write. Once the **Relation** object is appended to the database, the **ForeignName** property is read-only.

Only a  **Field2** object that belongs to the **Fields** collection of a **Relation** object can support the **ForeignName** property.

The  **[Name](5f4a95cd-63a3-aedf-df64-793158b2283d.md)** and **ForeignName** property settings for a **Field2** object specify the names of the corresponding fields in the primary and foreign tables of a relationship. The **[Table](cc4f64ef-c4e9-1a14-9263-5f8220d89840.md)** and **[ForeignTable](3f896433-2962-1c7c-f5a2-4e030ba8d4a0.md)** property settings for a **Relation** object determine the primary and foreign tables of a relationship.

For example, if you had a list of valid part codes (in a field named PartNo) stored in a ValidParts table, you could establish a relationship with an OrderItem table such that if a part code were entered into the OrderItem table, it would have to already exist in the ValidParts table. If the part code didn't exist in the ValidParts table and you had not set the  **[Attributes](8e6f6afb-1a89-7315-c129-cf7ff19e0ca9.md)** property of the **Relation** object to **dbRelationDontEnforce**, a trappable error would occur.

In this case, the ValidParts table is the foreign table, so the  **ForeignTable** property of the **Relation** object would be set to ValidParts and the **Table** property of the **Relation** object would be set to OrderItem. The **Name** and **ForeignName** properties of the **Field2** object in the **Relation** object's **Fields** collection would be set to PartNo.


## Example
<a name="sectionSection2"> </a>

This example shows how the  **Table**, **ForeignTable**, and **ForeignName** properties define the terms of a **Relation** between two tables.


```vb
Sub ForeignNameX() 
 
 Dim dbsNorthwind As Database 
 Dim relLoop As Relation 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
 Debug.Print "Relation" 
 Debug.Print " Table - Field" 
 Debug.Print " Primary (One) "; 
 Debug.Print ".Table - .Fields(0).Name" 
 Debug.Print " Foreign (Many) "; 
 Debug.Print ".ForeignTable - .Fields(0).ForeignName" 
 
 ' Enumerate the Relations collection of the Northwind 
 ' database to report on the property values of 
 ' the Relation objects and their Field objects. 
 For Each relLoop In dbsNorthwind.Relations 
 With relLoop 
 Debug.Print 
 Debug.Print .Name &; " Relation" 
 Debug.Print " Table - Field" 
 Debug.Print " Primary (One) "; 
 Debug.Print .Table &; " - " &; .Fields(0).Name 
 Debug.Print " Foreign (Many) "; 
 Debug.Print .ForeignTable &; " - " &; _ 
 .Fields(0).ForeignName 
 End With 
 Next relLoop 
 
 dbsNorthwind.Close 
 
End Sub
```

