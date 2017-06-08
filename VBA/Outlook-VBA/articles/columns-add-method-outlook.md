---
title: Columns.Add Method (Outlook)
keywords: vbaol11.chm2741
f1_keywords:
- vbaol11.chm2741
ms.prod: outlook
api_name:
- Outlook.Columns.Add
ms.assetid: d438cfeb-629f-4234-6f4f-ffa086ef9a41
ms.date: 06/08/2017
---


# Columns.Add Method (Outlook)

Adds the  **[Column](column-object-outlook.md)** specified by _Name_ to the **[Columns](columns-object-outlook.md)** collection and resets the **[Table](table-object-outlook.md)** .


## Syntax

 _expression_ . **Add**( **_Name_** )

 _expression_ A variable that represents a **Columns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that is being added as a column.|

### Return Value

A  **Column** object that represents the new column.


## Remarks

 **Columns.Add** adds the specified **Column** to the end of the **Columns** collection for the **Table** , and resets the **Table** by moving the current row to just before the first row of the **Table** . If **Columns.Add** returns an error, it will not change the current row.

 _Name_ can be an explicit built-in property name, or a property name referenced by namespace. It must be referenced as the name in the English locale. For more information on referencing properties by namespace, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx). 

If you are adding a property which is an explicit built-in property in the object model, for example,  **[Contact.FirstName](contactitem-firstname-property-outlook.md)** , you must specify _Name_ as the explicit built-in property name in English. For certain types of properties, the format used when adding these properties as columns affects how their values are expressed in the **Table** . For more information on property value representation in a **Table** , see[Factors Affecting Property Value Representation in the Table and View Classes](http://msdn.microsoft.com/library/13cf9945-a9e0-bb32-a2cb-74366a365ae1%28Office.15%29.aspx).

If you are adding a custom property to a  **Table** , referencing the property by the MAPI string namespace, you will have to explicitly append the type of the property to the end of the property reference. For example, to add the custom property `MyCustomProperty`, which has the type Unicode string, you will have to explicitly append the type 001f to the reference, resulting in:  `http://schemas.microsoft.com/mapi/string/{HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}/MyCustomProperty/0x0000001f`, where  `{HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}` represents the namespace GUID.

Certain properties cannot be added to a  **Table** using **Columns.Add** , including binary properties, computed properties, and HTML or RTF body content. For more information, see[Unsupported Properties in a Table Object or Table Filter](http://msdn.microsoft.com/library/0e37f03f-7677-ca29-d0b2-8b45c026e5f1%28Office.15%29.aspx).

While  **[Items.SetColumns](items-setcolumns-method-outlook.md)** can be used to facilitate caching certain properties for extremely fast access to those properties of an **[Items](items-object-outlook.md)** collection, some properties are restricted from **SetColumns** . Since these restrictions do not apply to **Columns.Add** , the **Table** object is a less restrictive alternative than **Items** .


## Example

The following code sample illustrates how to obtain a  **Table** object based on the **LastModificationTime** of items in the Inbox. It also shows how to remove the default columns of the **Table** , add specific columns, and print the values of the corresponding properties of these items.


```vb
Sub AddColumns() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```


## See also


#### Concepts


[Columns Object](columns-object-outlook.md)

