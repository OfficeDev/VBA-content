---
title: QueryTable.EditWebPage Property (Excel)
keywords: vbaxl10.chm518130
f1_keywords:
- vbaxl10.chm518130
ms.prod: excel
api_name:
- Excel.QueryTable.EditWebPage
ms.assetid: 4de607d1-266f-cbd4-c236-af748cfe0d03
ms.date: 06/08/2017
---


# QueryTable.EditWebPage Property (Excel)

Returns or sets the web page Uniform Resource Locator (URL) for a web query. Read/write  **Variant** .


## Syntax

 _expression_ . **EditWebPage**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

The  **EditWebPage** property returns **null** if not set. The **EditWebPage** property is only meaningful if the query type is Web or OLE.

If the  **EditWebPage** is not null then ignore the **[WebTables](querytable-webtables-property-excel.md)** property for refreshing. As a result an XML query and the **[WebTable](querytable-webtables-property-excel.md)** property refers to the table in the original Web page and should only be used in the edit case to pre-populate the **Web Query** dialog box.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **EditWebPage** property applies only to **QueryTable** objects.


## Example

In this example, Microsoft Excel displays to the user a Web page URL. This example assumes a  **QueryTable** object in cell A1 exists in the active worksheet and that a file called "MyHomepage.htm" exists on the C: drive.


```vb
Sub ReturnURL() 
 
 ' Set the EditWebPage property to a source. 
 Range("A1").QueryTable.EditWebPage = "C:\MyHomepage.htm" 
 
 ' Display the source to the user. 
 MsgBox Range("A1").QueryTable.EditWebPage 
 
End Sub
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

