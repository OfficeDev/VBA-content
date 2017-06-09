---
title: Application.LoadCustomUI Method (Access)
keywords: vbaac10.chm12616
f1_keywords:
- vbaac10.chm12616
ms.prod: access
api_name:
- Access.Application.LoadCustomUI
ms.assetid: 59be6be9-d7a0-98f3-b9c0-57ecba5651f6
ms.date: 06/08/2017
---


# Application.LoadCustomUI Method (Access)

Loads XML markup that represents a customized ribbon.


## Syntax

 _expression_. **LoadCustomUI**( ** _CustomUIName_**, ** _CustomUIXML_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomUIName_|Required|**String**|The name that will be used to identify the customized ribbon.|
| _CustomUIXML_|Required|**String**|The XML markup code that defines the customized ribbon.|

## Remarks

To create and make the ribbon available to Access, you first create a module in the database with a procedure that calls the  **LoadCustomUI** method, passing in the name of the ribbon and the XML customization markup. The XML markup can come from a **Recordset** object created from a table, from a source external to the database (such as an XML file that you must parse into a **String** ), or from XML markup embedded directly inside of the procedure.

You can make different ribbons available by using multiple calls to the  **LoadCustomUI** method, passing in different XML markup, as long as the name of each ribbon and the id attribute of the tabs that make up the ribbon are unique.

 After the procedure is complete, you then create an AutoExec macro that calls the procedure by using the RunCode action. That way, when the application is started, the **LoadCustomUI** method is automatically executed and all of the custom ribbons are made available to the application


## Example

The following code example creates a  **Recordset** from any table that contains the word "Ribbons" in its name. Then it calls the **LoadCustomUI** method to load the ribbons in order to make them available to the database. Finally, it closes the recordset and the reference to the **Database** object.


```vb
Function LoadRibbons() 
Dim i As Integer 
Dim db As DAO.Database 
Set db = Application.CurrentDb 
 
For i = 0 To (db.TableDefs.Count - 1) 
 If (InStr(1, db.TableDefs(i).Name, "Ribbons")) Then 
 Dim rs As DAO.Recordset 
 Set rs = CurrentDb.OpenRecordset(db.TableDefs(i).Name) 
 rs.MoveFirst 
 
 While Not rs.EOF 
 Application.LoadCustomUI rs("RibbonName").Value, rs("RibbonXml").Value 
 
 rs.MoveNext 
 Wend 
 
 rs.Close 
 Set rs = Nothing 
 End If 
Next i 
 
db.Close 
Set db = Nothing 
End Function
```


## See also


#### Concepts


[Application Object](application-object-access.md)

