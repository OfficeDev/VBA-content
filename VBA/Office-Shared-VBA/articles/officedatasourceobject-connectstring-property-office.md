---
title: OfficeDataSourceObject.ConnectString Property (Office)
keywords: vbaof11.chm232001
f1_keywords:
- vbaof11.chm232001
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.ConnectString
ms.assetid: 56c599a5-f493-ea5a-3d2b-a3dae973d71c
ms.date: 06/08/2017
---


# OfficeDataSourceObject.ConnectString Property (Office)

Gets or sets a  **String** that represents the connection to the specified mail merge data source. Read/write.


## Syntax

 _expression_. **ConnectString**

 _expression_ A variable that represents an **OfficeDataSourceObject** object.


## Example

This example checks if the connection string contains the characters ODSOOutlook and displays a message accordingly.


```
Sub VerifyCorrectDataSource() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 If InStr(appOffice.ConnectString, "ODSOOutlook") > 0 Then 
 MsgBox "Your Outlook address book is used as the data source." 
 Else 
 MsgBox "Your Outlook address book is not used as the data source." 
 End If 
 
End Sub
```


## See also


#### Concepts


[OfficeDataSourceObject Object](officedatasourceobject-object-office.md)
#### Other resources


[OfficeDataSourceObject Object Members](officedatasourceobject-members-office.md)

