---
title: Application.RefreshTitleBar Method (Access)
keywords: vbaac10.chm12551
f1_keywords:
- vbaac10.chm12551
ms.prod: access
api_name:
- Access.Application.RefreshTitleBar
ms.assetid: 9924e3ff-714f-023e-460f-d4aba7702829
ms.date: 06/08/2017
---


# Application.RefreshTitleBar Method (Access)

The  **RefreshTitleBar** method refreshes the Microsoft Access title bar after the **[AppTitle](apptitle-property.md)** or **[AppIcon](appicon-property.md)** property has been set in Visual Basic.


## Syntax

 _expression_. **RefreshTitleBar**

 _expression_ A variable that represents an **Application** object.


### Return Value

Nothing


## Remarks

For example, you can change the caption in the Microsoft Access title bar to "Contacts Database" by setting the  **AppTitle** property.

The  **AppTitle** and **AppIcon** properties enable you to customize your application by changing the title and icon that appear in the Microsoft Access title bar. The title bar is not automatically updated after you have set these properties. In order for the change to the title bar to appear, you must use the **RefreshTitleBar** method.


 **Note**  In a Microsoft Access database, you can reset the  **AppTitle** and **AppIcon** properties to their default value by deleting them from the **Properties** collection representing the current database. After you delete these properties, you must use the **RefreshTitleBar** method to restore the Microsoft Access defaults to the title bar.

It the path to the icon specified by the  **AppIcon** property is invalid, then no changes will be refled in the title bar when you call this method.


## Example

The following example sets the  **AppTitle** property of the current database and applies the **RefreshTitleBar** method to update the title bar.


```vb
Sub ChangeTitle() 
 Dim obj As Object 
 Const conPropNotFoundError = 3270 
 
 On Error GoTo ErrorHandler 
 ' Return Database object variable pointing to 
 ' the current database. 
 Set dbs = CurrentDb 
 ' Change title bar. 
 dbs.Properties!AppTitle = "Contacts Database" 
 ' Update title bar on screen. 
 Application.RefreshTitleBar 
 Exit Sub 
 
ErrorHandler: 
 If Err.Number = conPropNotFoundError Then 
 Set obj = dbs.CreateProperty("AppTitle", dbText, "Contacts Database") 
 dbs.Properties.Append obj 
 Else 
 MsgBox "Error: " &; Err.Number &; vbCrLf &; Err.Description 
 End If 
 Resume Next 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

