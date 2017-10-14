---
title: Application.GetNamespace Method (Outlook)
keywords: vbaol11.chm717
f1_keywords:
- vbaol11.chm717
ms.prod: outlook
api_name:
- Outlook.Application.GetNamespace
ms.assetid: 6175d0d9-5a61-ce45-35c0-b70895d757b3
ms.date: 06/08/2017
---


# Application.GetNamespace Method (Outlook)

Returns a  **[NameSpace](namespace-object-outlook.md)** object of the specified type.


## Syntax

 _expression_ . **GetNamespace**( **_Type_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **String**|The type of name space to return.|

### Return Value

A  **NameSpace** object that represents the specified namespace.


## Remarks

The only supported name space type is "MAPI". The  **GetNameSpace** method is functionally equivalent to the **Session** property.


## Example

This Visual Basic for Applications (VBA) example uses the  **[CurrentFolder](explorer-currentfolder-property-outlook.md)** property to change the displayed folder to the user's **Calendar** folder.


```vb
Sub ChangeCurrentFolder() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set Application.ActiveExplorer.CurrentFolder = _ 
 
 myNamespace.GetDefaultFolder(olFolderCalendar) 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)
#### Other resources


[How to: Obtain and Log On to an Instance of Outlook](http://msdn.microsoft.com/library/ef369364-6500-2759-3ef4-ed4411112e96%28Office.15%29.aspx)


