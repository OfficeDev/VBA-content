---
title: Application.FollowHyperlink Method (Project)
keywords: vbapj.chm1307
f1_keywords:
- vbapj.chm1307
ms.prod: project-server
api_name:
- Project.Application.FollowHyperlink
ms.assetid: d612e80b-93c1-7312-d164-be552b580370
ms.date: 06/08/2017
---


# Application.FollowHyperlink Method (Project)

Opens the document specified by a hyperlink address.


## Syntax

 _expression_. **FollowHyperlink**( ** _Address_**, ** _SubAddress_**, ** _AddHistory_**, ** _NewWindow_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Address_|Optional|**String**|The address of the target document. If  **Address** is omitted and a text field is selected, the text of the selected field is used. If **Address** is omitted and a text field is not selected, Project returns an error.|
| _SubAddress_|Optional|**String**|A location within the target document.|
| _AddHistory_|Optional|**Boolean**|**True** if the target document should be added to the History folder. The default value is **True**.|
| _NewWindow_|Optional|**Boolean**|**True** if the target document should display in a new window. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example opens a hyperlink to the Microsoft Web site in its own window.


```vb
Sub GoToMicrosoft() 
    Application.FollowHyperlink Address:="http://www.Microsoft.com", _ 
        NewWindow:=True, AddHistory:=True 
End Sub
```


