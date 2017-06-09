---
title: Application.EditHyperlink Method (Project)
keywords: vbapj.chm1310
f1_keywords:
- vbapj.chm1310
ms.prod: project-server
api_name:
- Project.Application.EditHyperlink
ms.assetid: d652ccc4-207e-933f-c281-a2d5d7db0b76
ms.date: 06/08/2017
---


# Application.EditHyperlink Method (Project)

Edits the hyperlink of the selected assignment, resource, or task.


## Syntax

 _expression_. **EditHyperlink**( ** _Name_**, ** _Address_**, ** _SubAddress_**, ** _ScreenTip_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the hyperlink as it appears in the Hyperlink field.|
| _Address_|Optional|**String**|The address of the target document.|
| _SubAddress_|Optional|**String**| A location within the target document.|
| _ScreenTip_|Optional|**String**|The ScreenTip text for the hyperlink.|

### Return Value

 **Boolean**


## Remarks

Using the  **EditHyperlink** method without specifying any arguments displays the **Edit Hyperlink** dialog box.


## Example

The following example first creates a hyperlink in the Gantt Chart view and then change the name to MyHyperLink.


```vb
Sub Edit_Hyperlink() 
 
 ViewApply Name:="&;Gantt Chart" 
 SelectRow Row:=2, RowRelative:=False 
 InsertHyperlink Name:="http://MSDN", Address:="http://msdn.microsoft.com/", SubAddress:="", ScreenTip:="" 
 
 EditHyperlink Name:="MyHyperLink" 
End Sub
```


