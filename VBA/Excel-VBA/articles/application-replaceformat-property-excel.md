---
title: Application.ReplaceFormat Property (Excel)
keywords: vbaxl10.chm133263
f1_keywords:
- vbaxl10.chm133263
ms.prod: excel
api_name:
- Excel.Application.ReplaceFormat
ms.assetid: df2242dc-9f23-b3c8-455d-1f0474eca873
ms.date: 06/08/2017
---


# Application.ReplaceFormat Property (Excel)

Sets the replacement criteria to use in replacing cell formats. The replacement criteria is then used in a subsequent call to the Replace method of the Range object.


## Syntax

 _expression_ . **ReplaceFormat**

 _expression_ A variable that represents an **Application** object.


## Example

The following example sets the search criteria to find cells containing Arial, Regular, Size 10 font, replaces their formats with Arial, Bold, Size 8 font, and then calls the Replace method, with the optional arguments of SearchFormat and ReplaceFormat set to True to actually make the changes.


```vb
Sub MakeBold() 
 
 ' Establish search criteria. 
 With Application.FindFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Regular" 
 .Size = 10 
 End With 
 
 ' Establish replacement criteria. 
 With Application.ReplaceFormat.Font 
 .Name = "Arial" 
 .FontStyle = "Bold" 
 .Size = 8 
 End With 
 
 ' Notify user. 
 With Application.ReplaceFormat.Font 
 MsgBox .Name &; "-" &; .FontStyle &; "-" &; .Size &; _ 
 " font is what the search criteria will replace cell formats with." 
 End With 
 
 ' Make the replacements in the worksheet. 
 Cells.Replace What:="", Replacement:="", _ 
 SearchFormat:=True, ReplaceFormat:=True 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

