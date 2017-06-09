---
title: Application.Dialogs Property (Excel)
keywords: vbaxl10.chm133118
f1_keywords:
- vbaxl10.chm133118
ms.prod: excel
api_name:
- Excel.Application.Dialogs
ms.assetid: 0d04aa87-9872-23e5-78e3-c9e3da2c8eb5
ms.date: 06/08/2017
---


# Application.Dialogs Property (Excel)

Returns a  **[Dialogs](dialogs-object-excel.md)** collection that represents all built-in dialog boxes. Read-only.


## Syntax

 _expression_ . **Dialogs**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays the  **Open** dialog box ( **File** menu).


```vb
Application.Dialogs(xlDialogOpen).Show
```



 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)



The following code example opens an e-mail message in Microsoft Outlook with the current workbook attached.




```vb
Sub SendIt() 
    Application.Dialogs(xlDialogSendMail).Show arg1:="ask@mrexcel.com", arg2:="This goes in the subject line" 
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Application Object](application-object-excel.md)

