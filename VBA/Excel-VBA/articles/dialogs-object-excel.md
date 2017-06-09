---
title: Dialogs Object (Excel)
keywords: vbaxl10.chm253072
f1_keywords:
- vbaxl10.chm253072
ms.prod: excel
api_name:
- Excel.Dialogs
ms.assetid: d1d54f0e-6057-92f5-4f4c-254c51e36040
ms.date: 06/08/2017
---


# Dialogs Object (Excel)

A collection of all the  **[Dialog](dialog-object-excel.md)** objects in Microsoft Excel.


## Remarks

 Each **Dialog** object represents a built-in dialog box. You cannot create a new built-in dialog box or add one to the collection. The only useful thing you can do with a **Dialog** object is use it with the **[Show](dialog-show-method-excel.md)** method to display the dialog corresponding dialog box.

The Microsoft Excel Visual Basic object library includes built-in constants for many of the built-in dialog boxes. Each constant is formed from the prefix "xlDialog" followed by the name of the dialog box. For example, the  **Apply Names** dialog box constant is **xlDialogApplyNames**, and the **Find File** dialog box constant is **xlDialogFindFile**. These constants are members of the **[XlBuiltinDialog](xlbuiltindialog-enumeration-excel.md)** enumerated type.


## Example

Use the [Dialogs](application-dialogs-property-excel.md) property to return the **Dialogs** collection. The following code example displays the number of available built-in Microsoft Excel dialog boxes.


```
MsgBox Application.Dialogs.Count
```

Use  **Dialogs** ( _index_ ), where _index_ is a built-in constant identifying the dialog box, to return a single **Dialog** object. The following example runs the built-in **File Open** dialog box.




```
dlgAnswer = Application.Dialogs(xlDialogOpen).Show
```



 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)

The following code example opens an e-mail message in Microsoft Outlook with the current workbook attached.




```
Sub SendIt() 
    Application.Dialogs(xlDialogSendMail).Show arg1:="ask@mrexcel.com", arg2:="This goes in the subject line" 
End Sub 

```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


## Properties
<a name="AboutContributor"> </a>



|**Name**|
|:-----|
|[Application](dialogs-application-property-excel.md)|
|[Count](dialogs-count-property-excel.md)|
|[Creator](dialogs-creator-property-excel.md)|
|[Item](dialogs-item-property-excel.md)|
|[Parent](dialogs-parent-property-excel.md)|

## See also
<a name="AboutContributor"> </a>


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
