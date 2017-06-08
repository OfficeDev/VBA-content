---
title: Dialog Object (Excel)
keywords: vbaxl10.chm255072
f1_keywords:
- vbaxl10.chm255072
ms.prod: excel
api_name:
- Excel.Dialog
ms.assetid: adabcd3b-fc48-d314-3ae5-f1b2ba148383
ms.date: 06/08/2017
---


# Dialog Object (Excel)

Represents a built-in Microsoft Excel dialog box.


## Remarks

 The **Dialog** object is a member of the **[Dialogs](dialogs-object-excel.md)** collection. The **Dialogs** collection contains all the built-in dialog boxes in Microsoft Excel. You cannot create a new built-in dialog box or add one to the collection. The only useful thing you can do with a **Dialog** object is use it with the **[Show](dialog-show-method-excel.md)** method to display the corresponding dialog box.

The Microsoft Excel Visual Basic object library includes built-in constants for many of the built-in dialog boxes. Each constant is formed from the prefix "xlDialog" followed by the name of the dialog box. For example, the  **Apply Names** dialog box constant is **xlDialogApplyNames**, and the **Find File** dialog box constant is **xlDialogFindFile**. These constants are members of the **[XlBuiltinDialog](xlbuiltindialog-enumeration-excel.md)** enumerated type.


## Example

Use  **[Dialogs](application-dialogs-property-excel.md)** ( _index_ ), where _index_ is a built-in constant identifying the dialog box, to return a single **Dialog** object. The following example runs the built-in **Open** dialog box ( **File** menu). The **Show** method returns **True** if Microsoft Excel successfully opens a file; it returns **False** if the user cancels the dialog box.


```
dlgAnswer = Application.Dialogs(xlDialogOpen).Show
```


## Methods



|**Name**|
|:-----|
|[Show](dialog-show-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](dialog-application-property-excel.md)|
|[Creator](dialog-creator-property-excel.md)|
|[Parent](dialog-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
