---
title: Application.Workbooks Property (Excel)
keywords: vbaxl10.chm183115
f1_keywords:
- vbaxl10.chm183115
ms.prod: excel
api_name:
- Excel.Application.Workbooks
ms.assetid: 5291a324-87d7-3916-ffee-34c3389cea13
ms.date: 06/08/2017
---


# Application.Workbooks Property (Excel)

Returns a  **[Workbooks](workbooks-object-excel.md)** collection that represents all the open workbooks. Read-only.


## Syntax

 _expression_ . **Workbooks**

 _expression_ A variable that represents an **Application** object.


## Remarks

Using this property without an object qualifier is equivalent to using `Application.Workbooks`.

The collection returned by the  **Workbooks** property doesn't include open add-ins, which are a special kind of hidden workbook. You can, however, return a single open add-in if you know the file name. For example, `Workbooks("Oscar.xla")`will return the open add-in named "Oscar.xla" as a  **Workbook** object.


 **Note**  A workbook displayed in a protected view window is not a member of the  **Workbooks** collection. Instead, use the **[Workbook](protectedviewwindow-workbook-property-excel.md)** property of the **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object to access a workbook that is displayed in a protected view window.


## Example

This example activates the workbook Book1.xls.


```vb
Workbooks("BOOK1").Activate
```

This example opens the workbook Large.xls.




```vb
Workbooks.Open filename:="LARGE.XLS"
```

This example saves changes to and closes all workbooks except the one that's running the example.




```vb
For Each w In Workbooks 
    If w.Name <> ThisWorkbook.Name Then 
        w.Close savechanges:=True 
    End If 
Next w
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

