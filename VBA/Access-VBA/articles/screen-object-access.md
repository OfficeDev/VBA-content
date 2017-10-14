---
title: Screen Object (Access)
keywords: vbaac10.chm12484
f1_keywords:
- vbaac10.chm12484
ms.prod: access
api_name:
- Access.Screen
ms.assetid: 00743775-071b-9ccd-7687-f3b992e9346e
ms.date: 06/08/2017
---


# Screen Object (Access)

The  **Screen** object refers to the particular form, report, or control that currently has the focus.


## Remarks

You can use the  **Screen** object together with its properties to refer to a particular form, report, or control that has the focus.

For example, you can use the  **Screen** object with the **ActiveForm** property to refer to the form in the active window without knowing the form's name. The following example displays the name of the form in the active window:




```
MsgBox Screen.ActiveForm.Name
```

Referring to the  **Screen** object doesn't make a form, report, or control active. To make a form, report, or control active, you must use the **SelectObject** method of the **[DoCmd](http://msdn.microsoft.com/library/3ce44cca-9979-0a1e-9787-079a52ce528f%28Office.15%29.aspx)** object.

If you refer to the  **Screen** object when there's no active form, report, or control, Microsoft Access returns a run-time error. For example, if a standard module is in the active window, the code in the preceding example would return an error.


## Example

The following example uses the  **Screen** object to print the name of the form in the active window and of the active control on that form:


```
Sub ActiveObjects() 
 Dim frm As Form, ctl As Control 
 
 ' Return Form object pointing to active form. 
 Set frm = Screen.ActiveForm 
 MsgBox frm.Name &amp; " is the active form." 
 ' Return Control object pointing to active control. 
 Set ctl = Screen.ActiveControl 
 MsgBox ctl.Name &amp; " is the active control " _ 
 &amp; "on this form." 
End Sub 

```



|**Name**|
|:-----|
|[ActiveControl](http://msdn.microsoft.com/library/01d76377-c88d-8f64-b13b-c80f4d296834%28Office.15%29.aspx)|
|[ActiveDatasheet](http://msdn.microsoft.com/library/cff189e7-9b8a-280f-e287-e4367f8ac134%28Office.15%29.aspx)|
|[ActiveForm](http://msdn.microsoft.com/library/5cf41661-656e-e62f-530e-0d2fa5466146%28Office.15%29.aspx)|
|[ActiveReport](http://msdn.microsoft.com/library/efcf6bfd-2749-5b5c-d7ca-a26168bfcb65%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/1d2fe0bb-5c08-8c16-2d09-9ed515d9eb43%28Office.15%29.aspx)|
|[MousePointer](http://msdn.microsoft.com/library/e7ee88cf-7eb8-a447-d671-1549cdbcb4fd%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/82df42fb-8049-8abc-09b3-ad70860a1c43%28Office.15%29.aspx)|
|[PreviousControl](http://msdn.microsoft.com/library/089a62f7-2f3f-93e8-8e84-1b77d4f12e79%28Office.15%29.aspx)|

## See also


#### Other resources


[Screen Object Members](http://msdn.microsoft.com/library/82c9e4cb-95a9-6842-2629-bcd71c81838f%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
