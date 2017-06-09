---
title: ProtectedViewWindow Object (Excel)
keywords: vbaxl10.chm914072
f1_keywords:
- vbaxl10.chm914072
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow
ms.assetid: 6a32240c-c90b-c51a-6f8e-c3ff496b9855
ms.date: 06/08/2017
---


# ProtectedViewWindow Object (Excel)

Represents a  **Protected View** window.


## Remarks

A  **Protected View** window is used to display a workbook from a potentially unsafe location. Unsafe locations are defined as the following:


- Files opened from the Internet.
    
- Attachments opened from Outlook.
    
- Files blocked by File Block Policy.
    
- Files that fail Office File Validation.
    
- Files explicitly opened in  **Protected View** by using the **Open in Protected View** command of the **Open** button in the **Open** dialog box.
    


Workbooks displayed in a  **Protected View** window cannot be edited and are restricted from running active content such as Visual Basic for Applications macros and data connections. For more information about **Protected View** windows, see[What is Protected View?](http://office.microsoft.com/en-us/excel-help/what-is-protected-view-HA010355931.aspx?CTT=1)

 To return a single **ProtectedViewWindow** object from the **[ProtectedViewWindows](protectedviewwindows-object-excel.md)** collection, use `ProtectedViewWindows(Index)`, where  _Index_ is the index number of the window you want to open. You can also access the **ProtectedViewWindow** object that represents the active **Protected View** window by using the **[ActiveProtectedViewWindow](application-activeprotectedviewwindow-property-excel.md)** property of the **[Application](application-object-excel.md)** object.

After you access a  **ProtectedViewWindow** object, use the **[Workbook](protectedviewwindow-workbook-property-excel.md)** property to access the **[Workbook](workbook-object-excel.md)** object that represents the workbook file that is open in the **Protected View** window. Because a **Protected View** window is designed to protect the user from potentially malicious code, the operations you can perform by using a **Workbook** object returned by a **ProtectedViewWindow** object will be limited. Operations that are not allowed will return an error.


## Example

 The following code example accesses the **Workbook** object that represents the workbook that is open in the first **Protected View** window.


```vb
Dim wbProtected As Workbook 
 
If Application.ProtectedViewWindows.Count > 0 Then 
    Set wbProtected = Application.ProtectedViewWindows(1).Workbook 
End If 

```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


