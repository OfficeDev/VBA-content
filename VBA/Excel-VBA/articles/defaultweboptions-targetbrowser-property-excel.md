---
title: DefaultWebOptions.TargetBrowser Property (Excel)
keywords: vbaxl10.chm660090
f1_keywords:
- vbaxl10.chm660090
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.TargetBrowser
ms.assetid: 785efc30-ef17-1745-874d-a3be861d450b
ms.date: 06/08/2017
---


# DefaultWebOptions.TargetBrowser Property (Excel)

Returns or sets an  **[MsoTargetBrowser](http://msdn.microsoft.com/library/6ce561d2-c327-b433-3c91-df1036e87a75%28Office.15%29.aspx)** constant indicating the browser version. Read/write.


## Syntax

 _expression_ . **TargetBrowser**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks





| **MsoTargetBrowser** can be one of these **MsoTargetBrowser** constants.|
| **msoTargetBrowserIE4** . Microsoft Internet Explorer 4.0 or later.|
| **msoTargetBrowserIE5** . Microsoft Internet Explorer 5.0 or later.|
| **msoTargetBrowserIE6** . Microsoft Internet Explorer 6.0 or later.|
| **msoTargetBrowserV3** . Microsoft Internet Explorer 3.0, Netscape Navigator 3.0, or later.|
| **msoTargetBrowserV4** . Microsoft Internet Explorer 4.0, Netscape Navigator 4.0, or later.|

## Example

In this example, Microsoft Excel determines if the browser version for Web options is IE5 and notifies the user.


```vb
Sub CheckWebOptions() 
 
    Dim wkbOne As Workbook 
 
    Set wkbOne = Application.Workbooks(1) 
 
    ' Determine if IE5 is the target browser. 
    If wkbOne.WebOptions.TargetBrowser = msoTargetBrowserIE5 Then 
        MsgBox "The target browser is IE5 or later." 
    Else 
        MsgBox "The target browser is not IE5 or later." 
    End If 
 
End Sub
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

