---
title: Application.PrintCommunication Property (Excel)
keywords: vbaxl10.chm133323
f1_keywords:
- vbaxl10.chm133323
ms.prod: excel
api_name:
- Excel.Application.PrintCommunication
ms.assetid: 8b8ad1c5-1999-d733-44f4-734b7a388986
ms.date: 06/08/2017
---


# Application.PrintCommunication Property (Excel)

Specifies whether communication with the printer is turned on.  **Boolean** Read/write


## Syntax

 _expression_ . **PrintCommunication**

 _expression_ A variable that returns an **Application** object.


### Return Value

 **True** , if communication with the printer is turned on; otherwise **False** .


## Remarks

Set the  **PrintCommunication** property to **False** to speed up the execution of code that sets **[PageSetup](pagesetup-object-excel.md)** properties. Set the **PrintCommunication** property to **True** after setting properties to commit all cached **PageSetup** commands.


## Example

The following example suspends communication with the printer while setting  **PageSetup** properties.


```vb
Application.PrintCommunication = False 
 With ActiveSheet.PageSetup 
 .PrintTitleRows = "" 
 .PrintTitleColumns = "" 
 End With 
Application.PrintCommunication = True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

