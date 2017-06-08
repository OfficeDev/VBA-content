---
title: OLEDBConnection.LocaleID Property (Excel)
keywords: vbaxl10.chm794107
f1_keywords:
- vbaxl10.chm794107
ms.prod: excel
api_name:
- Excel.OLEDBConnection.LocaleID
ms.assetid: 6a92f9ca-247a-8da8-a32e-ec239380894a
ms.date: 06/08/2017
---


# OLEDBConnection.LocaleID Property (Excel)

Returns or sets the locale identifier for the specified connection. Read/write


## Syntax

 _expression_ . **LocaleID**

 _expression_ A variable that represents an **[OLEDBConnection](oledbconnection-object-excel.md)** object.


### Return Value

 **Integer**


## Remarks

Before you set the  **LocaleID** property to a new locale, you must set the **[RetrieveInOfficeUILang](oledbconnection-retrieveinofficeuilang-property-excel.md)** property of the **OLEDBConnection** object to **False** . For more information about valid Locale ID (LCID) values, search the MSDN Web site for "Locale IDs Assigned by Microsoft".


## Example

The following code example switches the language of the connection to Spanish.


```vb
Dim myConnection As OLEDBConnection 
Set myConnection = ThisWorkbook.Connections(1) 
 
With myConnection 
 .RetrieveInOfficeUILang = False 
 .LocaleID = 3082 
End With
```


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

