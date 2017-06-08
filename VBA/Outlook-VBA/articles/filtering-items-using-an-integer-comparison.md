---
title: Filtering Items Using an Integer Comparison
ms.prod: outlook
ms.assetid: c67d95b2-f635-b751-d9c6-c7bdf406a01a
ms.date: 06/08/2017
---


# Filtering Items Using an Integer Comparison

You can compare an integer property with an integer value in a filter string using Microsoft Jet syntax or DAV Searching and Locating (DASL) syntax. You can specify the integer value with or without quotation marks as delimiters. The following three filter strings filter on the condition that the  **Importance** value is high:


```
criteria = "[Importance] = 2"
```


If you want to use a value from an integer enumeration, convert the value to a string and append it to the filter string. The following filters are equivalent and test for items with importance set to high: 




```
criteria = "[Importance] = " _ &; CStr(Outlook.OlImportance.olImportanceHigh)

criteria = "@SQL=" &; Chr(34) &; "urn:schemas:httpmail:importance" _ &; Chr(34) &; " = 2"
```


