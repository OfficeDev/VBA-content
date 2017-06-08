---
title: Filtering Items Using a Variable
ms.prod: outlook
ms.assetid: 4be50a96-a27e-ecbf-1f19-b8825a33c2eb
ms.date: 06/08/2017
---


# Filtering Items Using a Variable

You can use values from variables as part of a filter in Microsoft Jet syntax or DAV Searching and Locating (DASL) syntax. The following example illustrates the use of variables as part of a filter: 


```
fullname = "Dan Wilson" 
' This approach uses Chr(34) to delimit the value.  
criteria = "[FullName] = " &; Chr(34) &; fullname _ &; Chr(34) 
' This approach uses the double quotation mark to delimit the value.  
criteria = "[FullName] = """ &; fullname &; """" 
' This approach uses the escaped hexadecimal value of the double quotation mark to delimit the value.  
criteria = "[FullName] = "%22" &; fullname &; "%22"" 

```


