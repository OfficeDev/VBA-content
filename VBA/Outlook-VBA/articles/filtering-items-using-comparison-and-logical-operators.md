---
title: Filtering Items Using Comparison and Logical Operators
ms.prod: outlook
ms.assetid: 1f01f970-549f-5868-bbe8-e8792dfd657f
ms.date: 06/08/2017
---


# Filtering Items Using Comparison and Logical Operators




## Comparison Operators

You can use the following comparison operators in filter strings using Microsoft Jet syntax or DAV Searching and Locating (DASL) syntax:


- <
    
    Performs a less-than comparison.
    
- >
    
    Performs a greater-than comparison.
    
- <=
    
    Performs a less-than-or-equal-to comparison.
    
- >=
    
    Performs a greater-than-or-equal-to comparison.
    
- <>
    
    Performs a not-equal-to comparison.
    
- =
    
    Performs a equal-to comparison.
    

## Logical Operators

You can use the logical operators  **And**,  **Not**,  **Or** in filter strings in Jet or DASL syntax. The order of precedence of these operators, from the highest to the lowest, is: **Not**,  **And**,  **Or**. You can use parentheses to indicate specific precedence in a filter. Logical operators are case-insensitive. 


-  **Not**
    
    Performs a logical NOT on the condition. The following code retrieves all contacts whose first name is Jane and who do not work at Microsoft. 
    


```
  criteria = _ "[FirstName] = 'Jane' And Not([CompanyName] = 'Microsoft')"
```

-  **And**
    
    Performs a logical AND on the condition. The following code retrieves all contacts who work at Microsoft and whose first name is Mary.
    


```
  criteria = _ "[FirstName] = 'Mary' And [CompanyName] = 'Microsoft'"
```

-  **Or**
    
    Performs a logical OR on the condition. The following code returns all contact items that have either a first name of Peter or Paul. 
    


```
  criteria = "[FirstName] = 'Peter' Or [FirstName] = 'Paul'"
```


