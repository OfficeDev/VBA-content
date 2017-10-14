---
title: Filtering Items Using Query Keywords
ms.prod: outlook
ms.assetid: d7e6b169-c5fd-7acc-f077-658a153a921f
ms.date: 06/08/2017
---


# Filtering Items Using Query Keywords

You can use the following query keywords only in DAV Searching and Locating (DASL) queries. Keywords are case-insensitive. Microsoft Jet does not support any query keywords.


-  **ci_startwith** and **ci_phrasematch**
    
    These content indexer keywords provide prefix and phrase matching respectively for items in an indexed store. For example, the following DASL query creates a filter for last name starting with "Smith" and uses a content indexer keyword to return the results:
    


```
  criteria = "@SQL=" &; Chr(34) _ 
&; "urn:schemas:contacts:sn" &; Chr(34) _ 
&; " ci_startswith 'Smith'"
```


    The following DASL query creates a filter for last name being exactly "Smith" and uses a content indexer keyword to return the results: 
    


```
  criteria = "@SQL=" &; Chr(34) _ 
&; "urn:schemas:contacts:sn" &; Chr(34) _ 
&; " ci_phrasematch 'Smith'"
```

-  **Is Null**
    
    Evaluates if a property is null. Returns True if the property is null and False if the property is not null.
    
     **Is Null** operations are useful to determine if a date property has been set or if a string property is empty. If the date is null, the local time value of the date will be equal to 1/1/4501.
    
    The syntax of  **Is Null** is as follows:
    


```sql
  [PropertyName] IS NULL
```


    where  _PropertyName_ is the name of a property referenced by namespace.
    
    You can combine the  **Is Null** keywords with the **Not** operator to evaluate if a property is not null.
    
    The following DASL query retrieves all contacts where the custom property  **Order Date** is not null and the **[CompanyName](contactitem-companyname-property-outlook.md)** property is exactly Microsoft:
    


```sql
  criteria = "@SQL=" &; "(NOT(" _ 
&; Chr(34) &; "http://schemas.microsoft.com/mapi/string/" _ 
&; "{00020329-0000-0000-C000-000000000046}/Order%20Date" &; Chr(34) _ 
&; " IS NULL) AND " _ &; Chr(34) &; "urn:schemas-microsoft-com:office:office#Company" 
&; Chr(34) _ &; " = 'Microsoft')"
```





