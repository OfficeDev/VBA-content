---
title: Filtering Items Using a Boolean Comparison
ms.prod: outlook
ms.assetid: bd786159-f4eb-e649-e838-56d520b824cf
ms.date: 06/08/2017
---


# Filtering Items Using a Boolean Comparison

Boolean values are specified differently in a filter in Microsoft Jet syntax than in a filter in DAV Searching and Locating (DASL) syntax.


## Jet Queries

In Jet syntax, boolean operators such as True/False, Yes/No, On/Off, and so on, should be used as is and should not be converted to a string. For example, to create a filter to return unread items, you can use this filter:


```
criteria = "[UnRead] = True"
```


 **Note**  If you convert the boolean value to a comparison string by enclosing it in quotation marks, then a Jet filter using any non-empty comparison string and filtering on a boolean property will return items that have the property True. A Jet filter comparing an empty string with a boolean property will return items that have the property False. 


## DASL Queries

In DASL syntax, you must convert True/False to an integer value, where 0 represents False and 1 represents True; likewise for Yes/No and On/Off. The DASL filter to return unread items is as follows: 


```
criteria = "@SQL=" &; Chr(34) &; "urn:schemas:httpmail:read" &; Chr(34) _ &; " = 0"
```


