---
title: DoCmd.FindRecord Method (Access)
keywords: vbaac10.chm4151
f1_keywords:
- vbaac10.chm4151
ms.prod: access
api_name:
- Access.DoCmd.FindRecord
ms.assetid: dc48bc3d-5408-40a8-509b-e52b48b26187
ms.date: 06/08/2017
---


# DoCmd.FindRecord Method (Access)

The  **FindRecord** method carries out the FindRecord action in Visual Basic.


## Syntax

 _expression_. **FindRecord**( ** _FindWhat_**, ** _Match_**, ** _MatchCase_**, ** _Search_**, ** _SearchAsFormatted_**, ** _OnlyCurrentField_**, ** _FindFirst_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindWhat_|Required|**Variant**|An expression that evaluates to text, a number, or a date. The expression contains the data to search for.|
| _Match_|Optional|**AcFindMatch**|An  **[AcFindMatch](acfindmatch-enumeration-access.md)** constant that specifies where to search for the match. The default value is **acEntire**.|
| _MatchCase_|Optional|**Variant**|Use  **True** for a case-sensitive search and **False** for a search that's not case-sensitive. If you leave this argument blank, the default ( **False** ) is assumed.|
| _Search_|Optional|**AcSearchDirection**|An  **[AcSearchDirection](acsearchdirection-enumeration-access.md)** constant that specifies the direction to search. The default value is **acSearchAll**.|
| _SearchAsFormatted_|Optional|**Variant**|Use  **True** to search for data as it's formatted and **False** to search for data as it's stored in the database. If you leave this argument blank, the default ( **False** ) is assumed.|
| _OnlyCurrentField_|Optional|**AcFindField**|An  **[AcFindField](acfindfield-enumeration-access.md)** constant that specifies whether to search all fields, or only the current field. The default value is **acCurrent**.|
| _FindFirst_|Optional|**Variant**|Use  **True** to start the search at the first record. Use **False** to start the search at the record following the current record. If you leave this argument blank, the default ( **True** ) is assumed.|

## Remarks

When a procedure calls the FindRecord method, Access searches for the specified data in the records (the order of the search is determined by the setting of the Search argument). When Access finds the specified data, the data is selected in the record.

The  **FindRecord** method does not return a value indicating its success or failure. To determine whether a value exists in a recordset, use the **FindFirst**, **FindNext**, **FindPrevious** or **FindLast** method of the **Recordset** object. These methods set the value of the **NoMatch** property to **True** if the spcified value is not found.


## Example

The following example finds the first occurrence in the records of the name Smith in the current field. It doesn't find occurrences of smith or Smithson.


```vb
DoCmd.FindRecord "Smith",, True,, True
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

