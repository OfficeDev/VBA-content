---
title: DoCmd.OpenDataAccessPage Method (Access)
keywords: vbaac10.chm4648
f1_keywords:
- vbaac10.chm4648
ms.prod: access
api_name:
- Access.DoCmd.OpenDataAccessPage
ms.assetid: 130dcb88-e3e6-25a6-186c-bf541d114169
ms.date: 06/08/2017
---


# DoCmd.OpenDataAccessPage Method (Access)

The  **OpenDataAccessPage** method carries out the OpenDataAccessPage action in Visual Basic.


## Syntax

 _expression_. **OpenDataAccessPage**( ** _DataAccessPageName_**, ** _View_** )

 _expression_ An expression that returns a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataAccessPageName_|Required|**Variant**|A string expression that's the valid name of a data access page in the current database. If you execute Visual Basic code containing the  **OpenDataAccessPage** method in a library database, Microsoft Access looks for the form with this name, first in the library database, then in the current database.|
| _View_|Optional|**AcDataAccessPageView**|The view in which to open the data access page. In Access, this must be set to  **acDataAccessPageBrowse**.|

## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

