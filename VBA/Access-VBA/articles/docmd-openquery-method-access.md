---
title: DoCmd.OpenQuery Method (Access)
keywords: vbaac10.chm4162
f1_keywords:
- vbaac10.chm4162
ms.prod: access
api_name:
- Access.DoCmd.OpenQuery
ms.assetid: 3ea20a28-8dd4-e54c-831b-e7e5444aa793
ms.date: 06/08/2017
---


# DoCmd.OpenQuery Method (Access)

The  **OpenQuery** method carries out the **OpenQuery** action in Visual Basic.


## Syntax

 _expression_. **OpenQuery**( ** _QueryName_**, ** _View_**, ** _DataMode_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _QueryName_|Required|**Variant**|A string expression that's the valid name of a query in the current database. If you execute Visual Basic code containing the  **OpenQuery** method in a library database, Microsoft Access looks for the query with this name first in the library database, then in the current database.|
| _View_|Optional|**AcView**|A  **[AcView](acview-enumeration-access.md)** constant that specifies the view in which the query will open. The default value is **acViewNormal**.|
| _DataMode_|Optional|**AcOpenDataMode**|A  **[AcOpenDataMode](acopendatamode-enumeration-access.md)** constant that specifies the data entry mode for the query. The default value is **acEdit**.|

## Remarks

You can use the  **OpenQuery** method to open a select or crosstab query in Datasheet view, Design view, or Print Preview. This action runs an action query. You can also select a data entry mode for the query.


 **Note**  This method is only available in the Microsoft Access database environment. See the  **OpenView** or **OpenStoredProcedure** methods if using the Microsoft Access Project environment (.adp).

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Calling Action Queries](http://www.utteraccess.com/wiki/index.php/Calling_Action_Queries)
    

## Example

The following example opens Sales Totals Query in Datasheet view and enables the user to view but not to edit or add records:


```vb
DoCmd.OpenQuery "Sales Totals Query", , acReadOnly
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[DoCmd Object](docmd-object-access.md)

