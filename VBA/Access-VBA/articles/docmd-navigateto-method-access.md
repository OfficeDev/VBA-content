---
title: DoCmd.NavigateTo Method (Access)
keywords: vbaac10.chm5689
f1_keywords:
- vbaac10.chm5689
ms.prod: access
api_name:
- Access.DoCmd.NavigateTo
ms.assetid: 27a6e4ee-1c03-2652-3c5a-73c45f3109df
ms.date: 06/08/2017
---


# DoCmd.NavigateTo Method (Access)

You can use the  **NavigateTo** method to control the display of database objects in the Navigation Pane. .


## Syntax

 _expression_. **NavigateTo**( ** _Category_**, ** _Group_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Category_|Optional|**Variant**|The category by which you want the Navigation Pane to display objects. |
| _Group_|Optional|**Variant**|Determines which objects in the category appear in the Navigation Pane. If you leave the this argument blank, the Navigation Pane will display all database objects grouped by the criteria you specify in the  _Category_ argument. Examples of valid _Group_ arguments for the various _Category_ arguments are shown in the following table.|

## Remarks

For example, you can change how the database objects are categorized, and you can filter the objects so that only certain ones are displayed. 

This action is similar to selecting categories and groups from the title bar of the Navigation Pane.

Valid  _Group_ arguments depend on which _Category_ argument is used. If you enter an invalid _Group_ argument, an error message appears.

The following table contains examples of valid  _Group_ arguments for each _Category_ argument.



|**Category argument**|**Example Group arguments**|
|:-----|:-----|
|Object Type|Tables; Forms; Queries; Pages; Macros; Modules|
|Tables and Views|Names of specific tables in your database|
|Modified Date|Today; Yesterday; Last Month; Older|
|Created Date|Today; Yesterday; Last Month; Older|
|Custom Category|Names of groups you have created for the specified custom category|

 **Note**  To navigate to the top level of a category (for example,  **All Tables**,  **All Access Objects**, or  **All Dates**), you must leave the  _Group_ argument blank. For example, when the _Category_ argument is **Object Type**, entering **All Access Objects** as a _Group_ argument results in an error.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

