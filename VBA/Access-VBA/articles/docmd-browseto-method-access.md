---
title: DoCmd.BrowseTo Method (Access)
keywords: vbaac10.chm5973
f1_keywords:
- vbaac10.chm5973
ms.prod: access
api_name:
- Access.DoCmd.BrowseTo
ms.assetid: 7cfd2cc5-ad2d-4bf8-ed90-1fb6adf1890a
ms.date: 06/08/2017
---


# DoCmd.BrowseTo Method (Access)

The  **BrowseTo** method performs the BrowseTo action in Visual Basic.


## Syntax

 _expression_. **BrowseTo**( ** _ObjectType_**, ** _ObjectName_**, ** _PathtoSubformControl_**, ** _WhereCondition_**, ** _Page_**, ** _DataMode_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**[AcBrowseToObjectType](acbrowsetoobjecttype-enumeration-access.md)**|The object type to which to browse.|
| _ObjectName_|Required|**Variant**|The object that loads inside the subform control referenced by the PathtoSubformControl argument. |
| _PathtoSubformControl_|Optional|**Variant**|If specified, the path from the main form of the application to the target subform control that loads the object specified by the ObjectName argument.|
| _WhereCondition_|Optional|**Variant**|If specified, replaces the Where condition of the object record source.|
| _Page_|Optional|**Variant**|If specified, sets the page of the continuous form that will be made the current page. This argument is Web only.|
| _DataMode_|Optional|**[AcFormOpenDataMode](acformopendatamode-enumeration-access.md)**|If specified, the data entry mode of the form.|

## Remarks

Use the  **BrowseTo** method to navigate between objects in place. You can also change the source object of a subform control by specifying the PathToSubFormControl argument. You can use **BrowseTo** to navigate from form1 to form2 without opening up a new window.

The PathToSubFormControl argument must be specified using the syntax in the following example:

Main Form.SubForm Ctrl 1>Form 2.SubForm Ctrl 2>Form 3.SubFormCtrl3

In this example, the Main Form is the top level form in the Access client application. The PathToSubFormControl argument must alternately specify form and subform control names leading from the main form to the subform control that is the container of the object specified by the ObjectName argument. Each subform control specified must be a control on the form that precedes it. The path must end with a subform control.


## Example

The following code example opens the "EventDS" form in place in edit mode in the "NavigationSubform" subform control of the "Main" form.


```vb
DoCmd.BrowseTo ObjectType:=acBrowseToForm, _ 
ObjectName:="EventDS", _ 
PathToSubformControl:="Main.NavigationSubform", _ 
WhereCondition:="", _ 
Page:="", _ 
DataMode:=acFormEdit
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

