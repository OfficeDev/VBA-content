---
title: DoCmd.RunDataMacro Method (Access)
keywords: vbaac10.chm5978
f1_keywords:
- vbaac10.chm5978
ms.prod: access
api_name:
- Access.DoCmd.RunDataMacro
ms.assetid: e95b7a8e-a502-67c6-1941-dd5a06c08ef7
ms.date: 06/08/2017
---


# DoCmd.RunDataMacro Method (Access)

Use the  **RunDataMacro** method to run a named data macro from Visual Basic.


## Syntax

 _expression_. **RunDataMacro**( ** _MacroName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MacroName_|Required|**Variant**|Name of the saved macro. The name must include the name of the table to which the data macro is attached (for example, "Comments.AddComment").|

## Remarks

Use the  **RunDataMacro** method to reuse a named data macro in Visual Basic code.

If the data macro requires parameters, you must first create them using the  **[SetParameter](docmd-setparameter-method-access.md)** method prior to calling the **RunDataMacro** method. Each call to **SetParameter** creates a single named parameter.


## Example

The following code example creates two parameters for use by the AddComment data macro. The two parameters are named prmComment and prmRelatedID, respectively. The value of the txtComment textbox is stored in the prmComment parameter. The value of the txtId textbox is stored in the prmRelatedID parameter. The "Comments.AddComment" data macro is then run.


```vb
Private Sub cmdAddComment_Click() 
DoCmd.SetParameter "prmComment", Me.txtComment 
DoCmd.SetParameter "prmRelatedID", Me.txtId 
DoCmd.RunDataMacro "Comments.AddComment" 
End Sub
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

