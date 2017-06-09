---
title: DoCmd.SetParameter Method (Access)
keywords: vbaac10.chm5977
f1_keywords:
- vbaac10.chm5977
ms.prod: access
api_name:
- Access.DoCmd.SetParameter
ms.assetid: 55e64bab-1c5e-9da0-5425-c8ed7b0bb1c2
ms.date: 06/08/2017
---


# DoCmd.SetParameter Method (Access)

Use the  **SetParameter** method to create a parameter for use by the **[BrowseTo ](docmd-browseto-method-access.md)**, **[OpenForm](docmd-openform-method-access.md)**, **[OpenQuery](docmd-openquery-method-access.md)**, **[OpenReport](docmd-openreport-method-access.md)**, or **[RunDataMacro](docmd-rundatamacro-method-access.md)** methods.


## Syntax

 _expression_. **SetParameter**( ** _Name_**, ** _Expression_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**Variant**|The name of the parameter. The name must match the name of the parameter expected by the  **BrowseTo**, **OpenForm**, **OpenQuery**, **OpenReport**, or **RunDataMacro** method.|
| _Expression_|Required|**Variant**|An expression that evaluates to a value to assign to the parameter.|

## Remarks

You must create as many calls to the  **SetParameter** method as are necessary to create the parameters you need.

Each call to  **SetParameter** adds or updates a single parameter in an internal parameters collection. The parameters collection is passed to the **BrowseTo**, **OpenForm**, **OpenQuery**, **OpenReport**, or **RunDataMacro** method. When the method is run the parameters collection supplies the needed parameters. When the method is finished the parameters collection is cleared.

Because each of the methods that accepts parameters clears the parameters collection when it completes, you must ensure that your calls to  **SetParameter** immediately precede the call to the method that employs them.


## Example

The following code example creates two parameters for use by the AddComment data macro. The two parameters are named prmComment and prmRelatedID, respectively. The value of the txtComment textbox is stored in the prmComment parameter. The value of the txtId textbox is stored in the prmRelatedID parameter.


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

