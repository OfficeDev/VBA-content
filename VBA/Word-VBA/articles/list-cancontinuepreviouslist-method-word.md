---
title: List.CanContinuePreviousList Method (Word)
keywords: vbawd10.chm160563305
f1_keywords:
- vbawd10.chm160563305
ms.prod: word
api_name:
- Word.List.CanContinuePreviousList
ms.assetid: 5e235969-27ee-22eb-61ba-2b52a23447aa
ms.date: 06/08/2017
---


# List.CanContinuePreviousList Method (Word)

Returns a  **[WdContinue](wdcontinue-enumeration-word.md)** constant ( **wdContinueDisabled** , **wdResetList** , or **wdContinueList** ) that indicates whether the formatting from the previous list can be continued.


## Syntax

 _expression_ . **CanContinuePreviousList**( **_ListTemplate_** )

 _expression_ Required. A variable that represents a **[List](list-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ListTemplate_|Required| **[ListTemplate](listtemplate-object-word.md)**|A list template that's been applied to previous paragraphs in the document.|

## Remarks

This method returns the state of the  **Continue previous list** and **Restart numbering** options in the **Bullets and Numbering** dialog box for a specified list format. To change the settings of these options, set the ContinuePreviousList argument of the **ApplyListTemplate** method.


## Example

This example checks to see whether numbering from a previous list is disabled. If it isn't disabled, the current list template is applied with numbering set to continue from the previous list. The selection must be within the second list, or this example creates an error.


```vb
Dim lfTemp As ListFormat 
Dim intContinue As Integer 
 
Set lfTemp = Selection.Range.ListFormat 
 
intContinue = lfTemp.CanContinuePreviousList( _ 
 ListTemplate:=lfTemp.ListTemplate) 
If intContinue <> wdContinueDisabled Then 
 lfTemp.ApplyListTemplate _ 
 ListTemplate:=lfTemp.ListTemplate, _ 
 ContinuePreviousList:=True 
End If
```


## See also


#### Concepts


[List Object](list-object-word.md)

