---
title: ListBox.RemoveItem Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 4788abab-0798-f22e-5098-b76bb223f6c3
ms.date: 06/08/2017
---


# ListBox.RemoveItem Method (Outlook Forms Script)

Removes a row from the list in a  **[ListBox](listbox-object-outlook-forms-script.md)**.


## Syntax

 _expression_. **RemoveItem**( **_pvargIndex_**)

 _expression_A variable that represents a  **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pvargIndex|Required| **Variant**|Specifies the row to delete. The number of the first row is 0; the number of the second row is 1, and so on.|

### Return Value

A Boolean that returns  **True** if the method succeeds, **False** otherwise.


## Remarks

This method will not remove a row from the list if the  **ListBox** is data bound.


