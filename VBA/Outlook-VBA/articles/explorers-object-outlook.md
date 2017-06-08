---
title: Explorers Object (Outlook)
keywords: vbaol11.chm2995
f1_keywords:
- vbaol11.chm2995
ms.prod: outlook
api_name:
- Outlook.Explorers
ms.assetid: 8398532a-1fad-7390-6778-109ac5e6c67c
ms.date: 06/08/2017
---


# Explorers Object (Outlook)

Contains a set of  **[Explorer](explorer-object-outlook.md)** objects representing all explorers.


## Remarks

 An explorer need not be visible to be included in the **Explorers** collection.

Use the  **[Explorers](application-explorers-property-outlook.md)** property to return the **Explorers** object from the **[Application](application-object-outlook.md)** object.


## Example

The following example shows how to retrieve the  **Explorers** object in Microsoft Visual Basic and Microsoft Visual Basic for Applications (VBA).


```
Set myExplorers = Application.Explorers
```


## Events



|**Name**|
|:-----|
|[NewExplorer](explorers-newexplorer-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Add](explorers-add-method-outlook.md)|
|[Item](explorers-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](explorers-application-property-outlook.md)|
|[Class](explorers-class-property-outlook.md)|
|[Count](explorers-count-property-outlook.md)|
|[Parent](explorers-parent-property-outlook.md)|
|[Session](explorers-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
