---
title: Inspectors Object (Outlook)
keywords: vbaol11.chm2996
f1_keywords:
- vbaol11.chm2996
ms.prod: outlook
api_name:
- Outlook.Inspectors
ms.assetid: b65475d6-a212-fc96-459d-47390dfe5ee5
ms.date: 06/08/2017
---


# Inspectors Object (Outlook)

Contains a set of  **[Inspector](inspector-object-outlook.md)** objects representing all inspectors.


## Remarks

 An inspector need not be visible to be included in the collection.

Use the  **[Inspectors](application-inspectors-property-outlook.md)** property to return the **Inspectors** object from the **[Application](application-object-outlook.md)** object.


## Example

The following example shows how to retrieve the  **Inspectors** object in Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA).


```
Set myInspectors = Application.Inspectors
```


## Events



|**Name**|
|:-----|
|[NewInspector](inspectors-newinspector-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Add](inspectors-add-method-outlook.md)|
|[Item](inspectors-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](inspectors-application-property-outlook.md)|
|[Class](inspectors-class-property-outlook.md)|
|[Count](inspectors-count-property-outlook.md)|
|[Parent](inspectors-parent-property-outlook.md)|
|[Session](inspectors-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
