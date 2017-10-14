---
title: TextBox.DecimalPlaces Property (Access)
keywords: vbaac10.chm11045
f1_keywords:
- vbaac10.chm11045
ms.prod: access
api_name:
- Access.TextBox.DecimalPlaces
ms.assetid: cd032c51-34d1-18d3-c378-7473938ec1d7
ms.date: 06/08/2017
---


# TextBox.DecimalPlaces Property (Access)

You can use the  **DecimalPlaces** property to specify the number of decimal places Microsoft Access uses to display numbers. Read/write **Byte**.


## Syntax

 _expression_. **DecimalPlaces**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **DecimalPlaces** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Auto|255|(Default) Numbers appear as specified by the  **Format** property setting.|
|0 to 15|0 to 15|Digits to the right of the decimal separator appear with the specified number of decimal places; digits to the left of the decimal separator appear as specified by the  **Format** property setting.|
You should set the  **DecimalPlaces** property in the table's property sheet. A bound control you create on a form or report inherits the **DecimalPlaces** property set in the field in the underlying table or query, so you won't have to specify the property individually for every bound control you create.

The  **DecimalPlaces** property setting has no effect if the **Format** property is blank or is set to General Number.

The  **DecimalPlaces** property affects only the number of decimal places that display, not how many decimal places are stored. To change the way a number is stored you must change the **FieldSize** property in table Design view.

You can use the  **DecimalPlaces** property to display numbers differently from the **Format** property setting or from the way they are stored. For example, the Currency setting of the **Format** property displays only two decimal places ($5.35). To display Currency numbers with four decimal places (for example, $5.3523), set the **DecimalPlaces** property to 4.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

