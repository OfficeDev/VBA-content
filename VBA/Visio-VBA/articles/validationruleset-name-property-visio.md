---
title: ValidationRuleSet.Name Property (Visio)
keywords: vis_sdr.chm18213930
f1_keywords:
- vis_sdr.chm18213930
ms.prod: visio
api_name:
- Visio.ValidationRuleSet.Name
ms.assetid: 4b8c8063-debc-a2ef-a9a5-94fa88713858
ms.date: 06/08/2017
---


# ValidationRuleSet.Name Property (Visio)

Specifies the name of the  **[ValidationRuleSet](validationruleset-object-visio.md)** object that appears in the user interface. The default property of the object. Read/write.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **ValidationRuleSet** object.


### Return Value

 **String**


## Remarks

You cannot set the  **Name** property to a value that exceeds 255 characters or to an empty string.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to various Visio objects. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you do not want to change a name each time a solution is localized. Use the  **Name** property to get or set an object's local name. Use the **[NameU](validationruleset-nameu-property-visio.md)** property to get or set its universal name.


