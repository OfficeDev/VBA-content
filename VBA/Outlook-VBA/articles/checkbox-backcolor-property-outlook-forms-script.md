---
title: CheckBox.BackColor Property (Outlook Forms Script)
keywords: olfm10.chm2000770
f1_keywords:
- olfm10.chm2000770
ms.prod: outlook
ms.assetid: c0c3a00c-2679-68fb-6a4e-6f8bb9946694
ms.date: 06/08/2017
---


# CheckBox.BackColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the background color of the object. Read/write.


## Syntax

 _expression_. **BackColor**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```

You can only see the background color of an object if the  **[BackStyle](checkbox-backstyle-property-outlook-forms-script.md)** property is set to 1.


