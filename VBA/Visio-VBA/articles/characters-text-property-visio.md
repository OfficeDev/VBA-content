---
title: Characters.Text Property (Visio)
keywords: vis_sdr.chm10214515
f1_keywords:
- vis_sdr.chm10214515
ms.prod: visio
api_name:
- Visio.Characters.Text
ms.assetid: ebfa0548-4150-f6a6-8362-8bd3c2c36f93
ms.date: 06/08/2017
---


# Characters.Text Property (Visio)

Returns the range of text represented by a  **Characters** object, which may be a subset of the shape's text depending on the values of the **Characters** object's **Begin** and **End** properties.Read/write.


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Variant


## Remarks

The text for a  **Characters** object is returned in a **Variant** of type **String** , as opposed to in a **String** . This is typically transparent if you are using Microsoft Visual Basic.

In the text returned by a  **Characters** object, fields are expanded to the number of characters that are visible in the drawing window. For example, if a shape's text contains a field that displays the file name of a drawing, the **Text** property of a **Characters** object returns the expanded file name (provided the **Begin** and **End** properties were not altered).

If a  **Characters** object represents the text of a shape that is a group, it will always return the text of the group.

Objects from other applications and guides don't have a  **Text** property.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters.Text**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the  **Text** property of a **Characters** object.


```vb
 
Public Sub CharactersText_Example()  
 
    Dim vsoOval As Visio.Shape  
    Dim vsoCharacters As Visio.Characters  
 
    'Create a shape and add text. 
    Set vsoOval = ActivePage.DrawOval(2, 5, 5, 7)  
    vsoOval.Text = "Oval Shape"  
 
    'Get a Characters object from the shape. 
    Set vsoCharacters = vsoOval.Characters  
 
    'Get the text from the Characters object. 
    Debug.Print vsoCharacters.Text  
 
End Sub
```


