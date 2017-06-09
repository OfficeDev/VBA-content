---
title: TextBox.OnUndo Property (Access)
keywords: vbaac10.chm11146
f1_keywords:
- vbaac10.chm11146
ms.prod: access
api_name:
- Access.TextBox.OnUndo
ms.assetid: fa62ba10-c8e8-f4d4-5d48-ab73c074f2ef
ms.date: 06/08/2017
---


# TextBox.OnUndo Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **Undo** event occurs. Read/write..


## Syntax

 _expression_. **OnUndo**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the Undo event for the specified object, or "= _functionname_()" where  _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **Undo** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).OnUndo = "[Event Procedure]"
```

The following example specifies that when the  **Undo** event occurs in any text box on the first form of the current project, the associated event procedure should run.




```vb
Dim ctlLoop As Control 
 
For Each ctlLoop In Forms(0).Controls 
 If ctlLoop.Type = acTextBox Then 
 ctlLoop.OnUndo = "[Event Procedure]" 
 End If 
Next ctlLoop 

```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

