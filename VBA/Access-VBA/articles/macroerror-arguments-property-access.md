---
title: MacroError.Arguments Property (Access)
keywords: vbaac10.chm14048
f1_keywords:
- vbaac10.chm14048
ms.prod: access
api_name:
- Access.MacroError.Arguments
ms.assetid: 0c5a6589-bd2c-e818-c9b0-5d3bc094c368
ms.date: 06/08/2017
---


# MacroError.Arguments Property (Access)

Gets the arguments specified for the macro action that was executing when an error occurred. Read-only  **String**.


## Syntax

 _expression_. **Arguments**

 _expression_ A variable that represents a **MacroError** object.


## Remarks

When an error occurs in a macro, information about the error is stored in the  **MacroError** object. If you have not used the **OnError** action to suppress error messages, the macro stops and the error information is displayed in a standard error message. However, if you have used the **OnError** action to suppress error messages, you may want to use the information stored in the **MacroError** object in a condition or a custom error message.

After an error has been handled, the information in the  **MacroError** object is out of date, so it is a good idea to clear the object using the **ClearMacroError** action. This resets the error number in the **MacroError** object back to zero, and clears any other information about the error that is stored in the object, such as the error description, macro name, action name, condition, and arguments. This way, you can inspect the **MacroError** object again later to see if another error has occurred.


## See also


#### Concepts


[MacroError Object](macroerror-object-access.md)

