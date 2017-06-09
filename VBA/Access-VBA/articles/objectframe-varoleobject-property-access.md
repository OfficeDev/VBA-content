---
title: ObjectFrame.VarOleObject Property (Access)
keywords: vbaac10.chm10954
f1_keywords:
- vbaac10.chm10954
ms.prod: access
api_name:
- Access.ObjectFrame.VarOleObject
ms.assetid: e04e769d-07fb-dacc-aa70-ddd3a064d785
ms.date: 06/08/2017
---


# ObjectFrame.VarOleObject Property (Access)

 Gets a pointer to an **IOLEObject** that represents the memory address of an OLE object. Read-only **Variant**.


## Syntax

 _expression_. **VarOleObject**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks


 **Note**  The  **VarOleObject** property is compatible with 32-bit and 64-bit OLE objects. It replaces the **LpOLEObject**, which was available in earlier versions of Microsoft Access.

You can use the  **VarOleObject** property to refer to the address of the active OLE object. If no object is displayed, the setting is 0.

This property setting is not available in Design view and is read-only in other views.

In the OLE dynamic link libraries (DLLs), many function calls require the address of the OLE object as an argument. You can pass the value specified by the  **VarOleObject** property when you make application programming interface (API) calls to the OLE DLLs.

If an API call is made to a function that performs a callback to the control, the results are unpredictable.


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

