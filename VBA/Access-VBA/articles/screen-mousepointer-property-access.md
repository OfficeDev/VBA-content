---
title: Screen.MousePointer Property (Access)
keywords: vbaac10.chm12492
f1_keywords:
- vbaac10.chm12492
ms.prod: access
api_name:
- Access.Screen.MousePointer
ms.assetid: e7ee88cf-7eb8-a447-d671-1549cdbcb4fd
ms.date: 06/08/2017
---


# Screen.MousePointer Property (Access)

You can use the  **MousePointer** property together with the **[Screen](screen-object-access.md)** object to specify or determine the type of mouse pointer currently displayed. Read/write **Integer**.


## Syntax

 _expression_. **MousePointer**

 _expression_ A variable that represents a **Screen** object.


## Remarks

The setting for the  **MousePointer** property is an **Integer** value representing one of the following pointers.



|**Setting**|**Description**|
|:-----|:-----|
|0|(Default) The shape is determined by Microsoft Access|
|1|Normal Select (Arrow)|
|3|Text Select (I-Beam)|
|7|Vertical Resize (Size N, S)|
|9|Horizontal Resize (Size E, W)|
|11|Busy (Hourglass)|

 **Note**  Setting the  **MousePointer** property to an integer other than one that appears in the preceding table will cause the property to be set to 0.

The  **MousePointer** property affects the appearance of the mouse pointer over the entire screen. Some custom controls have a **MousePointer** property that, if set, will specify how the mouse pointer is displayed when it's positioned over the control.

You could use the  **MousePointer** property to indicate that your application is busy by setting the property to 11 to display an hourglass icon. You can also read the **MousePointer** property to determine what's being displayed. This could be useful if you wanted to prevent a user from clicking a command button while the mouse pointer is displaying an hourglass icon.

Setting the  **MousePointer** property to 11 is the same as passing the **True** (?1) argument to the **[Hourglass](docmd-hourglass-method-access.md)** method of the **[DoCmd](docmd-object-access.md)** object. Conversely, passing the **True** argument to the **Hourglass** method also sets the **MousePointer** property to 11.


## Example

The following example changes the mouse pointer to an hourglass.


```vb
Screen.MousePointer = 11
```


## See also


#### Concepts


[Screen Object](screen-object-access.md)

