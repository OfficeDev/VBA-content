---
title: NavigationFolder.Position Property (Outlook)
keywords: vbaol11.chm2907
f1_keywords:
- vbaol11.chm2907
ms.prod: outlook
api_name:
- Outlook.NavigationFolder.Position
ms.assetid: cfa86104-c191-51f8-4da3-dc3c26d6a7ed
ms.date: 06/08/2017
---


# NavigationFolder.Position Property (Outlook)

Returns or sets an  **Long** value that represents the ordinal position of the **[NavigationFolder](navigationfolder-object-outlook.md)** object when displayed in the Navigation Pane. Read/write.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **NavigationFolder** object.


## Remarks

This property can only be set to a value between 1 and the value of the  **[Count](navigationfolders-count-property-outlook.md)** property for the parent **[NavigationFolders](navigationfolders-object-outlook.md)** object. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a  **NavigationFolder** object changes the **Position** values of other navigation folders contained by a **NavigationFolders** collection, depending on the relative change between the new value and the original value of the **Position** property for that **NavigationFolder** object:


- If the new value is less than the original value, then the specified  **NavigationFolder** object moves up to the new position and pushes the other navigation folders already at or below that new position down.
    
- If the new value is greater than the original value, then the specified  **NavigationFolder** object moves down to the new position and pushes the other navigation folders between the old position and the new position up, filling the old position.
    
If the navigation folder has been removed from the Navigation Pane, then this property returns -1 to indicate that the navigation folder is no longer part of the navigation group.


## See also


#### Concepts


[NavigationFolder Object](navigationfolder-object-outlook.md)

