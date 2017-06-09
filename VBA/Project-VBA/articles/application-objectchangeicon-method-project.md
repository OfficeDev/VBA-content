---
title: Application.ObjectChangeIcon Method (Project)
keywords: vbapj.chm235
f1_keywords:
- vbapj.chm235
ms.prod: project-server
api_name:
- Project.Application.ObjectChangeIcon
ms.assetid: 8153748e-9b46-5d57-eaaf-0f09564c55e4
ms.date: 06/08/2017
---


# Application.ObjectChangeIcon Method (Project)

Displays the  **Change Icon** dialog box to enable changing the icon of an active bitmap or drawing object that is added in a Gantt chart or other view.


## Syntax

 _expression_. **ObjectChangeIcon**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

The bitmap or drawing object must be displayed as an icon and selected. The  **ObjectChangeIcon** method is equivalent to the **Change Icon** command in the **Convert** dialog box. To open the **Convert** dialog box by using the Project user interface, do the following:


1. Open the  **Project Options** dialog box, choose the **Customize Ribbon** tab, and then choose the list of commands not in the Ribbon.
    
2. In the  **Customize the Ribbon** drop-down list, select **Main Tabs**, and then choose  **New Tab**. Rename the tab, for example,  **Old Methods**. 
    
3. Choose  **New Group** to add a group to the **Old Methods** tab. Rename the group, for example, **Objects**.
    
4. Select the  **Objects** group, add the **Object** and **Convert** commands to the group from the list of commands not in the Ribbon, and then choose **OK**.
    
5. On the Gantt chart, choose  **Object** in the **Old Methods** tab. In the **Insert Object** dialog box, choose **Bitmap Image**. You can create a new image or add it from a file. Check  **Display As Icon**.
    
6. Select the bitmap image object on the Gantt chart, and then choose  **Convert** on the **Old Methods** tab of the Ribbon. In the **Convert** dialog box, choose **Change Icon**.
    



