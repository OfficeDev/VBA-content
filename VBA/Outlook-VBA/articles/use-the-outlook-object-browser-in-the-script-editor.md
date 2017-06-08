---
title: Use the Outlook Object Browser in the Script Editor
keywords: olfm10.chm3077348
f1_keywords:
- olfm10.chm3077348
ms.prod: outlook
ms.assetid: 0b201674-66b4-38e2-fb67-74f6c56d447b
ms.date: 06/08/2017
---


# Use the Outlook Object Browser in the Script Editor

## To view the Outlook object browser


1.  On the **Developer** tab, in the **Custom Forms** group, click **Design a Form**.
    
     **Note**  If you do not see the  **Developer** tab, see [Run in Developer Mode in Outlook](run-in-developer-mode-in-outlook.md) to activate the **Developer** tab.
2. In the  **Design Form** dialog box, select the form that you would like to use the object browser with, and click **Open** to open the form in design mode.
    
3. On the  **Developer** tab, in the **Form** group, click **View Code** to open the selected form in the Script Editor.
    
4. In the Script Editor, click  **Object Browser** on the **Script** menu or press **F2**.
    
All of the available Outlook objects are listed in the  **Classes** pane of the object browser in alphabetical order.

To view the members of an object, select the object in the  **Classes** pane. The members of this object appear in alphabetical order in the **Members of** pane. The heading at the top of this pane will reflect the name of the object that you select. For example, if you select the **AppointmentItem** object in the **Classes** pane, the heading of the **Members of** pane will appear as **Members of AppointmentItem**.

The details pane shows the definition of the selected member. This text is read-only and cannot be copied and pasted into the Script Editor.


## To insert an item from the object browser into the Script Editor


1.  In the Script Editor, position your cursor at the location for insertion.
    
2. Select the desired object in the  **Classes** pane.
    
3. Select the desired member of this object in the  **Members of** pane.
    
4. Click  **Insert**.
    

 **Note**  The  **Insert** button remains unavailable until a member of the object is selected.


