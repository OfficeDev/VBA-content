---
title: Unable to unload within this context (Error 365)
keywords: vblr6.chm1117812
f1_keywords:
- vblr6.chm1117812
ms.prod: office
ms.assetid: 845a5c20-95d1-4920-eb1c-df62dbefc97b
ms.date: 06/08/2017
---


# Unable to unload within this context (Error 365)

In some situations you are not allowed to unload a form or a control on a form. This error has the following causes and solutions:



- There is an  **Unload** statement in the Paint event for the form or for a control on the form that has the Paint event. Remove the **Unload** statement from the Paint event.
    
- There is an  **Unload** statement in the Change, Click, or DropDown events of a **ComboBox**. Remove the **Unload** statement from the event.
    
- There is an  **Unload** statement in the Scroll event of an **HScrollBar** or **VScrollBar** control. Remove the **Unload** statement from the event.
    
- There is an  **Unload** statement in the Resize event of a **Data**, **Form**, **MDIForm**, or **PictureBox** control. Remove the **Unload** statement from the event.
    
- There is an  **Unload** statement in the Resize event of an **MDIForm** that is trying to unload an MDI child form. Remove the **Unload** statement from the event.
    
- There is an  **Unload** statement in the RePosition event or Validate event of a **Data** control. Remove the **Unload** statement from the event.
    
- There is an  **Unload** statement in the ObjectMove event of an **OLE Container** control. Remove the **Unload** statement from the event.
    


