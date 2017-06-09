---
title: Assign a Control Tip, Caption, or Accelerator Key to a Page or Tab
ms.prod: outlook
ms.assetid: 69ea2e05-fa0e-f4b8-c9fb-52cdbb5c0f71
ms.date: 06/08/2017
---


# Assign a Control Tip, Caption, or Accelerator Key to a Page or Tab

This procedure sets properties on a  [Page](page-object-outlook-forms-script.md) or [Tab](tab-object-outlook-forms-script.md) in a [MultiPage](multipage-object-outlook-forms-script.md) or [TabStrip](tabstrip-object-outlook-forms-script.md) control only.


1. In the Forms Designer, select a page or tab in a  **MultiPage** or **TabStrip** control. For more information, see [How to: Select and Edit a Control Within a Group](select-and-edit-a-control-within-a-group.md) and [How to: Assign a Control Tip, Caption, or Accelerator Key to a Control](assign-a-control-tip-caption-or-accelerator-key-to-a-control.md). 
    
     **Note**  Be sure to select an individual page or tab, and not the corresponding  **MultiPage** or **TabStrip**. When you select a page or tab, a rectangle appears around its caption.
2. Right-click the caption of the selected page or tab, and then click  **Rename**. 
    
3. In the  **Control Tip Text** box, type the text that you want to use as the control tip.
    
4. In the  **Caption** box, type the text that you want to use as the caption.
    
5. In the  **Accelerator Key** box, type a single character from the caption of the control. Note that the selected character is underlined in the control caption.
    

 **Tip**  To assign a control tip for a  **MultiPage** or **TabStrip**, use the  **ControlTipText** property. If you assign a control tip to a **MultiPage** or a **TabStrip**, control tips for the individual page or tab objects within the  **MultiPage** do not appear.

 For more information about the **ControlTipText** property to set for a specific control, see:

- The  **[ControlTipText](page-controltiptext-property-outlook-forms-script.md)** property for the **[Page](page-object-outlook-forms-script.md)** control.
    
- The  **[ControlTipText](tab-controltiptext-property-outlook-forms-script.md)** property for the **[Tab](tab-object-outlook-forms-script.md)** control.
    


