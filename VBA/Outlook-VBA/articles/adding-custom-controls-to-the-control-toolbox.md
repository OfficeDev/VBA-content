---
title: Adding Custom Controls to the Control Toolbox
ms.prod: outlook
ms.assetid: 81b5bba3-076d-4a02-9aa3-034fab9f1e85
ms.date: 06/08/2017
---


# Adding Custom Controls to the Control Toolbox

You can add a modified control (based on modifications made to the advanced properties) to the  [Control Toolbox](show-or-hide-the-control-toolbox.md). You can also add other custom controls to the  **Control Toolbox**, such as ActiveX controls that are not part of Outlook.

You can use a variety of custom controls in Outlook forms, but there are some limitations with Outlook form pages. Form pages support most ActiveX properties and methods but do not support custom event handling. The  **[Click](add-a-click-event-for-a-control-in-a-custom-form-page.md)** event is the only event for which you can write code. To access the methods of an ActiveX control, use the Visual Basic Application (VBA) Object Browser to browse ActiveX control methods.

There are no similar limitations on adding custom controls when using form regions to customize Outlook forms; form regions support the full event model for any control.

For more information, see the following topics:

-  [How to: Add a Modified Control to the Control Toolbox](add-a-modified-control-to-the-control-toolbox.md)
    
-  [Adding Other Custom Controls to the Control Toolbox](add-other-custom-controls-to-the-control-toolbox.md)
    
-  [Form Regions](form-regions.md)
    

