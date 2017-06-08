---
title: CommandBarControl Members (Office)
ms.prod: office
ms.assetid: 1d2360e4-7511-a3a4-9959-2f7c8282bf99
ms.date: 06/08/2017
---


# CommandBarControl Members (Office)
Represents a command bar control. The  **CommandBarControl** object is a member of the **CommandBarControls** collection. The properties and methods of the **CommandBarControl** object are all shared by the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects.

Represents a command bar control. The  **CommandBarControl** object is a member of the **CommandBarControls** collection. The properties and methods of the **CommandBarControl** object are all shared by the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Copy](commandbarcontrol-copy-method-office.md)|Copies a command bar control to an existing command bar.|
|[Delete](commandbarcontrol-delete-method-office.md)|Deletes the  **CommandBarControl** object from its collection.|
|[Execute](commandbarcontrol-execute-method-office.md)|Runs the procedure or built-in command assigned to the specified  **CommandBarControl** control.|
|[Move](commandbarcontrol-move-method-office.md)|Moves the specified  **CommandBarControl** to an existing command bar.|
|[Reset](commandbarcontrol-reset-method-office.md)|Resets a built-in command bar to its default configuration, or resets a built-in  **CommandBarControl** to its original function and face.|
|[SetFocus](commandbarcontrol-setfocus-method-office.md)|Moves the keyboard focus to the specified CommandBarControl. If the control is disabled or isn't visible, this method will fail.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](commandbarcontrol-application-property-office.md)|Gets an  **Application** object that represents the container application for the **CommandBarControl** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BeginGroup](commandbarcontrol-begingroup-property-office.md)|Gets  **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.|
|[BuiltIn](commandbarcontrol-builtin-property-office.md)|Gets  **True** if the specified command bar control is a built-in control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.|
|[Caption](commandbarcontrol-caption-property-office.md)|Gets or sets the caption text for a command bar control. Read/write.|
|[Creator](commandbarcontrol-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **CommandBarControl** object was created. Read-only.|
|[DescriptionText](commandbarcontrol-descriptiontext-property-office.md)|Gets or sets the description for a command bar control. Read/write.|
|[Enabled](commandbarcontrol-enabled-property-office.md)|Gets or sets a  **Boolean** value specifying if the **CommandBarControl** is enabled. Read/write.|
|[Height](commandbarcontrol-height-property-office.md)|Gets or sets the height of a  **CommandBarControl** control. Read/write.|
|[HelpContextId](commandbarcontrol-helpcontextid-property-office.md)|Gets or sets the Help context Id number for the Help topic attached to the  **CommandBarControl**. Read/write.|
|[HelpFile](commandbarcontrol-helpfile-property-office.md)|Gets or sets the file name for the Help topic attached to the  **CommandBarControl**. Read/write.|
|[Id](commandbarcontrol-id-property-office.md)|Gets the ID for a built-in  **CommandBarControl**. Read-only.|
|[Index](commandbarcontrol-index-property-office.md)|Gets a ** Long** representing the index number for a **CommandBarControl** object in the collection. Read-only.|
|[IsPriorityDropped](commandbarcontrol-isprioritydropped-property-office.md)|Gets  **True** if the control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the **Visible** property). Read-only.|
|[Left](commandbarcontrol-left-property-office.md)|Gets the horizontal position of the specified  **CommandBarControl** (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.|
|[OLEUsage](commandbarcontrol-oleusage-property-office.md)|Gets or sets the OLE client and OLE server roles in which a  **CommandBarControl** will be used when two Microsoft Office applications are merged. Read/write.|
|[OnAction](commandbarcontrol-onaction-property-office.md)|Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a  **CommandBarControl**. Read/write.|
|[Parameter](commandbarcontrol-parameter-property-office.md)|Gets or sets a string that an application can use to execute a command from a  **CommandBarControl**. Read/write.|
|[Parent](commandbarcontrol-parent-property-office.md)|Gets the  **Parent** object for the **CommandBarControl** object. Read-only.|
|[Priority](commandbarcontrol-priority-property-office.md)|Gets or sets the priority of a  **CommandBarControl**. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.|
|[Tag](commandbarcontrol-tag-property-office.md)|Gets or sets information about the  **CommandBarControl**, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.|
|[TooltipText](commandbarcontrol-tooltiptext-property-office.md)|Gets or sets the text displayed in a  **CommandBarControl's** **ScreenTip**. Read/write.|
|[Top](commandbarcontrol-top-property-office.md)|Gets the distance (in pixels) from the top edge of the specified  **CommandBarControl** to the top edge of the screen. Read-only.|
|[Type](commandbarcontrol-type-property-office.md)|Gets the type of  **CommandBarControl**. Read-only.|
|[Visible](commandbarcontrol-visible-property-office.md)|Gets or sets the  **Visible** property of the **CommandBarControl**. **True** if the **CommandBarControl** is visible. Read/write.|
|[Width](commandbarcontrol-width-property-office.md)|Gets or sets the width (in pixels) of the specified  **CommandBarControl**. Read/write.|

