---
title: CommandBarButton Members (Office)
ms.prod: office
ms.assetid: 69fe57fe-dabc-9379-283c-d0a51a775592
ms.date: 06/08/2017
---


# CommandBarButton Members (Office)
Represents a button control on a command bar.

Represents a button control on a command bar.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Click](commandbarbutton-click-event-office.md)|Occurs when the user clicks a  **CommandBarButton** object.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Copy](commandbarbutton-copy-method-office.md)|Copies a command bar button control to an existing command bar.|
|[CopyFace](commandbarbutton-copyface-method-office.md)|Copies the face of a command bar button control to the Clipboard.|
|[Delete](commandbarbutton-delete-method-office.md)|Deletes the  **CommandBarButton** object from its collection.|
|[Execute](commandbarbutton-execute-method-office.md)|Runs the procedure or built-in command assigned to the specified  **CommandBarButton** control.|
|[Move](commandbarbutton-move-method-office.md)|Moves the specified  **CommandBarButton** control to an existing command bar.|
|[PasteFace](commandbarbutton-pasteface-method-office.md)|Pastes the contents of the Clipboard onto a  **CommandBarButton**.|
|[Reset](commandbarbutton-reset-method-office.md)|Resets a built-in  **CommandBarButton** control to its original function and face.|
|[SetFocus](commandbarbutton-setfocus-method-office.md)|Moves the keyboard focus to the specified  **CommandBarButton** control. If the button is disabled or isn't visible, this method will fail.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](commandbarbutton-application-property-office.md)|Gets an  **Application** object that represents the container application for the **CommandBarButton** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BeginGroup](commandbarbutton-begingroup-property-office.md)|Gets True if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.|
|[BuiltIn](commandbarbutton-builtin-property-office.md)|Is  **True** if the specified command bar control is a control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.|
|[BuiltInFace](commandbarbutton-builtinface-property-office.md)|Is  **True** if the face of a command bar button control is its original built-in face. Read/write.|
|[Caption](commandbarbutton-caption-property-office.md)|Gets or sets the caption text for a command bar control. Read/write.|
|[Creator](commandbarbutton-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **CommandBarButton** object was created. Read-only.|
|[DescriptionText](commandbarbutton-descriptiontext-property-office.md)|Gets or sets the description for a command bar button control. Read/write.|
|[Enabled](commandbarbutton-enabled-property-office.md)|**True** if the specified **CommandBar** or **CommandBarControl** is enabled. Read/write .|
|[FaceId](commandbarbutton-faceid-property-office.md)|Gets or sets the Id number for the face of a  **CommandBarButton** control. Read/write.|
|[Height](commandbarbutton-height-property-office.md)|Gets or sets the height of a command bar control. Read/write.|
|[HelpContextId](commandbarbutton-helpcontextid-property-office.md)|Gets or sets the Help context Id number for the Help topic attached to the  **CommandBarButton** control. Read/write.|
|[HelpFile](commandbarbutton-helpfile-property-office.md)|Gets or sets the file name for the Help topic attached to the  **CommandBarButton** control. Read/write.|
|[HyperlinkType](commandbarbutton-hyperlinktype-property-office.md)|Sets or gets a  **MsoCommandBarButtonHyperlinkType** constant that represents the type of hyperlink associated with the specified command bar button. Read/write.|
|[Id](commandbarbutton-id-property-office.md)|Gets the ID for a built-in  **CommandBarButton** control. Read-only.|
|[Index](commandbarbutton-index-property-office.md)|Gets a  **Long** representing the index number for an **CommandBarButton** object in the collection. Read-only.|
|[IsPriorityDropped](commandbarbutton-isprioritydropped-property-office.md)|Gets  **True** if the **CommandBarButton** control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the Visible property). Read-only.|
|[Left](commandbarbutton-left-property-office.md)|Set or gets the horizontal position of the specified  **CommandBarButton** control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.|
|[Mask](commandbarbutton-mask-property-office.md)|Gets or sets an  **IPictureDisp** object representing the mask image of a **CommandBarButton** object. The mask image determines what parts of the button image are transparent. Read/write.|
|[OLEUsage](commandbarbutton-oleusage-property-office.md)|Gets or sets the OLE client and OLE server roles in which a  **CommandBarButton** control will be used when two Microsoft Office applications are merged. Read/write.|
|[OnAction](commandbarbutton-onaction-property-office.md)|Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a  **CommandBarButton** control. Read/write.|
|[Parameter](commandbarbutton-parameter-property-office.md)|Gets or sets a string that an application can use to execute a command from a  **CommandBarButton** control. Read/write.|
|[Parent](commandbarbutton-parent-property-office.md)|Gets the  **Parent** object for the **CommandBarButton** object. Read-only.|
|[Picture](commandbarbutton-picture-property-office.md)|Gets or sets an  **IPictureDisp** object representing the image of a **CommandBarButton** object. Read/write.|
|[Priority](commandbarbutton-priority-property-office.md)|Gets or sets the priority of a CommandBarButton control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.|
|[ShortcutText](commandbarbutton-shortcuttext-property-office.md)|Gets or sets the shortcut key text displayed next to a  **CommandBarButton** control when the button appears on a menu, submenu, or shortcut menu. Read/write.|
|[State](commandbarbutton-state-property-office.md)|Gets or sets the appearance of a CommandBarButton control. Read/write.|
|[Style](commandbarbutton-style-property-office.md)|Gets or sets the way a  **CommandBarButton** control is displayed. Read/write.|
|[Tag](commandbarbutton-tag-property-office.md)|Gets or sets information about the  **CommandBarButton** control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.|
|[TooltipText](commandbarbutton-tooltiptext-property-office.md)|Gets or sets the text displayed in a  **CommandBarButton's** **ScreenTip**. Read/write.|
|[Top](commandbarbutton-top-property-office.md)|Gets the distance (in pixels) from the top edge of the specified  **CommandBarButton** control to the top edge of the screen. Read-only.|
|[Type](commandbarbutton-type-property-office.md)|Gets the type of  **CommandBarButton** control. Read-only.|
|[Visible](commandbarbutton-visible-property-office.md)|Gets or sets the  **Visible** property of the **CommandBarButton** control. **True** if the **CommandBarButton** is visible. Read/write.|
|[Width](commandbarbutton-width-property-office.md)|Gets or sets the width (in pixels) of the specified  **CommandBarButton** control. Read/write.|

