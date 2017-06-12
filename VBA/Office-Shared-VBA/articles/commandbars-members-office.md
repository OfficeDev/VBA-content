---
title: CommandBars Members (Office)
ms.prod: office
ms.assetid: c11db22d-b7bb-20a2-a455-e441cb8d5bc0
ms.date: 06/08/2017
---


# CommandBars Members (Office)
A collection of  **CommandBar** objects that represent the command bars in the container application.

A collection of  **CommandBar** objects that represent the command bars in the container application.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[OnUpdate](commandbars-onupdate-event-office.md)|Occurs when any change is made to a command bar.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](commandbars-add-method-office.md)|Creates a new command bar and adds it to the collection of command bars.|
|[CommitRenderingTransaction](commandbars-commitrenderingtransaction-method-office.md)|Commits the rendering transaction. Returns  **Nothing**.|
|[ExecuteMso](commandbars-executemso-method-office.md)|Executes the control identified by the  **idMso** parameter.|
|[FindControl](commandbars-findcontrol-method-office.md)|Gets a  **CommandBarControl** object that fits a specified criteria.|
|[FindControls](commandbars-findcontrols-method-office.md)|Gets the  **CommandBarControls** collection that fits the specified criteria.|
|[GetEnabledMso](commandbars-getenabledmso-method-office.md)|Returns True if the control identified by the  **idMso** parameter is enabled.|
|[GetImageMso](commandbars-getimagemso-method-office.md)|Returns an  **IPictureDisp** object of the control image identified by the **idMso** parameter scaled to the dimensions specified by width and height.|
|[GetLabelMso](commandbars-getlabelmso-method-office.md)|Returns the label of the control identified by the  **idMso** parameter as a String.|
|[GetPressedMso](commandbars-getpressedmso-method-office.md)|Returns a value indicating whether the toggleButton control identified by the  **idMso** parameter is pressed.|
|[GetScreentipMso](commandbars-getscreentipmso-method-office.md)|Returns the screentip of the control identified by the  **idMso** parameter as a String.|
|[GetSupertipMso](commandbars-getsupertipmso-method-office.md)|Returns the supertip of the control identified by the  **idMso** parameter as a String.|
|[GetVisibleMso](commandbars-getvisiblemso-method-office.md)|Returns True if the control identified by the  **idMso** parameter is visible.|
|[ReleaseFocus](commandbars-releasefocus-method-office.md)|Releases the user interface focus from all command bars.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActionControl](commandbars-actioncontrol-property-office.md)|Gets the  **CommandBarControl** object whose **OnAction** property is set to the running procedure. Read-only.|
|[ActiveMenuBar](commandbars-activemenubar-property-office.md)|Gets a  **CommandBar** object that represents the active menu bar in the container application. Read-only.|
|[AdaptiveMenus](commandbars-adaptivemenus-property-office.md)|This property checks or unchecks the check box control for the option to show menus in Microsoft Office as full or personalized. Read/write.|
|[Application](commandbars-application-property-office.md)|Gets an  **Application** object that represents the container application for the **CommandBars** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Count](commandbars-count-property-office.md)|Gets a count of the number of command bars in the host application. Read-only.|
|[Creator](commandbars-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **CommandBars** object was created. Read-only.|
|[DisableAskAQuestionDropdown](commandbars-disableaskaquestiondropdown-property-office.md)|Is  **True** if the **Answer Wizard** dropdown menu is enabled. Read/write.|
|[DisableCustomize](commandbars-disablecustomize-property-office.md)|Is  **True** if toolbar customization is disabled. Read/write.|
|[DisplayFonts](commandbars-displayfonts-property-office.md)|Is  **True** if the font names in the **Font** box are displayed in their actual fonts. Read/write.|
|[DisplayKeysInTooltips](commandbars-displaykeysintooltips-property-office.md)|Is  **True** if shortcut keys are displayed in the **ToolTips** for each command bar control. Read/write.|
|[DisplayTooltips](commandbars-displaytooltips-property-office.md)|Is  **True** if ScreenTips are displayed whenever the user positions the pointer over command bar controls. Read/write.|
|[Item](commandbars-item-property-office.md)|Gets a  **CommandBar** object from the **CommandBars** collection. Read-only.|
|[LargeButtons](commandbars-largebuttons-property-office.md)|Is  **True** if the toolbar buttons displayed are larger than normal size. Read/write.|
|[MenuAnimationStyle](commandbars-menuanimationstyle-property-office.md)|Gets or sets a  **MsoMenuAnimation** that represents the way a command bar is animated. Read/write.|
|[Parent](commandbars-parent-property-office.md)|Gets the  **Parent** object for the **CommandBars** object. Read-only.|

