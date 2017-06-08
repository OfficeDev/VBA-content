---
title: CustomControl Object (Access)
keywords: vbaac10.chm12062
f1_keywords:
- vbaac10.chm12062
ms.prod: access
api_name:
- Access.CustomControl
ms.assetid: a6ded8cf-4cf8-26ff-bade-f37a7ac52b02
ms.date: 06/08/2017
---


# CustomControl Object (Access)

When setting the properties of an ActiveX control, you may need or prefer to use the control's custom properties dialog box. This custom properties dialog box provides an alternative to the list of properties in the Microsoft Access property sheet for setting ActiveX control properties in Design view.


## Remarks


 **Note**  This information only applies to ActiveX controls in a Microsoft Access database environment.

 **Two Ways to Set Properties**

The reason for the custom properties dialog box is that not all applications that use ActiveX controls provide a property sheet like the one in Microsoft Access. The custom properties dialog box provides an interface for setting key control properties regardless of the interface supplied by the hosting application.

For some ActiveX control properties, you can choose either of these two locations to set the property:


- The Microsoft Access property sheet.
    
- The ActiveX control's custom properties dialog box.
    
In some cases, the custom properties dialog box is the only way to set a property in Design view. This is usually the situation when the interface needed to set a property doesn't work inside the Microsoft Access property sheet. For example, the  **GridFont** property for the Calendar control has a number of arguments; you can't set more than one argument per property in the Microsoft Access property sheet.

 **Finding the Custom Properties Dialog Box**

Not all ActiveX controls provide a custom properties dialog box. To see whether a control provides this custom properties dialog box, look for the  **Custom** property in the Microsoft Access property sheet for this control. If the list of properties contains the name **Custom**, then the control provides the custom properties dialog box.

After you click the  **Custom** property box in the Microsoft Access property sheet, click the **Build** button to the right of the property box to display the control's custom properties dialog box, often presented as a tabbed dialog box. Choose the tab that contains the interface for setting the properties that you want to set.

 **Using the Custom Properties Dialog Box**

After you make changes on one tab, you can often apply those changes immediately by clicking the  **Apply** button (if provided). You can click other tabs to set other properties as needed. To approve all changes made in the custom properties dialog box, click the **OK** button. To return to the Microsoft Access property sheet without changing any property settings, click the **Cancel** button.

You can also view the custom properties dialog box by clicking the  **Properties** subcommand of the ActiveX control **Object** command (for example, **Calendar Control Object** ) on the **Edit** menu, or by clicking this same subcommand on the shortcut menu for the ActiveX control. In addition, some properties in the Microsoft Access property sheet for the ActiveX control, like the **GridFontColor** property of the Calendar control, have a **Build** button to the right of the property box. When you click the **Build** button, the custom properties dialog box is displayed, with the appropriate tab selected (for example, **Colors** ).


## Events



|**Name**|
|:-----|
|[Enter](customcontrol-enter-event-access.md)|
|[Exit](customcontrol-exit-event-access.md)|
|[GotFocus](customcontrol-gotfocus-event-access.md)|
|[LostFocus](customcontrol-lostfocus-event-access.md)|
|[Updated](customcontrol-updated-event-access.md)|

## Methods



|**Name**|
|:-----|
|[Move](customcontrol-move-method-access.md)|
|[Requery](customcontrol-requery-method-access.md)|
|[SetFocus](customcontrol-setfocus-method-access.md)|
|[SizeToFit](customcontrol-sizetofit-method-access.md)|

## Properties



|**Name**|
|:-----|
|[About](customcontrol-about-property-access.md)|
|[Application](customcontrol-application-property-access.md)|
|[BorderColor](customcontrol-bordercolor-property-access.md)|
|[BorderShade](customcontrol-bordershade-property-access.md)|
|[BorderStyle](customcontrol-borderstyle-property-access.md)|
|[BorderThemeColorIndex](customcontrol-borderthemecolorindex-property-access.md)|
|[BorderTint](customcontrol-bordertint-property-access.md)|
|[BorderWidth](customcontrol-borderwidth-property-access.md)|
|[BottomPadding](customcontrol-bottompadding-property-access.md)|
|[Cancel](customcontrol-cancel-property-access.md)|
|[Class](customcontrol-class-property-access.md)|
|[Controls](customcontrol-controls-property-access.md)|
|[ControlSource](customcontrol-controlsource-property-access.md)|
|[ControlTipText](customcontrol-controltiptext-property-access.md)|
|[ControlType](customcontrol-controltype-property-access.md)|
|[Custom](customcontrol-custom-property-access.md)|
|[Default](customcontrol-default-property-access.md)|
|[DisplayWhen](customcontrol-displaywhen-property-access.md)|
|[Enabled](customcontrol-enabled-property-access.md)|
|[EventProcPrefix](customcontrol-eventprocprefix-property-access.md)|
|[GridlineColor](customcontrol-gridlinecolor-property-access.md)|
|[GridlineStyleBottom](customcontrol-gridlinestylebottom-property-access.md)|
|[GridlineStyleLeft](customcontrol-gridlinestyleleft-property-access.md)|
|[GridlineStyleRight](customcontrol-gridlinestyleright-property-access.md)|
|[GridlineStyleTop](customcontrol-gridlinestyletop-property-access.md)|
|[GridlineWidthBottom](customcontrol-gridlinewidthbottom-property-access.md)|
|[GridlineWidthLeft](customcontrol-gridlinewidthleft-property-access.md)|
|[GridlineWidthRight](customcontrol-gridlinewidthright-property-access.md)|
|[GridlineWidthTop](customcontrol-gridlinewidthtop-property-access.md)|
|[Height](customcontrol-height-property-access.md)|
|[HelpContextId](customcontrol-helpcontextid-property-access.md)|
|[HorizontalAnchor](customcontrol-horizontalanchor-property-access.md)|
|[InSelection](customcontrol-inselection-property-access.md)|
|[IsVisible](customcontrol-isvisible-property-access.md)|
|[Layout](customcontrol-layout-property-access.md)|
|[LayoutID](customcontrol-layoutid-property-access.md)|
|[Left](customcontrol-left-property-access.md)|
|[LeftPadding](customcontrol-leftpadding-property-access.md)|
|[Locked](customcontrol-locked-property-access.md)|
|[Name](customcontrol-name-property-access.md)|
|[Object](customcontrol-object-property-access.md)|
|[ObjectPalette](customcontrol-objectpalette-property-access.md)|
|[ObjectVerbs](customcontrol-objectverbs-property-access.md)|
|[ObjectVerbsCount](customcontrol-objectverbscount-property-access.md)|
|[OldBorderStyle](customcontrol-oldborderstyle-property-access.md)|
|[OldValue](customcontrol-oldvalue-property-access.md)|
|[OLEClass](customcontrol-oleclass-property-access.md)|
|[OnEnter](customcontrol-onenter-property-access.md)|
|[OnExit](customcontrol-onexit-property-access.md)|
|[OnGotFocus](customcontrol-ongotfocus-property-access.md)|
|[OnLostFocus](customcontrol-onlostfocus-property-access.md)|
|[OnUpdated](customcontrol-onupdated-property-access.md)|
|[Parent](customcontrol-parent-property-access.md)|
|[Properties](customcontrol-properties-property-access.md)|
|[RightPadding](customcontrol-rightpadding-property-access.md)|
|[Section](customcontrol-section-property-access.md)|
|[SpecialEffect](customcontrol-specialeffect-property-access.md)|
|[TabIndex](customcontrol-tabindex-property-access.md)|
|[TabStop](customcontrol-tabstop-property-access.md)|
|[Tag](customcontrol-tag-property-access.md)|
|[Top](customcontrol-top-property-access.md)|
|[TopPadding](customcontrol-toppadding-property-access.md)|
|[Value](customcontrol-value-property-access.md)|
|[VarOleObject](customcontrol-varoleobject-property-access.md)|
|[Verb](customcontrol-verb-property-access.md)|
|[VerticalAnchor](customcontrol-verticalanchor-property-access.md)|
|[Visible](customcontrol-visible-property-access.md)|
|[Width](customcontrol-width-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
