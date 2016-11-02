
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
|[Enter](f62c7d3b-c5af-58a5-f65f-fbcafef724f8.md)|
|[Exit](3e78fb94-69d0-0192-d5e9-f14d8bbf8c4e.md)|
|[GotFocus](c0329ab1-bd08-31be-cd57-636540f58539.md)|
|[LostFocus](061c8169-f11a-db5a-3bfe-5f43d1a33a74.md)|
|[Updated](4c7820ba-d712-7ace-483f-8c943eec16f6.md)|

## Methods



|**Name**|
|:-----|
|[Move](8494088f-0c10-6446-e01e-d70680b0597d.md)|
|[Requery](0055d270-ce36-40da-4eaf-7851da6d5dec.md)|
|[SetFocus](bb608976-d178-0e6e-fc8e-b362108314d5.md)|
|[SizeToFit](12d27756-3f97-4856-571d-cf9b811cc1e0.md)|

## Properties



|**Name**|
|:-----|
|[About](39126d90-a587-ef35-83c8-9d94241f2642.md)|
|[Application](54b56ba5-f624-acc4-cab3-1e007a09a890.md)|
|[BorderColor](7fcfa9d0-bb08-0bdf-81c0-5f171b487138.md)|
|[BorderShade](43cf768f-ad41-5729-e5bf-41d445b54efa.md)|
|[BorderStyle](f0cb73d3-1841-031c-5a5f-0e08d90774ee.md)|
|[BorderThemeColorIndex](4e4d6aeb-dd68-b16f-375a-be4c3cf95286.md)|
|[BorderTint](30e29d2f-df31-457d-eb2d-520d6e6bb9a4.md)|
|[BorderWidth](ac847423-f5ad-4d56-655d-25c468f82240.md)|
|[BottomPadding](37fe735a-4fc8-c772-1cc9-a0208b2fe2e4.md)|
|[Cancel](013feb6d-44e9-dbdf-0342-c07ff743f747.md)|
|[Class](c745856b-c447-af0a-ed9e-9945d3917d10.md)|
|[Controls](9e8e9948-94eb-87d3-6917-be95224da5c4.md)|
|[ControlSource](1f773a09-7bcc-45ec-9380-3ab5ee13f024.md)|
|[ControlTipText](40564070-a355-632b-0578-0bd98f1ccc53.md)|
|[ControlType](9160eff6-cf44-d0fb-0ff0-436a6d62b1c6.md)|
|[Custom](9ce0028d-92a7-c113-c4c8-87caab8c4a5b.md)|
|[Default](ffe92e84-4bfa-56a2-298e-00d448f8dc29.md)|
|[DisplayWhen](5d53befd-6210-12b6-7397-2e1eea8bc5d3.md)|
|[Enabled](d84b19c0-173d-ffbd-dfb3-47a47709d130.md)|
|[EventProcPrefix](578dc1f6-0977-e8b9-e96f-ae3408118456.md)|
|[GridlineColor](a07d7fb0-f538-a186-f016-0236a32ab276.md)|
|[GridlineStyleBottom](6cacbac2-3960-3f3e-45a1-d5b0d8fd3ac0.md)|
|[GridlineStyleLeft](594c56fb-d8d5-a9af-dc40-d29a9dffd02d.md)|
|[GridlineStyleRight](1bafb68b-5ab3-f1da-1a48-858829006755.md)|
|[GridlineStyleTop](5d04ce0c-648f-894b-dd67-06fcc9e4afe4.md)|
|[GridlineWidthBottom](b40d8316-64c5-7039-bd72-27faf3ab4caa.md)|
|[GridlineWidthLeft](94a8129a-ff41-f252-6af6-33f9c6dd9eaf.md)|
|[GridlineWidthRight](ac6c59a2-c074-6678-29fc-200ed3e6b6a9.md)|
|[GridlineWidthTop](9cecf573-f2d5-5e5e-e507-1920ede22d0b.md)|
|[Height](1e482282-30f7-139e-dd89-40cf89139a2e.md)|
|[HelpContextId](a96b1981-3366-f6e9-67c6-5276bfc590d9.md)|
|[HorizontalAnchor](1ccbf207-3b60-d7e7-dd69-355c2e3a1a60.md)|
|[InSelection](5b2a7bf0-e779-681f-f748-97798c119c6f.md)|
|[IsVisible](e432566c-c8a9-6d08-4b01-5c5949248ba9.md)|
|[Layout](5954580e-18f6-87c0-107b-902065cebc90.md)|
|[LayoutID](87fab4f4-cd1a-73cd-a36d-d735723c7511.md)|
|[Left](583eb7ee-2df5-aac2-f103-b343a8d315eb.md)|
|[LeftPadding](5beaff4a-d129-6039-4552-3afe589bae03.md)|
|[Locked](e6b42627-6560-2fab-ecb0-e9ff32d3fe4e.md)|
|[Name](927f6470-53d1-c8bf-4bf0-56f0dbec8c7e.md)|
|[Object](d578251a-7768-6843-c9af-77084a26a737.md)|
|[ObjectPalette](d9712689-b62f-9e18-90d8-4e6327e2b2db.md)|
|[ObjectVerbs](fae2e8b8-6326-143f-15cd-ba1f1c541f5d.md)|
|[ObjectVerbsCount](f7c74900-3f0d-b6b1-3606-ca8d206f85b3.md)|
|[OldBorderStyle](e22d9cd8-e155-aaab-35e0-d9edf7811ef3.md)|
|[OldValue](76a696b3-1ffc-d909-e22e-51eb4fc5347f.md)|
|[OLEClass](d9aad7b9-6388-3365-881a-6e42ebebcfd6.md)|
|[OnEnter](c2ca822a-2b67-5b06-0d5c-ff602b21226b.md)|
|[OnExit](a634b83c-fd5a-1277-44b2-d9e2c4b13436.md)|
|[OnGotFocus](75c6d494-5524-f628-5d27-aff11dc9e358.md)|
|[OnLostFocus](5bbe697b-d9e7-a534-d4b2-ec2e05452682.md)|
|[OnUpdated](6cd30c42-d645-6ca8-5c9e-7a5951283fd9.md)|
|[Parent](04bd9bf4-a19e-83c0-b5c5-d78449a22f97.md)|
|[Properties](d2da3527-c234-3c3b-e0ac-45c324c39a1a.md)|
|[RightPadding](eaa9ae99-22f9-f237-da25-9515d3b8d8a6.md)|
|[Section](6969ee7c-8fdb-6e8b-1bc7-b08424a14df9.md)|
|[SpecialEffect](cad6b92e-b927-fa6f-518c-f019dba0f879.md)|
|[TabIndex](2240626f-2aa8-be76-ce5a-b706bcb70de6.md)|
|[TabStop](d1cb97a8-49f8-deb7-66d6-e402da45fc74.md)|
|[Tag](7be610c6-9d2f-4c06-bda7-8de246badf54.md)|
|[Top](a79a5dba-acdc-d17e-76fb-d90629ea84d5.md)|
|[TopPadding](77604178-a2b7-9ad9-2a2d-91d60843c31c.md)|
|[Value](3661428e-b852-e87d-2758-618c293f4c92.md)|
|[VarOleObject](7de5433c-a2da-bb8e-35d2-9c7aae1ff2cd.md)|
|[Verb](dffd74b7-2a69-9b18-dde2-d0fd02754f15.md)|
|[VerticalAnchor](0a4658e3-3406-a9f6-58e8-e284e95fe616.md)|
|[Visible](574563a0-937c-271b-b106-67c9f48a18aa.md)|
|[Width](659a7481-cf4e-1909-38b7-358b59002a83.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)