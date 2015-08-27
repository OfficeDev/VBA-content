
# ApplicationSettings Members (Visio)
Represents various application settings for Microsoft Visio.

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [GetRasterExportResolution](526d2970-006b-6596-bfef-49446dd58610.md)|Returns the raster export resolution settings.|
| [GetRasterExportSize](70591d2c-ac80-5637-996e-3ebef6be0c51.md)|Gets the raster export size.|
| [SetRasterExportResolution](18b28fe1-4460-940c-0de7-566a608a8f04.md)|Specifies the raster export resolution settings.|
| [SetRasterExportSize](763157d2-014b-0aa4-7c55-a0fb71fb5e23.md)|Sets the raster export size.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](5a8f32a8-4e27-1924-8c67-9be08e38ad66.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
| [ApplyBackgroundToDocument](441f1147-7f91-a4ac-f69c-0f1258400499.md)|Determines whether the selected background or border is applied to all pages in the document ( **True**) or only to the current page ( **False**). Read/write.|
| [ApplyThemesOnShapeAdd](c2a83004-852e-83d7-718f-f27f254aae58.md)|Gets or sets whether to apply themes to new shapes when they are added to the drawing page. Read/write.|
| [AsianTextUI](b317afda-5014-6c53-44e1-a713dabee111.md)|Gets whether Asian text is displayed in the Microsoft Visio user interface. Read-only.|
| [BIDITextUI](a358e155-9ba0-42ca-3192-3fc90ee19559.md)|Gets the current setting for display of right-to-left languages. Read-only.|
| [CenterSelectionOnZoom](ad7a3867-7200-84de-a599-7826b55e12d0.md)|Determines whether when the user zooms in, the selection appears in the center of the window. Read/write.|
| [ComplexTextUI](b4ea05ad-ef40-6886-de68-c9bfb6826a88.md)|Gets whether complex text is displayed in the Microsoft Visio user interface. Read-only.|
| [ConnectorSplittingEnabled](f13df7d6-13b6-b39e-1671-a2287505c923.md)|Determines whether connector splitting is enabled in Microsoft Visio. Read/write.|
| [DefaultSaveFormat](892953a8-1e69-000a-3099-c6f4baa69079.md)|Determines the default format for saving Microsoft Visio files. Read/write.|
| [DeleteConnectorsEnabled](adb52279-5837-08be-ce73-231656ef7640.md)|Determines whether connectors are deleted when a shape to which they are connected is deleted. Read/write.|
| [DeveloperMode](db078edb-e8cb-6362-14e1-096186a197f5.md)|Determines if certain user interface functions for the development environment in Microsoft Visio are enabled. Read/write.|
| [DrawingAids](af1a1565-b304-43be-1a56-44d1c9ee6000.md)|Determines whether drawing aids are currently active in Microsoft Visio. Read/write.|
| [DrawingBackgroundColor](c07d8268-d0f6-afc7-8c6f-da16a3f643a0.md)|Determines the background color of the Microsoft Visio drawing window for the current session. Read/write.|
| [DrawingBackgroundColorGradient](3bd4693b-4312-3b99-5f48-a4d7909cf41c.md)|Determines the background gradient color of the Microsoft Visio drawing window for the current session. Read/write. |
| [DrawingPageColor](7ae90e3a-d7ed-588e-2416-eecc3e8619e7.md)|Determines the page color of the Microsoft Visio drawing window for the current session. Read/write. |
| [EnableAutoConnect](9aef5e1c-7f46-0edb-1237-bbb9412a8aa5.md)|Determines whether the  **AutoConnect** feature is enabled in the Microsoft Visio user interface (UI). Read/write.|
| [EnableFormulaAutoComplete](c5860206-378b-1d21-54cc-4fb939daf5ef.md)|Indicates whether ShapeSheet formula AutoComplete is enabled. Read/write.|
| [EnterCommitsText](ba9ce9fa-d224-cdc3-668d-46c1849911c7.md)|Returns or sets a  **Boolean** that determines whether pressing **Enter** commits shape text ( **True**) or writes a new line ( **False**, the default). Read/write.|
| [FreeformDrawingPrecision](3822238b-cd63-1883-88a6-894b289765d7.md)|Determines the margin of error allowed when the  **Freeform** tool is drawing a straight line before it switches to drawing a spline. Read/write.|
| [FreeformDrawingSmoothing](55526b81-324a-8c6f-1654-bf7e1244ccf2.md)|Determines how precisely mouse movements are smoothed when drawing a spline. Read/write.|
| [KanaFindAndReplace](09616d8b-1a81-2c45-c8e5-7b8fa961a4e2.md)|Gets whether additional options specific to Japanese in the  **Find** and **Replace** dialog boxes are available. Read-only.|
| [KashidaTextUI](84270b9c-2ae9-4050-8a68-c04dee0297bf.md)|Gets the current setting for display of Kashida text-justification in certain cursive languages. Read-only.|
| [ObjectType](f3bce466-0497-6744-f96a-5c4bb7afca08.md)|Returns an object's type. Read-only.|
| [RasterExportBackgroundColor](25591439-b332-af75-dec0-562cd261a453.md)|Determines the background color that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.|
| [RasterExportColorFormat](8306b2c1-d0a0-41ae-16de-0deb4d881604.md)|Determines the color format that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a BMP, JPG, PNG, or TIFF file. Read/write.|
| [RasterExportColorReduction](7897f3aa-d7d1-4dcc-d4f1-9c38771c0122.md)|Determines the color reduction that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a BMP, GIF, PNG, or TIFF file. Read/write.|
| [RasterExportDataCompression](cec938db-1368-7c05-a264-b69ae334a249.md)|Determines the data compression algorithm that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a BMP or TIFF file. Read/write.|
| [RasterExportDataFormat](e07c3f2e-469e-33bc-cd6d-0261cf7ec267.md)|Determines whether the exported raster image is interlaced or non-interlaced when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a GIF or PNG file. Read/write.|
| [RasterExportFlip](1aa94fd4-7d2e-a2db-3291-c86ac4e22573.md)|Determines the flip that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.|
| [RasterExportOperation](7f53b4a6-6497-01ca-2219-575065d4c9f4.md)|Determines the export operation that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a JPG file. Read/write.|
| [RasterExportQuality](6864bbfd-bb2d-721f-4146-f66974318929.md)|Determines the export quality that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a JPG file. Read/write.|
| [RasterExportRotation](660b22ff-11b6-bfaf-1949-18e5e9c57d64.md)|Determines the rotation that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.|
| [RasterExportTransparencyColor](39806af2-1bdd-d659-134f-9cd86110e195.md)|Determines the transparency color that is applied to the exported image when you call the  **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a GIF or PNG file. Read/write.|
| [RasterExportUseTransparencyColor](1fd93b1b-8b35-a82a-17f5-0fa2ffa819a7.md)|Determines whether Microsoft Visio applies, to the exported image, the transparency color that is specified in the  **RasterExportTransparencyColor** property when you call the **Export** method of the ** [Master](1a69e4d7-2b72-f712-d36c-c565af64c278.md)**,  ** [Page](7a7f37ab-b448-eb70-b4f1-c185dfbd511e.md)**,  ** [Selection](e5734140-6dbe-7de8-9695-1a22fb4ac628.md)**, or  ** [Shape](da7a8872-4ebb-a607-e0ed-eebf68ff5630.md)** object to export the specified object to a GIF or PNG file. Read/write.|
| [RecentFilesListSize](8057f3d5-ccaf-28a2-9e70-1844f858d51d.md)|Determines the number of entries in the  **Recent Documents** list in the Microsoft Visio user interface. Read/write.|
| [RecentFoldersListSize](http://msdn.microsoft.com/library/e75dae99-a302-12e9-a6fe-b897ea7c343b%28Office.15%29.aspx)||
| [SATextUI](e8bdb2bd-a54b-01e4-8ee7-c3d5c3156854.md)|Gets the current setting for display of South Asian languages. Read-only.|
| [ShowChooseDrawingTypePane](1551b56c-92a1-127b-38d5-28e50af4a0e0.md)|Determines if the  **New** tab appears when the user opens Microsoft Visio. Read/write.|
| [ShowFileOpenWarnings](643daf77-2109-5609-6761-3d3d0be066c0.md)|Determines if warning messages appear when the user attempts to open files in XML format that contain errors, such as invalid XML code. Read/write.|
| [ShowFileSaveWarnings](86a2bb68-da8e-9661-9f03-13806debc29f.md)|Determines if warning messages appear when the user attempts to save in XML format drawings that contain errors, such as invalid XML code. Read/write.|
| [ShowMoreShapeHandlesOnHover](159a3c68-b882-4352-b14d-0a0be2bcdf29.md)|Gets or sets whether to show additional shape handles when the mouse is paused over a shape. Read/write.|
| [ShowShapeSearchPane](41c07355-5ce8-25fb-ff34-75c24c6c1e89.md)|Gets or sets whether the  **Shape Search** pane is visible in the Microsoft Visio user interface (UI). Read/write.|
| [ShowSmartTags](36df74db-1d60-7bcc-0f52-ed12084a383a.md)|Determines whether display of smart tags in Microsoft Visio is enabled. Read/write.|
| [SnapStrengthExtensionsX](45fb7005-34af-860f-ea59-a48e5a0b7a01.md)|Specifies the distance in pixels along the  _x_-axis that shape extension lines pull when snapping is enabled. Read/Write.|
| [SnapStrengthExtensionsY](01540007-8cbb-e551-6917-85295c99185a.md)|Specifies the distance in pixels along the  _y-_axis that shape extension lines pull when snapping is enabled. Read/write.|
| [SnapStrengthGeometryX](8b0b9a83-fbbb-46f0-445d-35fa429a1e11.md)|Specifies the distance in pixels along the  _x_-axis that shape geometry pulls when snapping is enabled. Read/write.|
| [SnapStrengthGeometryY](8e5b3bf3-4cb6-af1c-1812-863c247608b9.md)|Specifies the distance in pixels along the  _y_-axis that shape geometry pulls when snapping is enabled. Read/write.|
| [SnapStrengthGridX](ebe2489d-6643-4303-30fd-720446a4e19d.md)|Specifies the distance in pixels along the x-axis that gridlines pull when snapping is enabled. Read/write.|
| [SnapStrengthGridY](0fc60e09-0315-d981-7375-9c5fd71ec6bd.md)|Specifies the distance in pixels along the y-axis that gridlines pull when snapping is enabled. Read/write.|
| [SnapStrengthGuidesX](d4a8fcca-1aee-c093-c92f-6a3ba2a6b319.md)|Specifies the distance in pixels along the x-axis that guides pull when snapping is enabled. Read/write.|
| [SnapStrengthGuidesY](64d2c688-d900-c5e7-28c7-a0c24dcc854a.md)|Specifies the distance in pixels along the y-axis that guides pull when snapping is enabled. Read/write.|
| [SnapStrengthPointsX](7f18b1bc-0164-48d5-b50c-d269b68c1f31.md)|Specifies the distance in pixels along the x-axis that points pull when snapping is enabled. Read/write.|
| [SnapStrengthPointsY](7719694e-993a-2792-3f6f-3d697ef34790.md)|Specifies the distance in pixels along the y-axis that points pull when snapping is enabled. Read/write.|
| [SnapStrengthRulerX](594b4730-94ac-de20-12df-97ae0df4b7f6.md)|Specifies the distance in pixels along the x-axis that rulers pull when snapping is enabled. Read/write.|
| [SnapStrengthRulerY](b0b6a3da-a87d-496e-901c-e6850e6c612b.md)|Specifies the distance in pixels along the y-axis that rulers pull when snapping is enabled. Read/write.|
| [Stat](dd322ca5-6f48-94ab-8632-f60896dd3228.md)|Returns status information for an object. Read-only.|
| [StencilBackgroundColor](a1cbf151-96b8-7c9b-9ceb-2cf5768d41ff.md)|Determines the background color of the Microsoft Visio stencil window for the current session. Read/write.|
| [StencilBackgroundColorGradient](e73b2f5a-6ddf-0e46-62f1-8409e7e0608c.md)|Determines the background gradient color of the Microsoft Visio stencil window for the current session. Read/write. |
| [StencilCharactersPerLine](e69c1c58-6383-f614-fcd4-d32505f53206.md)|For shapes on stencils, determines approximately how many characters of each shape's name appear on each line before the text wraps to the next line. Read/write.|
| [StencilLinesPerMaster](0d962d29-2cb5-5a9f-342f-1a35905a3438.md)|For shapes on stencils in Microsoft Visio, determines how many lines of text of each shape's name can appear below the shape before the text is truncated and "..." is appended. Read/write.|
| [StencilTextColor](4e71f784-0d1a-c49f-7e9f-e0b96fdc0f6e.md)|Determines the color of text in stencil windows in Microsoft Visio for the current session. Read/write.|
| [SVGExportFormat](9e7ca1cb-5ace-b75b-0e59-61566b9a0169.md)|Returns or sets a  [VisSVGExportFormat](d8ca8c3f-41d9-4e9d-8f6d-f5567361b14e.md) constant that specifies the options for saving a Microsoft Visio file in SVG format. Read/write.|
| [TransitionsEnabled](af3b25b8-eee2-110f-9189-5133144d3a43.md)|Determines whether Microsoft Visio uses an animated transition to show certain shape movements, such as re-layout of shapes. Read/write.|
| [UndoLevels](5d4ad370-254d-3b99-21d9-2cbdf60842a6.md)|Determines the number of consecutive actions the user can undo in Microsoft Visio. Read/write.|
| [UseLocalUserInfo](http://msdn.microsoft.com/library/db86193c-0af3-2cd5-359c-193afa23c17f%28Office.15%29.aspx)||
| [UserInitials](cb134c38-52aa-53ef-cfc8-708d3b4a7887.md)|Determines the user initials associated with the Microsoft Visio file. Read/write.|
| [UserName](a22d5f51-15a4-6a89-4c3a-2e96f9cf7b4e.md)|Gets or sets the user name of an  **Application** object. Read/write.|
| [ZoomOnRoll](27475650-3703-4a95-f71c-d979ba2066f6.md)|Determines whether zooming in to and out from a Microsoft Visio drawing by rolling the wheel of the mouse is enabled. Read/write.|
