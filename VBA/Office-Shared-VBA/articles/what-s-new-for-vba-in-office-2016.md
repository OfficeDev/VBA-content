---
title: What's New for VBA in Office 2016
ms.prod: office
ms.assetid: c0294abb-bc0e-495d-b387-4398378dd3ad
ms.date: 06/08/2017
---


# What's New for VBA in Office 2016
The following tables summarize the new VBA language updates for Office 2016.




## Access





|**Name**|**Description**|
|:-----|:-----|
|**[CodeProject.IsSQLBackend Property (Access)](http://msdn.microsoft.com/library/c0b0f9bb-5ad4-69c1-9553-2caf420870f1%28Office.15%29.aspx)**|Returns the  **Boolean** value **true** if the code project was created in Access 2013 and newer, and **false** if otherwise.|
|**[CurrentProject.IsSQLBackend Property (Access)](http://msdn.microsoft.com/library/39e312e0-9b58-e1fe-7a98-be5e225a3c0c%28Office.15%29.aspx)**|Returns  **true** if the current project was created in Access 2013 and onwards and **false** if the current project was created prior to Access 2013. Read-only **Boolean**.|

## Excel





|**Name**|**Description**|
|:-----|:-----|
|**[Chart.ShowExpandCollapseEntireFieldButtons Property (Excel)](http://msdn.microsoft.com/library/8fc5a821-ab24-2e48-1100-cec590786cd1%28Office.15%29.aspx)**|**True** to display the **Expand Entire Field** and **Collapse Entire Field** buttons on the specified pivot chart. Read/write **Boolean**.|
|**[ChartGroup.BinsCountValue Property (Excel)](http://msdn.microsoft.com/library/933ce137-4421-54c1-e3f7-f51466f6012d%28Office.15%29.aspx)**|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|**[ChartGroup.BinsOverflowEnabled Property (Excel)](http://msdn.microsoft.com/library/3af8d552-94e1-6f15-df2b-38fb7d3a0be1%28Office.15%29.aspx)**|Specifies whether a bin for values above the [BinsOverflowValue](http://msdn.microsoft.com/library/411856a7-ac17-e9eb-35bd-c851c0cfdfdc%28Office.15%29.aspx) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsOverflowValue Property (Excel)](http://msdn.microsoft.com/library/411856a7-ac17-e9eb-35bd-c851c0cfdfdc%28Office.15%29.aspx)**|If an [BinsOverflowEnabled](http://msdn.microsoft.com/library/3af8d552-94e1-6f15-df2b-38fb7d3a0be1%28Office.15%29.aspx) is **True**, specifies the value above which an overflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinsType Property (Excel)](http://msdn.microsoft.com/library/7230c44b-2e93-9790-2f27-d584688c6172%28Office.15%29.aspx)**|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](http://msdn.microsoft.com/library/99482ffa-a40c-c2b4-a062-ce5ce2ad5b29%28Office.15%29.aspx).|
|**[ChartGroup.BinsUnderflowEnabled Property (Excel)](http://msdn.microsoft.com/library/719d315a-c3ed-77e9-3b42-0f6300b6bf8d%28Office.15%29.aspx)**|Specifies whether a bin for values below the [BinsUnderflowValue](http://msdn.microsoft.com/library/39a9ec75-8283-e603-fddd-e165a1641203%28Office.15%29.aspx) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsUnderflowValue Property (Excel)](http://msdn.microsoft.com/library/39a9ec75-8283-e603-fddd-e165a1641203%28Office.15%29.aspx)**|If an [BinsUnderflowEnabled](http://msdn.microsoft.com/library/719d315a-c3ed-77e9-3b42-0f6300b6bf8d%28Office.15%29.aspx) is **True**, specifies the value below which an underflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinWidthValue Property (Excel)](http://msdn.microsoft.com/library/1b67cbda-6982-d16d-e039-6455cf7a845e%28Office.15%29.aspx)**|Specifies the number of points in each range. Read/write  **Double**.|
|**[CubeField.AutoGroup Method (Excel)](http://msdn.microsoft.com/library/72e1f6e7-edc5-6a9d-6632-a86064984e03%28Office.15%29.aspx)**|Automatically groups the cube fields in an OLAP cube, optionally in the specified orientation and/or at the specified position.|
|**[Model.ModelFormatBoolean Property (Excel)](http://msdn.microsoft.com/library/e38f7c66-2af8-8952-5d11-877d53b29d9e%28Office.15%29.aspx)**|Returns a [ModelFormatBoolean](http://msdn.microsoft.com/library/b6a43c30-1dd9-39e0-86dc-fd229bb51c87%28Office.15%29.aspx) object that represents formatting of type true/false in the data model. Read-only.|
|**[Model.ModelFormatCurrency Property (Excel)](http://msdn.microsoft.com/library/a8b6af70-624f-b018-be15-c8eeafa30d3a%28Office.15%29.aspx)**|Returns a [ModelFormatCurrency](http://msdn.microsoft.com/library/acb863b6-c188-5ed3-afe4-5e1ab6bb20bf%28Office.15%29.aspx) object that represents formatting of type currency in the data model. Read-only.|
|**[Model.ModelFormatDate Property (Excel)](http://msdn.microsoft.com/library/d1f5bd11-4b82-6dad-5e98-1a085d10fa47%28Office.15%29.aspx)**|Returns a [ModelFormatDate](http://msdn.microsoft.com/library/fe0be1f5-bd51-11cf-f0ba-f7c1ff228ecd%28Office.15%29.aspx) object that represents formatting of type date in the data model. Read-only.|
|**[Model.ModelFormatDecimalNumber Property (Excel)](http://msdn.microsoft.com/library/402b7073-0a6b-7856-5316-fc82ec980fc6%28Office.15%29.aspx)**|Returns a [ModelFormatDecimalNumber](http://msdn.microsoft.com/library/1080e484-4ec0-abdc-6322-5d83201c59fb%28Office.15%29.aspx) object that represents formatting of type decimal number in the data model. Read-only.|
|**[Model.ModelFormatGeneral Property (Excel)](http://msdn.microsoft.com/library/bac1d3bb-430e-8b0c-effb-81b2bc0ecf8c%28Office.15%29.aspx)**|Returns a [ModelFormatGeneral](http://msdn.microsoft.com/library/4fc68fb0-37aa-da83-f303-40ff96efb4a7%28Office.15%29.aspx) object that represents formatting of type general in the data model. Read-only.|
|**[Model.ModelFormatPercentageNumber Property (Excel)](http://msdn.microsoft.com/library/0efc53f5-bb5e-e367-8c23-0c65be87ea0c%28Office.15%29.aspx)**|Returns a [ModelFormatPercentageNumber](http://msdn.microsoft.com/library/1a7134a3-2645-e762-c2dd-1ca8ab8b6e73%28Office.15%29.aspx) object that represents formatting of type percentage number in the data model. Read-only.|
|**[Model.ModelFormatScientificNumber Property (Excel)](http://msdn.microsoft.com/library/6f7968b7-c765-0d7b-4485-852d54f3d471%28Office.15%29.aspx)**|Returns a [ModelFormatScientificNumber](http://msdn.microsoft.com/library/0099a473-0848-05ad-abe5-b36b70d4a2da%28Office.15%29.aspx) object that represents formatting of type scientific number in the data model. Read-only.|
|**[Model.ModelFormatWholeNumber Property (Excel)](http://msdn.microsoft.com/library/6a17f683-0617-f5eb-9cc9-040a68c8e452%28Office.15%29.aspx)**|Returns a [ModelFormatWholeNumber](http://msdn.microsoft.com/library/1a3d96ac-a2d7-cf26-5afa-6cfc8da846d5%28Office.15%29.aspx) object that represents formatting of type whole number in the data model. Read-only.|
|**[Model.ModelMeasures Property (Excel)](http://msdn.microsoft.com/library/b92f52fc-7c11-accc-bf3a-ba62c87daf71%28Office.15%29.aspx)**|Returns a [ModelMeasures](http://msdn.microsoft.com/library/b0edac9a-e10d-ec51-d9e7-6fa8a29dcda8%28Office.15%29.aspx) object that represents the collection of model measures in the data model. Read-only.|
|**[ModelConnection.CalculatedMembers Property (Excel)](http://msdn.microsoft.com/library/2969824d-b7a2-fb88-1066-cf5d36d8e9bb%28Office.15%29.aspx)**|Returns a [CalculatedMembers](http://msdn.microsoft.com/library/2969824d-b7a2-fb88-1066-cf5d36d8e9bb%28Office.15%29.aspx) object that represents the calculated members in the model connection.|
|**[ModelFormatBoolean Object (Excel)](http://msdn.microsoft.com/library/b6a43c30-1dd9-39e0-86dc-fd229bb51c87%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatBoolean.Application Property (Excel)](http://msdn.microsoft.com/library/a04b24b3-fb9c-0c59-06aa-aa5198e2017e%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatBoolean.Creator Property (Excel)](http://msdn.microsoft.com/library/b32a70e5-a6ae-e1ef-cc10-e86ca88f1578%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatBoolean.Parent Property (Excel)](http://msdn.microsoft.com/library/b581cf67-d77b-c17b-1878-1029d73682ff%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatCurrency Object (Excel)](http://msdn.microsoft.com/library/acb863b6-c188-5ed3-afe4-5e1ab6bb20bf%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatCurrency.Application Property (Excel)](http://msdn.microsoft.com/library/62fb4288-dc98-4831-a039-b2b81f407159%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatCurrency.Creator Property (Excel)](http://msdn.microsoft.com/library/069eb7ee-2168-0820-1018-61c1498c7929%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatCurrency.DecimalPlaces Property (Excel)](http://msdn.microsoft.com/library/99cd87a8-4aff-f507-05e3-59a28f676828%28Office.15%29.aspx)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatCurrency.Parent Property (Excel)](http://msdn.microsoft.com/library/ed21cd7a-ce59-2f39-c342-3292af63c079%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only|
|**[ModelFormatCurrency.Symbol Property (Excel)](http://msdn.microsoft.com/library/67c90858-43f0-5506-735b-747b7b9dcb07%28Office.15%29.aspx)**|Specifies the symbol to use to represent the currency. Read/write  **String**.|
|**[ModelFormatDate Object (Excel)](http://msdn.microsoft.com/library/fe0be1f5-bd51-11cf-f0ba-f7c1ff228ecd%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatDate.Application Property (Excel)](http://msdn.microsoft.com/library/a0932f28-60a5-34ce-a678-78cb72b33b31%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatDate.Creator Property (Excel)](http://msdn.microsoft.com/library/4f7b44a5-70da-be7d-306c-9a2d2c9ea724%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatDate.FormatString Property (Excel)](http://msdn.microsoft.com/library/2752f9be-4bb1-5bb6-7bee-eecaecafe0d9%28Office.15%29.aspx)**|Specifies the date format, for example, " _dd/mm/yy_ ". Read/write **String**.|
|**[ModelFormatDate.Parent Property (Excel)](http://msdn.microsoft.com/library/064b3651-ad76-c8fd-1250-b25fa669cf92%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatDecimalNumber Object (Excel)](http://msdn.microsoft.com/library/1080e484-4ec0-abdc-6322-5d83201c59fb%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatDecimalNumber.Application Property (Excel)](http://msdn.microsoft.com/library/71e9a26b-4e0b-0fdc-71b4-812fc5e96546%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatDecimalNumber.Creator Property (Excel)](http://msdn.microsoft.com/library/106db87c-9b52-1e74-e899-3da9de73bd3e%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatDecimalNumber.DecimalPlaces Property (Excel)](http://msdn.microsoft.com/library/74b0f4ba-a44d-8ca0-be24-11caecbc1fdd%28Office.15%29.aspx)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatDecimalNumber.Parent Property (Excel)](http://msdn.microsoft.com/library/f45fe29b-d869-e439-0aae-ab1bbe3b0793%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatDecimalNumber.UseThousandSeparator Property (Excel)](http://msdn.microsoft.com/library/97d01ea7-a6cf-9262-701b-3b0fac3ca571%28Office.15%29.aspx)**|Specifies whether to display commas between thousands. Read/write  **Boolean**.|
|**[ModelFormatGeneral Object (Excel)](http://msdn.microsoft.com/library/4fc68fb0-37aa-da83-f303-40ff96efb4a7%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatGeneral.Application Property (Excel)](http://msdn.microsoft.com/library/d05c69ad-e39e-e021-b827-04a73542e816%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatGeneral.Creator Property (Excel)](http://msdn.microsoft.com/library/828ced24-d35d-bee5-c9a6-b63e102c8cfb%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatGeneral.Parent Property (Excel)](http://msdn.microsoft.com/library/e6338b39-2822-d3f2-1067-e17deda4d71e%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatPercentageNumber Object (Excel)](http://msdn.microsoft.com/library/1a7134a3-2645-e762-c2dd-1ca8ab8b6e73%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatPercentageNumber.Application Property (Excel)](http://msdn.microsoft.com/library/bdcf764e-771f-9efe-d24f-ce03b047959c%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatPercentageNumber.Creator Property (Excel)](http://msdn.microsoft.com/library/1ff943c2-e52f-c01e-d337-d5dd7c02983e%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatPercentageNumber.DecimalPlaces Property (Excel)](http://msdn.microsoft.com/library/5828ab2d-1748-8ed9-8cad-10db422a6b8a%28Office.15%29.aspx)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatPercentageNumber.Parent Property (Excel)](http://msdn.microsoft.com/library/6f277a29-dc95-aff7-5b6f-1ffda156e3f1%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatPercentageNumber.UseThousandSeparator Property (Excel)](http://msdn.microsoft.com/library/f5f585ed-58db-f44a-525e-5c44c1a32168%28Office.15%29.aspx)**|Specifies whether to display commas between thousands. Read/write  **Boolean**.|
|**[ModelFormatScientificNumber Object (Excel)](http://msdn.microsoft.com/library/0099a473-0848-05ad-abe5-b36b70d4a2da%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatScientificNumber.Application Property (Excel)](http://msdn.microsoft.com/library/4caf7286-bcba-8628-15a4-d01d5e8cd575%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatScientificNumber.Creator Property (Excel)](http://msdn.microsoft.com/library/b764b8cb-b6f4-dca8-9bab-6add833dc61b%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatScientificNumber.DecimalPlaces Property (Excel)](http://msdn.microsoft.com/library/70cf2e5b-d7e1-d259-a7b8-188dfa0387d1%28Office.15%29.aspx)**|Specifies the number of decimal places after the dot. Read/write  **Long**.|
|**[ModelFormatScientificNumber.Parent Property (Excel)](http://msdn.microsoft.com/library/4eff9a0e-2a7d-5c76-3d0e-3d011908c118%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatWholeNumber Object (Excel)](http://msdn.microsoft.com/library/1a3d96ac-a2d7-cf26-5afa-6cfc8da846d5%28Office.15%29.aspx)**|Represents the format to be used for a model measure in the data model.|
|**[ModelFormatWholeNumber.Application Property (Excel)](http://msdn.microsoft.com/library/a2dada26-38f7-967a-be69-3f75f911c05e%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelFormatWholeNumber.Creator Property (Excel)](http://msdn.microsoft.com/library/82f16ccb-6f50-273e-5ed4-e16db1262ecc%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelFormatWholeNumber.Parent Property (Excel)](http://msdn.microsoft.com/library/0455db3e-b4db-635b-7e72-ee6fcc366512%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelFormatWholeNumber.UseThousandSeparator Property (Excel)](http://msdn.microsoft.com/library/7378fadd-cd13-c0f7-525a-7e30eb59d4bb%28Office.15%29.aspx)**|Specifies whether to display commas between thousands. Read/write  **Boolean**.|
|**[ModelMeasure Object (Excel)](http://msdn.microsoft.com/library/0df4620a-e250-a68e-7000-6959cde08f3e%28Office.15%29.aspx)**|Represents a single  **ModelMeasure** object in the[ModelMeasures](http://msdn.microsoft.com/library/b0edac9a-e10d-ec51-d9e7-6fa8a29dcda8%28Office.15%29.aspx) collection.|
|**[ModelMeasure.Application Property (Excel)](http://msdn.microsoft.com/library/fa2112fb-ce9d-48bc-63cf-34faa2cf3488%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelMeasure.AssociatedTable Property (Excel)](http://msdn.microsoft.com/library/a51f7698-38a4-211e-3973-de9c8b62e66d%28Office.15%29.aspx)**|Specifies the table that contains the model measure, as displayed in the  **Field List** task pane. Read/write[ModelTable](http://msdn.microsoft.com/library/c853beb6-f2e7-dda0-b33a-8110a6c23de8%28Office.15%29.aspx).|
|**[ModelMeasure.Creator Property (Excel)](http://msdn.microsoft.com/library/abdbcd11-2b57-043f-20ba-e8b4ed603130%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelMeasure.Delete Method (Excel)](http://msdn.microsoft.com/library/19ec1efa-5c1e-8130-3845-7e3f55017041%28Office.15%29.aspx)**|Deletes the model measure from the data model.|
|**[ModelMeasure.Description Property (Excel)](http://msdn.microsoft.com/library/f80228a3-ea61-4d00-6a93-609914c3a21e%28Office.15%29.aspx)**|The description of the model measure. Read/write  **String**.|
|**[ModelMeasure.FormatInformation Property (Excel)](http://msdn.microsoft.com/library/26ad6641-c4fe-ae9d-b8dd-d429f5806790%28Office.15%29.aspx)**|The format of the model measure. Read/write  **Variant**.|
|**[ModelMeasure.Formula Property (Excel)](http://msdn.microsoft.com/library/fc6f6b80-1a80-3400-e98f-4cb74ad8e6e8%28Office.15%29.aspx)**|The Data Analysis Expressions (DAX) formula of the model measure. Read/write  **String**.|
|**[ModelMeasure.Name Property (Excel)](http://msdn.microsoft.com/library/11e33619-17c3-3d78-29ad-048949fd06b3%28Office.15%29.aspx)**|The name of the model measure. Read/write  **String**.|
|**[ModelMeasure.Parent Property (Excel)](http://msdn.microsoft.com/library/900b9ddd-5567-ca95-5122-02fc670abf0f%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelMeasures Object (Excel)](http://msdn.microsoft.com/library/b0edac9a-e10d-ec51-d9e7-6fa8a29dcda8%28Office.15%29.aspx)**|Represents: a collection of  **ModelMeasure** objects.|
|**[ModelMeasures.Add Method (Excel)](http://msdn.microsoft.com/library/abc0f260-abdb-2f60-928f-b325fbb976f3%28Office.15%29.aspx)**|Adds a model measure to the model.|
|**[ModelMeasures.Application Property (Excel)](http://msdn.microsoft.com/library/bf2c2284-b45b-5a68-b02a-c2cc88babcd4%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[ModelMeasures.Count Property (Excel)](http://msdn.microsoft.com/library/c420f7e8-ecc1-988b-5438-280f3fb3b5e1%28Office.15%29.aspx)**|Returns an integer that represents the number of objects in the collection.|
|**[ModelMeasures.Creator Property (Excel)](http://msdn.microsoft.com/library/575d569a-5932-8e3e-66fa-61e7e67e3afa%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[ModelMeasures.Item Method (Excel)](http://msdn.microsoft.com/library/cbadae47-2225-4633-eac6-8697227384f4%28Office.15%29.aspx)**|Returns a single object from a collection|
|**[ModelMeasures.Parent Property (Excel)](http://msdn.microsoft.com/library/61d981d0-bc20-efea-1fdd-49c6e188670c%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[ModelRelationships.DetectRelationships Method (Excel)](http://msdn.microsoft.com/library/e6db4a4b-09c4-7564-f3c7-3aed719dcc16%28Office.15%29.aspx)**|Detects model relationships in the specified [PivotTable](http://msdn.microsoft.com/library/a9c1d4a0-78a9-f9a6-6daf-91cb63e45842%28Office.15%29.aspx).|
|**[PivotField.AutoGroup Method (Excel)](http://msdn.microsoft.com/library/b8806ccf-a4c0-3dfc-a04b-3244ccfb3163%28Office.15%29.aspx)**|Automatically groups the pivot fields in a pivot table.|
|**[Point.IsTotal Property (Excel)](http://msdn.microsoft.com/library/65269b0f-cb65-eb9c-b2d3-0b73d7677801%28Office.15%29.aspx)**|**True** if the point represents a total. Read/write **Boolean**.|
|**[Queries Object (Excel)](http://msdn.microsoft.com/library/3c16b2f6-8189-352a-4c4e-513bdb9c01d5%28Office.15%29.aspx)**|The collection of [WorkbookQuery](http://msdn.microsoft.com/library/2a27186f-5e02-f026-bee2-b4c7aa852711%28Office.15%29.aspx) objects|
|**[Queries.Add Method (Excel)](http://msdn.microsoft.com/library/184711c0-2ce4-ba6e-df56-1f7fdd60ab2c%28Office.15%29.aspx)**|Adds a new [WorkbookQuery](http://msdn.microsoft.com/library/2a27186f-5e02-f026-bee2-b4c7aa852711%28Office.15%29.aspx) object to the **Queries** collection.|
|**[Queries.Application Property (Excel)](http://msdn.microsoft.com/library/83778da5-1c09-1465-f651-88eb00179da3%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[Queries.Count Property (Excel)](http://msdn.microsoft.com/library/b9553330-01ff-8c31-ba10-62176f1ba0b7%28Office.15%29.aspx)**|Returns an integer that represents the number of objects in the collection.|
|**[Queries.Creator Property (Excel)](http://msdn.microsoft.com/library/1e20a980-6f8d-e780-dd0e-3f0b428d97ea%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[Queries.FastCombine Property (Excel)](http://msdn.microsoft.com/library/6d34ab2f-5dd4-6dd9-74c0-b49c600db45b%28Office.15%29.aspx)**|**True** to enable the fast combine feature, as long as the workbook is open. Read/write **Boolean**.|
|**[Queries.Item Method (Excel)](http://msdn.microsoft.com/library/d87f5019-dde2-972a-67f8-de7bf5d07b66%28Office.15%29.aspx)**|Returns a single object from a collection.|
|**[Queries.Parent Property (Excel)](http://msdn.microsoft.com/library/01f66159-a7bd-bffd-29a7-ff13c20fadb0%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[Series.ParentDataLabelOption Property (Excel)](http://msdn.microsoft.com/library/ba86d954-7442-5023-e663-eea3626588e6%28Office.15%29.aspx)**|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XLParentDataLabelOptions](http://msdn.microsoft.com/library/eb2c2212-e538-e6a4-2a76-c14808ff679c%28Office.15%29.aspx).|
|**[Series.QuartileCalculationInclusiveMedian Property (Excel)](http://msdn.microsoft.com/library/eda57981-1903-d1a8-1c53-80272191e077%28Office.15%29.aspx)**|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|**[SoundNote Object (Excel)](http://msdn.microsoft.com/library/518e9707-4696-e7ad-7547-b746131e311b%28Office.15%29.aspx)**|Represents a recorded sound note.|
|**[SoundNote.Application Property (Excel)](http://msdn.microsoft.com/library/3adf2c05-3fc5-6a29-8c4f-ea6021db2802%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[SoundNote.Creator Property (Excel)](http://msdn.microsoft.com/library/3b59f17a-56ca-16b0-4094-ead8e42ffd89%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[SoundNote.Parent Property (Excel)](http://msdn.microsoft.com/library/373f7b42-ca1d-1eb9-e499-18120c5353d3%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[SoundNote.Delete Method (Excel)](http://msdn.microsoft.com/library/a5700e45-9ee6-8aba-205d-fe7927b367d2%28Office.15%29.aspx)**|Deletes the sound note.|
|**[SoundNote.Import Method (Excel)](http://msdn.microsoft.com/library/c5fe13cb-aa95-c150-3290-b6a6e45616af%28Office.15%29.aspx)**|Imports the specified sound note.|
|**[SoundNote.Play Method (Excel)](http://msdn.microsoft.com/library/c7a78257-75d3-d131-2d46-d01bf4598de5%28Office.15%29.aspx)**|Plays the sound note.|
|**[SoundNote.Record Method (Excel)](http://msdn.microsoft.com/library/cc17091c-38e7-508f-80e3-3ac7e320c9ed%28Office.15%29.aspx)**|Records the sound note.|
|**[Workbook.CreateForecastSheet Method (Excel)](http://msdn.microsoft.com/library/bec7b60b-7840-af15-6d5f-f5c184ea7aee%28Office.15%29.aspx)**|If you have historical time-based data, you can use  **CreateForecastSheet** to create a forecast. When you create a forecast, a new worksheet is created that contains a table of the historical and predicted values and a chart showing this. A forecast can help you predict things like future sales, inventory requirements, or consumer trends.|
|**[WorkbookQuery Object (Excel)](http://msdn.microsoft.com/library/2a27186f-5e02-f026-bee2-b4c7aa852711%28Office.15%29.aspx)**|An object that represents a query that was created by Power Query.|
|**[WorkbookQuery.Application Property (Excel)](http://msdn.microsoft.com/library/b025538e-ac17-60c9-337e-0b6ce4a7943f%28Office.15%29.aspx)**|When used without an object qualifier, this property returns an [Application](http://msdn.microsoft.com/library/19b73597-5cf9-4f56-8227-b5211f657f6f%28Office.15%29.aspx) object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|**[WorkbookQuery.Creator Property (Excel)](http://msdn.microsoft.com/library/82e257ca-9e3f-0acc-66a7-84f7e7e07ff8%28Office.15%29.aspx)**|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
|**[WorkbookQuery.Delete Method (Excel)](http://msdn.microsoft.com/library/05f42f34-1814-870f-081a-c0538b438aec%28Office.15%29.aspx)**|Deletes this query and its underlying connection and removes it from the  **Queries** collection.|
|**[WorkbookQuery.Description Property (Excel)](http://msdn.microsoft.com/library/1175e1df-0788-99aa-2bb3-9dfa545125f3%28Office.15%29.aspx)**|The description of the query. Read/write  **String**.|
|**[WorkbookQuery.Formula Property (Excel)](http://msdn.microsoft.com/library/62c5fcfa-8359-5fab-1a5d-fdbb2793bf53%28Office.15%29.aspx)**|The Power Query M formula for the object. Read-only  **String**.|
|**[WorkbookQuery.Name Property (Excel)](http://msdn.microsoft.com/library/afc6c679-8dda-08f9-c896-775b395b5e92%28Office.15%29.aspx)**|The name of the query. Read/write  **String**.|
|**[WorkbookQuery.Parent Property (Excel)](http://msdn.microsoft.com/library/246acb77-2a0b-b988-48ba-a18f0d6e0361%28Office.15%29.aspx)**|Returns the parent object for the specified object. Read-only.|
|**[WorksheetFunction.Forecast_ETS Method (Excel)](http://msdn.microsoft.com/library/de915259-3d2a-485a-8027-290dc9cb95a5%28Office.15%29.aspx)**|Calculates or predicts a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm. |
|**[WorksheetFunction.Forecast_ETS_ConfInt Method (Excel)](http://msdn.microsoft.com/library/23d6cb35-58c8-6ef0-ed4f-5c693974ccd2%28Office.15%29.aspx)**|Returns a confidence interval for the forecast value at the specified target date.|
|**[WorksheetFunction.Forecast_ETS_Seasonality Method (Excel)](http://msdn.microsoft.com/library/aad7c233-1745-64e3-22a9-ade62e5e177d%28Office.15%29.aspx)**|Returns the length of the repetitive pattern Excel detects for the specified time series.|
|**[WorksheetFunction.Forecast_ETS_STAT Method (Excel)](http://msdn.microsoft.com/library/6b1c0256-3146-4dc5-3f8a-27e61a982fee%28Office.15%29.aspx)**|Returns a statistical value as a result of time series forecasting.|
|**[WorksheetFunction.Forecast_Linear Method (Excel)](http://msdn.microsoft.com/library/71b85d12-0c81-f82d-99fe-ad712f2530e5%28Office.15%29.aspx)**|Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. You can use this function to predict future sales, inventory requirements, or consumer trends.|
|**[XlBinsType Enumeration (Excel)](http://msdn.microsoft.com/library/99482ffa-a40c-c2b4-a062-ce5ce2ad5b29%28Office.15%29.aspx)**|Constants passed to and returned by the [ChartGroup.BinsType](http://msdn.microsoft.com/library/7230c44b-2e93-9790-2f27-d584688c6172%28Office.15%29.aspx) property.|
|**[XlForecastAggregation Enumeration (Excel)](http://msdn.microsoft.com/library/00df6eeb-05ab-e004-7cee-56f520096f72%28Office.15%29.aspx)**|Constants passed to various  **WorksheetFunction** and **Workbook** statistical forecasting methods.|
|**[XlForecastChartType Enumeration (Excel)](http://msdn.microsoft.com/library/7296fb27-dccf-6ad4-3565-453e9fae1b77%28Office.15%29.aspx)**|Constants passed to the [Workbook.CreateForecastSheet](http://msdn.microsoft.com/library/bec7b60b-7840-af15-6d5f-f5c184ea7aee%28Office.15%29.aspx) Method.|
|**[XlForecastDataCompletion Enumeration (Excel)](http://msdn.microsoft.com/library/0407a50c-2efe-1522-7666-b5a8b4e72a83%28Office.15%29.aspx)**|Constants passed to various  **WorksheetFunction** and **Workbook** statistical forecasting methods.|
|**[XlParentDataLabelOptions Enumeration (Excel)](http://msdn.microsoft.com/library/eb2c2212-e538-e6a4-2a76-c14808ff679c%28Office.15%29.aspx)**|Constants passed to and returned by the  **Series.ParentDataLabelOption** property.|

## Outlook





|**Name**|**Description**|
|:-----|:-----|
|**[ExchangeDistributionList.GetUnifiedGroup Method (Outlook)](http://msdn.microsoft.com/library/9b129256-02c0-438a-9098-c0925ec60388%28Office.15%29.aspx)**|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](http://msdn.microsoft.com/library/9ee27465-3ea5-7316-feec-2f255ff08f6b%28Office.15%29.aspx)|
|**[ExchangeDistributionList.GetUnifiedGroupFromStore Method (Outlook)](http://msdn.microsoft.com/library/8565a110-d9d9-bc58-a5c8-a0ac9da8ec0c%28Office.15%29.aspx)**|Determines if the object is a unified group (by way of a call to [IsUnifiedGroup](http://msdn.microsoft.com/library/9ee27465-3ea5-7316-feec-2f255ff08f6b%28Office.15%29.aspx)) and returns the  **Outlook.Folder** object associated with the group using the[GetUnifiedGroup](http://msdn.microsoft.com/library/9b129256-02c0-438a-9098-c0925ec60388%28Office.15%29.aspx) and **GetUnifiedGroupFromStore** methods.|
|**[ExchangeDistributionList.IsUnifiedGroup Method (Outlook)](http://msdn.microsoft.com/library/9ee27465-3ea5-7316-feec-2f255ff08f6b%28Office.15%29.aspx)**|Determines if the object is a unified group.|
|**[ExchangeUser.GetUnifiedGroup Method (Outlook)](http://msdn.microsoft.com/library/ec0f58fa-969d-ed38-705b-2c99ccbf3c86%28Office.15%29.aspx)**|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](http://msdn.microsoft.com/library/46f9564a-1c0a-fe6c-3f06-989fb5f36adf%28Office.15%29.aspx).|
|**[ExchangeUser.GetUnifiedGroupFromStore Method (Outlook)](http://msdn.microsoft.com/library/38a901d3-670f-afd2-a385-3b2bb859cb81%28Office.15%29.aspx)**|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](http://msdn.microsoft.com/library/46f9564a-1c0a-fe6c-3f06-989fb5f36adf%28Office.15%29.aspx).|
|**[ExchangeUser.IsUnifiedGroup Method (Outlook)](http://msdn.microsoft.com/library/46f9564a-1c0a-fe6c-3f06-989fb5f36adf%28Office.15%29.aspx)**|Determines if the object is a unified group.|
|**[Explorer.DisplayMode Property (Outlook)](http://msdn.microsoft.com/library/8e6bcc0d-5a37-2c8f-d059-28706b638dee%28Office.15%29.aspx)**|Indicates the display mode: Normal, Portrait View, or Portrait Reading Pane.|
|**[Explorer.DisplayModeChange Event (Outlook)](http://msdn.microsoft.com/library/cee77aad-8905-efed-466e-c2e88cfeeaa2%28Office.15%29.aspx)**|Occurs when the user performs an action that changes the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|
|**[Explorer.PreviewPane Property (Outlook)](http://msdn.microsoft.com/library/5f3edb49-b9f6-db03-8f83-3fe27f0aaf08%28Office.15%29.aspx)**|The [PreviewPane](http://msdn.microsoft.com/library/fd4f497b-7085-6e0f-018b-17845f4dfe61%28Office.15%29.aspx) object displays content in a "single pane mode" by showing only the Preview Pane view.|
|**[ExplorerEvents_10.DisplayModeChange Method (Outlook)](http://msdn.microsoft.com/library/8805ec85-d6b2-dec4-2179-9de0b08a2a7b%28Office.15%29.aspx)**|Occurs when the user performs an action that changes the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|
|**[OlDisplayMode Enumeration (Outlook)](http://msdn.microsoft.com/library/a5312dea-ccde-d417-6f40-013e63c107f8%28Office.15%29.aspx)**|Describes the nature of the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|
|**[OlUnifiedGroupFolderType Enumeration (Outlook)](http://msdn.microsoft.com/library/7ee0ae00-17e4-320b-8e52-f759193f6232%28Office.15%29.aspx)**|Specifies the folder to be obtained for unified groups. Because groups have both a mail and a calendar folder, you can specify either the  **olGroupMailFolder** or **olGroupCalendarFolder**.|
|**[OlUnifiedGroupType Enumeration (Outlook)](http://msdn.microsoft.com/library/e750a22a-4e76-9458-fccd-7f2babcf9485%28Office.15%29.aspx)**|Specifies the group type as public or private for the [CreateUnifiedGroup](http://msdn.microsoft.com/library/45f70f08-f198-22a2-79c5-26dc3247e164%28Office.15%29.aspx) method.|
|**[PreviewPane Members (Outlook)](http://msdn.microsoft.com/library/42ded67c-b3cb-a479-a110-fd3db9548d3b%28Office.15%29.aspx)**|Displays content in a "single pane mode" by showing only the Preview Pane view.|
|**[PreviewPane Object (Outlook)](http://msdn.microsoft.com/library/fd4f497b-7085-6e0f-018b-17845f4dfe61%28Office.15%29.aspx)**|Displays content in a "single pane mode" by showing only the Preview Pane view.|
|**[PreviewPane.Application Property (Outlook)](http://msdn.microsoft.com/library/19fc11bc-a777-349e-f708-3cfa9f24ecbd%28Office.15%29.aspx)**|Returns the [Application](http://msdn.microsoft.com/library/797003e7-ecd1-eccb-eaaf-32d6ddde8348%28Office.15%29.aspx) object that represents the parent application (Outlook) for the[PreviewPane](http://msdn.microsoft.com/library/fd4f497b-7085-6e0f-018b-17845f4dfe61%28Office.15%29.aspx) Object. Read-only.|
|**[PreviewPane.Class Property (Outlook)](http://msdn.microsoft.com/library/e6dd78bb-01d6-a351-156c-cb278435c922%28Office.15%29.aspx)**|Returns a constant in the [OlObjectClass](http://msdn.microsoft.com/library/33d724b3-df3c-2a7f-a80f-93b66d96f588%28Office.15%29.aspx) enumeration indicating the class of the[PreviewPane](http://msdn.microsoft.com/library/fd4f497b-7085-6e0f-018b-17845f4dfe61%28Office.15%29.aspx) Object. Read-only.|
|**[PreviewPane.Parent Property (Outlook)](http://msdn.microsoft.com/library/ab92d2d9-ebc6-d9f0-ca37-04a61ee33f3f%28Office.15%29.aspx)**|Returns the parent property for the [PreviewPane](http://msdn.microsoft.com/library/fd4f497b-7085-6e0f-018b-17845f4dfe61%28Office.15%29.aspx) Object. Read only.|
|**[PreviewPane.Session Property (Outlook)](http://msdn.microsoft.com/library/54509e05-d255-b96e-f037-14282791ea55%28Office.15%29.aspx)**|Returns the [NameSpace](http://msdn.microsoft.com/library/f0dcaa19-07f5-5d42-a3bf-2e42b7885644%28Office.15%29.aspx) for the current session. Read-only.|
|**[PreviewPane.WordEditor Property (Outlook)](http://msdn.microsoft.com/library/8c50e511-99ed-a691-352e-ae8f0942dbe5%28Office.15%29.aspx)**|Returns the Microsoft Word Document Object Model of the message being displayed. Read-only.|
|**[Store.CreateUnifiedGroup Method (Outlook)](http://msdn.microsoft.com/library/45f70f08-f198-22a2-79c5-26dc3247e164%28Office.15%29.aspx)**|Enables a unified group to be created.|
|**[Store.DeleteUnifiedGroup Method (Outlook)](http://msdn.microsoft.com/library/53c15736-f88a-33ad-2b21-29a2c9c6d402%28Office.15%29.aspx)**|Enables a unified group to be deleted.|

## Project





|**Name**|**Description**|
|:-----|:-----|
|**[Application.AddEngagement Method (Project)](http://msdn.microsoft.com/library/61fbd902-1fa1-d591-5618-697e5dc9338d%28Office.15%29.aspx)**|Adds a  **Resource Plan** view, enabling users to display and edit engagement data to Project when connected to Project Online. Introduced in Office 2016.|
|**[Application.EngagementInfo Method (Project)](http://msdn.microsoft.com/library/4e95d901-77a0-f1f7-b754-aefeb720e5ea%28Office.15%29.aspx)**|Displays the engagement information dialog box user interface for the  **Resource Plan** view. Introduced in Office 2016.|
|**[Application.GetDpiScaleFactor Method (Project)](http://msdn.microsoft.com/library/d1e7f1e5-095c-aa4c-0550-1a077c1a2de3%28Office.15%29.aspx)**|Indicates the  **DPI Scale Factor**, used for optimizing scale settings. Introduced in Office 2016.|
|**[Application.InsertTimelineBar Method (Project)](http://msdn.microsoft.com/library/2cb9d639-3363-79e3-ced6-73b0a574986a%28Office.15%29.aspx)**|Adds a  **timeline** bar to the view.|
|**[Application.Inspector Method (Project)](http://msdn.microsoft.com/library/f386160f-232a-7e4d-37e0-9c090a58df8a%28Office.15%29.aspx)**|Indicates the  **Task Inspector** for use with engagement data.|
|**[Application.LocaleName Method (Project)](http://msdn.microsoft.com/library/989d8c73-3452-2abe-fbaa-f68d532e353e%28Office.15%29.aspx)**|Language name that is used by Project, such as en-us or za-ch.|
|**[Application.ProjectSummaryInfoEx Method (Project)](http://msdn.microsoft.com/library/2827f735-6a7b-9f33-c1c6-2c5f1f7492f6%28Office.15%29.aspx)**|Returns information about project summary, including the Project Utilization type and Project Utilization date information.|
|**[Application.RefreshEngagementsForProject Method (Project)](http://msdn.microsoft.com/library/f0530b2b-18de-70b8-d27d-a51ded376fe3%28Office.15%29.aspx)**|Refreshes the engagements for the project using engagement state on the server.|
|**[Application.RemoveTimelineBar Method (Project)](http://msdn.microsoft.com/library/8385d889-b81e-5422-a032-c7073fa7c65d%28Office.15%29.aspx)**|Removes a  **Timeline** bar from the view.|
|**[Application.SubmitAllEngagementsForProject Method (Project)](http://msdn.microsoft.com/library/7e695f9f-5c0b-bbbf-9abe-a695e72591a1%28Office.15%29.aspx)**|Submits all the engagements in the project to the resource manager for review.|
|**[Application.SubmitSelectedEngagementsForProject Method (Project)](http://msdn.microsoft.com/library/bfa4d8b5-5806-54d9-009e-ff8fcb96d994%28Office.15%29.aspx)**|Submits all selected engagements in the project to the resource manager for review.|
|**[Application.TaskOnTimelineEx Method (Project)](http://msdn.microsoft.com/library/4307f842-0ccc-d7ac-f386-ec8d259011c6%28Office.15%29.aspx)**|Manages tasks on the Timeline pane or for a specified custom timeline, including specifying the bar that you want to add or remove.|
|**[Application.TimelineBarDateRange Method (Project)](http://msdn.microsoft.com/library/a1d257f3-92b7-6719-4ce5-5b959823e702%28Office.15%29.aspx)**|Modifies the start and finish dates for a  **Timeline** bar.|
|**[Application.UpdateEngagementsForProject Method (Project)](http://msdn.microsoft.com/library/cda633ec-2143-0f6e-80eb-2d9751d8782f%28Office.15%29.aspx)**|Update the Engagements for a Project.|
|**[Assignment.Compliant Property (Project)](http://msdn.microsoft.com/library/bceddf30-8cb4-4098-c354-46c044a97b0a%28Office.15%29.aspx)**|Gets the compliant for a task assignment in Project. Read-only.|
|**[Cell.Engagement Property (Project)](http://msdn.microsoft.com/library/14cbaf04-f1dc-dfe8-e40f-5d92446ee491%28Office.15%29.aspx)**|Gets or sets the engagement resource for a cell.|
|**[Chart Members (Project)](http://msdn.microsoft.com/library/7632e4dd-46f3-e463-5623-e8f2b2096390%28Office.15%29.aspx)**|The  **Chart** object represents a chart on a report in Project.|
|**[Engagement Object (Project)](http://msdn.microsoft.com/library/3e7f7bed-e575-a5f4-25e5-1c1cbe1880bb%28Office.15%29.aspx)**||
|**[Engagement.Application Property (Project)](http://msdn.microsoft.com/library/c5f3a831-22e9-a747-30c7-997ac97ff3e9%28Office.15%29.aspx)**||
|**[Engagement.Comments Property (Project)](http://msdn.microsoft.com/library/3d0a850a-6edf-e359-4c2d-dbba1c7c5d6e%28Office.15%29.aspx)**||
|**[Engagement.CommittedFinish Property (Project)](http://msdn.microsoft.com/library/9f2e166d-a609-1816-3c52-3127e3f21dd0%28Office.15%29.aspx)**||
|**[Engagement.CommittedMaxUnits Property (Project)](http://msdn.microsoft.com/library/84765743-234a-e293-9d3a-e6dd1a51790b%28Office.15%29.aspx)**||
|**[Engagement.CommittedStart Property (Project)](http://msdn.microsoft.com/library/793a9ba6-5250-54af-4f26-349abf59d5fc%28Office.15%29.aspx)**||
|**[Engagement.CommittedWork Property (Project)](http://msdn.microsoft.com/library/cd30cfc3-b1fa-19e2-49a1-f77eab1981d6%28Office.15%29.aspx)**||
|**[Engagement.CreatedDate Property (Project)](http://msdn.microsoft.com/library/22ad79fa-2d98-4f79-d5ed-91ac93c2b5c9%28Office.15%29.aspx)**||
|**[Engagement.Delete Method (Project)](http://msdn.microsoft.com/library/87c34ec9-157f-5f76-150d-036161f35363%28Office.15%29.aspx)**||
|**[Engagement.DraftFinish Property (Project)](http://msdn.microsoft.com/library/ae298776-46f2-c39a-5fa4-9b56499526d5%28Office.15%29.aspx)**||
|**[Engagement.DraftMaxUnits Property (Project)](http://msdn.microsoft.com/library/fa77a2ac-445f-ccbd-88fc-b8bd66e31783%28Office.15%29.aspx)**||
|**[Engagement.DraftStart Property (Project)](http://msdn.microsoft.com/library/352ffdd1-364b-ec22-286f-babf39bf6bb5%28Office.15%29.aspx)**||
|**[Engagement.DraftWork Property (Project)](http://msdn.microsoft.com/library/dfcc1702-1004-bf5b-c70f-88e30c331304%28Office.15%29.aspx)**||
|**[Engagement.GetField Method (Project)](http://msdn.microsoft.com/library/2c16e270-d7ad-e085-437f-a401cd10f26e%28Office.15%29.aspx)**||
|**[Engagement.Guid Property (Project)](http://msdn.microsoft.com/library/bd65661c-982d-8a1d-8d1b-24a41c9c5abd%28Office.15%29.aspx)**||
|**[Engagement.Index Property (Project)](http://msdn.microsoft.com/library/5d55800f-ea9f-de13-e81e-d6450e0297cb%28Office.15%29.aspx)**||
|**[Engagement.ModifiedByGuid Property (Project)](http://msdn.microsoft.com/library/390a65a7-21c1-bd3d-da88-a62108287465%28Office.15%29.aspx)**||
|**[Engagement.ModifiedByName Property (Project)](http://msdn.microsoft.com/library/97a04489-324b-454b-66a8-3a5915bfd5cb%28Office.15%29.aspx)**||
|**[Engagement.ModifiedDate Property (Project)](http://msdn.microsoft.com/library/a15d070c-f714-267a-b0bc-8ae9920bc68c%28Office.15%29.aspx)**||
|**[Engagement.Name Property (Project)](http://msdn.microsoft.com/library/f889308f-e395-67da-5691-c7a53a1856f3%28Office.15%29.aspx)**||
|**[Engagement.Parent Property (Project)](http://msdn.microsoft.com/library/33522e59-e840-b3af-79f3-3f92035853d9%28Office.15%29.aspx)**||
|**[Engagement.ProjectGuid Property (Project)](http://msdn.microsoft.com/library/93dfc0f4-06ad-7c4b-de6b-e224a5151689%28Office.15%29.aspx)**||
|**[Engagement.ProjectName Property (Project)](http://msdn.microsoft.com/library/b1a82d6e-850d-e519-1d17-1699b1ecb56f%28Office.15%29.aspx)**||
|**[Engagement.ProposedFinish Property (Project)](http://msdn.microsoft.com/library/2c2233f2-ee0b-5054-1300-ed4afdfd4c5f%28Office.15%29.aspx)**||
|**[Engagement.ProposedMaxUnits Property (Project)](http://msdn.microsoft.com/library/e0cee0d4-b9b8-9368-18dc-d39733996ec8%28Office.15%29.aspx)**||
|**[Engagement.ProposedStart Property (Project)](http://msdn.microsoft.com/library/ba467fd7-2930-a8b1-6477-0b28a731b9af%28Office.15%29.aspx)**||
|**[Engagement.ProposedWork Property (Project)](http://msdn.microsoft.com/library/85a93a89-8516-b72b-7aff-3b738a419547%28Office.15%29.aspx)**||
|**[Engagement.ResourceGuid Property (Project)](http://msdn.microsoft.com/library/9b92c2a6-891d-c7d0-97a8-aee2deee7277%28Office.15%29.aspx)**||
|**[Engagement.ResourceID Property (Project)](http://msdn.microsoft.com/library/11a1cb67-e799-5dbb-8361-8668a991eaee%28Office.15%29.aspx)**||
|**[Engagement.ResourceName Property (Project)](http://msdn.microsoft.com/library/0fd48448-b63c-207c-6aa3-eae693ea47e8%28Office.15%29.aspx)**||
|**[Engagement.ReviewedByGuid Property (Project)](http://msdn.microsoft.com/library/f7080dbb-93f2-fe06-38c2-ed56fd3d73c0%28Office.15%29.aspx)**||
|**[Engagement.ReviewedByName Property (Project)](http://msdn.microsoft.com/library/264c2472-cf6d-7fb5-956d-857c40a016b9%28Office.15%29.aspx)**||
|**[Engagement.ReviewedDate Property (Project)](http://msdn.microsoft.com/library/a7cddc80-6ebe-7fd7-553c-ad7f478b8cab%28Office.15%29.aspx)**||
|**[Engagement.SetField Method (Project)](http://msdn.microsoft.com/library/2f5f578f-a172-512c-1309-6910018281f0%28Office.15%29.aspx)**||
|**[Engagement.Status Property (Project)](http://msdn.microsoft.com/library/d928fab4-e451-2384-8b0d-1493b444b390%28Office.15%29.aspx)**||
|**[Engagement.SubmittedByGuid Property (Project)](http://msdn.microsoft.com/library/48885af4-e230-b4df-ae40-b1a285080e89%28Office.15%29.aspx)**||
|**[Engagement.SubmittedByName Property (Project)](http://msdn.microsoft.com/library/1b310aec-2e0d-1386-c3ba-875356abd704%28Office.15%29.aspx)**||
|**[Engagement.SubmittedDate Property (Project)](http://msdn.microsoft.com/library/b241f0da-0a2c-3faf-4a3b-44bfa327e3b5%28Office.15%29.aspx)**||
|**[EngagementComment Members (Project)](http://msdn.microsoft.com/library/739c0d51-7f6a-90d6-5160-c8634c6dffe3%28Office.15%29.aspx)**||
|**[EngagementComment Object (Project)](http://msdn.microsoft.com/library/4ca86b23-f8a2-0939-3cc5-196e72d06f01%28Office.15%29.aspx)**||
|**[EngagementComment Properties (Project)](http://msdn.microsoft.com/library/32d609b0-8df1-0eaa-d9f9-7735bd5e5289%28Office.15%29.aspx)**||
|**[EngagementComment.Application Property (Project)](http://msdn.microsoft.com/library/7c74fa87-932a-6f46-72cd-3f0ad3dfa245%28Office.15%29.aspx)**||
|**[EngagementComment.AuthorResEmail Property (Project)](http://msdn.microsoft.com/library/a14c0d6c-2163-b7ce-86a8-b44ab691a386%28Office.15%29.aspx)**||
|**[EngagementComment.AuthorResGuid Property (Project)](http://msdn.microsoft.com/library/551e1ae2-346a-aac5-7fca-ac92f6983cc6%28Office.15%29.aspx)**||
|**[EngagementComment.AuthorResName Property (Project)](http://msdn.microsoft.com/library/1c148709-ce9b-ff90-3f4c-932e2c6f79aa%28Office.15%29.aspx)**||
|**[EngagementComment.CreatedDate Property (Project)](http://msdn.microsoft.com/library/1d963d79-e219-9c91-2332-6c977dd346fa%28Office.15%29.aspx)**||
|**[EngagementComment.Guid Property (Project)](http://msdn.microsoft.com/library/d36b982b-bf3a-cdfe-d910-f1cd2bdab769%28Office.15%29.aspx)**||
|**[EngagementComment.Message Property (Project)](http://msdn.microsoft.com/library/b54430ec-7d99-76eb-2895-7c54abea6bc2%28Office.15%29.aspx)**||
|**[EngagementComment.Parent Property (Project)](http://msdn.microsoft.com/library/d27685a9-4a21-9095-d6e0-8a3978faf11d%28Office.15%29.aspx)**||
|**[EngagementComments Members (Project)](http://msdn.microsoft.com/library/5d231a50-8c3a-c299-b5c9-81da32fedccc%28Office.15%29.aspx)**||
|**[EngagementComments Methods (Project)](http://msdn.microsoft.com/library/8c5db094-4141-bdcc-392f-ded97b647004%28Office.15%29.aspx)**||
|**[EngagementComments Object (Project)](http://msdn.microsoft.com/library/6df493a2-5580-f6bc-373e-565ce1be6828%28Office.15%29.aspx)**||
|**[EngagementComments Properties (Project)](http://msdn.microsoft.com/library/335727af-b898-0629-adee-e3c9d6c89fc4%28Office.15%29.aspx)**||
|**[EngagementComments.Add Method (Project)](http://msdn.microsoft.com/library/a36d5592-068f-3cda-c4e5-301ddbe1cbbb%28Office.15%29.aspx)**||
|**[EngagementComments.Application Property (Project)](http://msdn.microsoft.com/library/12894229-3a2c-2b1b-2d31-39da1fc3b443%28Office.15%29.aspx)**||
|**[EngagementComments.Count Property (Project)](http://msdn.microsoft.com/library/8767e8f8-7e89-5644-a53a-5d28e34dc75d%28Office.15%29.aspx)**||
|**[EngagementComments.Item Property (Project)](http://msdn.microsoft.com/library/04bdc594-4bc4-0d5a-354d-e53c5dacdb5a%28Office.15%29.aspx)**||
|**[EngagementComments.Parent Property (Project)](http://msdn.microsoft.com/library/5d8aec1a-c197-5bf1-4461-b580bd8dd5a8%28Office.15%29.aspx)**||
|**[Engagements Members (Project)](http://msdn.microsoft.com/library/a1851a7d-96e5-c523-4ccb-66c5a91220b0%28Office.15%29.aspx)**||
|**[Engagements Methods (Project)](http://msdn.microsoft.com/library/d8eb0c5a-dc7d-1e42-ac23-882b3bfd06e1%28Office.15%29.aspx)**||
|**[Engagements Object (Project)](http://msdn.microsoft.com/library/4986802b-1d53-7bc6-0bc7-6a5b83855628%28Office.15%29.aspx)**||
|**[Engagements Properties (Project)](http://msdn.microsoft.com/library/0a07b5a0-8eae-3b6e-e290-25ab0ef63b32%28Office.15%29.aspx)**||
|**[Engagements.Add Method (Project)](http://msdn.microsoft.com/library/c3871f6a-ce2f-d0ae-ed91-658afaae25dd%28Office.15%29.aspx)**||
|**[Engagements.Application Property (Project)](http://msdn.microsoft.com/library/6e4c0204-6955-9298-e47a-357f1a600b5f%28Office.15%29.aspx)**||
|**[Engagements.Count Property (Project)](http://msdn.microsoft.com/library/e0d95ca6-50e9-c180-81bb-d1579a6d2405%28Office.15%29.aspx)**||
|**[Engagements.Item Property (Project)](http://msdn.microsoft.com/library/959abd12-3c55-25b9-2411-36a5b1f3bed7%28Office.15%29.aspx)**||
|**[Engagements.Parent Property (Project)](http://msdn.microsoft.com/library/dfd17c98-de11-ab6d-b7bb-9c0df3b1114e%28Office.15%29.aspx)**||
|**[Engagements.UniqueID Property (Project)](http://msdn.microsoft.com/library/35e9e64a-5ab9-ffda-2002-cb5a2b40eb7e%28Office.15%29.aspx)**||
|**[PjAssignmentWarnings Enumeration (Project)](http://msdn.microsoft.com/library/ecc27ae7-cc86-21aa-8c7f-aed8a7d22d38%28Office.15%29.aspx)**|Defines the different types of warnings that may appear on assignments triggering indicators in the indicator column in sheet views.|
|**[PjEngagementViolationType Enumeration (Project)](http://msdn.microsoft.com/library/e65cf9c5-e122-a4ef-f8c1-efb27899e27b%28Office.15%29.aspx)**|Defines the different types of engagement violation types triggering indicators in the indicator column in sheet views on tasks/resources and assignments. Used internally for setting the right violation types on tasks and resources.|
|**[PjEngagementWarnings Enumeration (Project)](http://msdn.microsoft.com/library/6a6b606b-6040-e278-642a-aa54fb690be2%28Office.15%29.aspx)**|Defines the different types of warnings that may appear on engagements triggering indicators in the indicator column in sheet views.|
|**[PjResourceWarnings Enumeration (Project)](http://msdn.microsoft.com/library/91d4ddd9-8ca2-e1e8-2948-37a856f944b6%28Office.15%29.aspx)**|Defines the different types of warnings that may appear on resources triggering indicators in the indicator column in sheet views. |
|**[Project.Engagements Property (Project)](http://msdn.microsoft.com/library/00ebca26-b9f6-05e4-f0ab-ba54b9dc0124%28Office.15%29.aspx)**|Returns the root object for all Engagement properties.|
|**[Project.LastWssSyncDate Property (Project)](http://msdn.microsoft.com/library/fc8aadd9-0ab1-b0b3-1ebc-7f1ef8dbe945%28Office.15%29.aspx)**|Returns the last date on which Project was synced with Wss. Read-only  **DateType**.|
|**[Project.Timeline Property (Project)](http://msdn.microsoft.com/library/6e463f3b-28fb-79dc-c51f-c3512183a310%28Office.15%29.aspx)**|Returns the root object for all Timeline properties. Read/write  **object**.|
|**[Project.UtilizationDate Property (Project)](http://msdn.microsoft.com/library/f63aced1-4bdf-585e-ae72-92d6f45699b7%28Office.15%29.aspx)**|Used for portfolio analysis, such as Project Plan, Resource Engagements, or Project Plan until. Read-only. Project Plan uses the project plan to calculate resource availability, Resource Engagements uses Resource Engagements, and Project Plan until is a combination of Project Plan and Resource Engagements.|
|**[Project.UtilizationType Property (Project)](http://msdn.microsoft.com/library/5b6aa424-b84d-4ad6-c6e5-d7a54a63a63f%28Office.15%29.aspx)**|If the Project.UtilizationType Property (Project) property is Project Plan until, this date is used to switch between using the project plan to calculate resource availability or when resource engagements are used. Read-only.|
|**[Resource.Compliant Property (Project)](http://msdn.microsoft.com/library/269f8d19-ff75-a017-9a84-1c5889918ea8%28Office.15%29.aspx)**|**True** if the resource is compliant with its engagement. Read/write **Boolean**.|
|**[Resource.EngagementCommittedFinish Property (Project)](http://msdn.microsoft.com/library/8507b1d2-095e-b0ab-aaa4-58fbf4037cab%28Office.15%29.aspx)**|Returns the committed finish date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementCommittedMaxUnits Property (Project)](http://msdn.microsoft.com/library/571c89e8-5168-eb41-c995-77371a7e4039%28Office.15%29.aspx)**|Returns the committed max units for the engagement. Read-only  **Integer**.|
|**[Resource.EngagementCommittedStart Property (Project)](http://msdn.microsoft.com/library/a9cd2ee0-aee7-2048-cf58-361e091ecf4c%28Office.15%29.aspx)**|Returns the committed start date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementCommittedWork Property (Project)](http://msdn.microsoft.com/library/60fc5fdf-4341-f034-61ae-2055515150c9%28Office.15%29.aspx)**|Returns the committed work for the engagement. Read-only  **Double**.|
|**[Resource.EngagementDraftFinish Property (Project)](http://msdn.microsoft.com/library/46d61f27-eacd-2546-637a-695be8c7d98d%28Office.15%29.aspx)**|Returns the draft finish date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementDraftMaxUnits Property (Project)](http://msdn.microsoft.com/library/d3129ad0-2883-1294-d494-c1927121fc2c%28Office.15%29.aspx)**|Returns the draft max units for the engagement. Read-only  **Integer**.|
|**[Resource.EngagementDraftStart Property (Project)](http://msdn.microsoft.com/library/54f5399f-e7df-0a5a-6008-054423d30be8%28Office.15%29.aspx)**|Returns the draft start date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementDraftWork Property (Project)](http://msdn.microsoft.com/library/577433e5-7d23-5ce8-f3e8-18d0da9e67ac%28Office.15%29.aspx)**|Returns the draft work for the engagement. Read-only  **Double**.|
|**[Resource.EngagementProposedFinish Property (Project)](http://msdn.microsoft.com/library/ab8917ef-edb5-592b-f87f-8db9aefc85ff%28Office.15%29.aspx)**|Returns the proposed finish date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementProposedMaxUnits Property (Project)](http://msdn.microsoft.com/library/635dd13d-7b4f-0325-638b-af0a1baf505d%28Office.15%29.aspx)**|Returns the proposed maximum units for the engagement. Read-only  **Integer**.|
|**[Resource.EngagementProposedStart Property (Project)](http://msdn.microsoft.com/library/f00b3441-1f20-3f66-9c0c-208a04d3f6a2%28Office.15%29.aspx)**|Returns the proposed start date for the engagement. Read-only  **DateType**.|
|**[Resource.EngagementProposedWork Property (Project)](http://msdn.microsoft.com/library/cb1e6aae-cbc6-3ee0-9e0f-c75c9954ec66%28Office.15%29.aspx)**|Returns the proposed work for the engagement. Read-only  **Double**.|
|**[Resource.IsLocked Property (Project)](http://msdn.microsoft.com/library/56525d08-e779-b9e4-c41a-24664ed68538%28Office.15%29.aspx)**|Indicates whether the resource is or is not locked. If resource is locked, you need an engagement for a resource. Read-only Return  **Boolean**.|
|**[Task.Compliant Property (Project)](http://msdn.microsoft.com/library/d2e43c4a-a7c6-c179-70f3-c67b813be3b8%28Office.15%29.aspx)**||
|**[Timeline Members (Project)](http://msdn.microsoft.com/library/ac50eced-d876-ee09-f8f4-01fb2272ddf0%28Office.15%29.aspx)**||
|**[Timeline Object (Project)](http://msdn.microsoft.com/library/8e02e775-1999-edf8-e724-02e4a0d59bad%28Office.15%29.aspx)**||
|**[Timeline Properties (Project)](http://msdn.microsoft.com/library/37a5da93-b5b1-dca7-2fac-0cdd5baa3f61%28Office.15%29.aspx)**||
|**[Timeline.Application Property (Project)](http://msdn.microsoft.com/library/4e9beeb2-5fd9-3631-b60e-1f41666f50b4%28Office.15%29.aspx)**|Gets the Project  **Application** object.|
|**[Timeline.BarCount Property (Project)](http://msdn.microsoft.com/library/8c4f6fa2-62d5-3be4-a4e8-0b3301d1fd85%28Office.15%29.aspx)**|Indicates the number of bars in a  **Timeline** view.|
|**[Timeline.FinishDate Property (Project)](http://msdn.microsoft.com/library/d0f51644-63ba-9e7f-2da3-92995ec73551%28Office.15%29.aspx)**|Indicates the finish date for a  **Timeline** bar based on the input argument.|
|**[Timeline.Label Property (Project)](http://msdn.microsoft.com/library/8456d32e-c389-232a-2279-e7f73b4cd05e%28Office.15%29.aspx)**|Returns the timeline for the  **Timeline** object.|
|**[Timeline.StartDate Property (Project)](http://msdn.microsoft.com/library/960deebd-d7c3-eee0-2658-ba170bf40fcd%28Office.15%29.aspx)**|Indicates the start date for a  **Timeline** bar based on the input argument.|

## PowerPoint





|**Name**|**Description**|
|:-----|:-----|
|**[ChartGroup.BinsCountValue Property (PowerPoint)](http://msdn.microsoft.com/library/7af8e5a7-ec62-f447-275c-b694b7ed37b7%28Office.15%29.aspx)**|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|**[ChartGroup.BinsOverflowEnabled Property (PowerPoint)](http://msdn.microsoft.com/library/9d5e5296-b80c-f6dd-b418-1d0cd3a9adce%28Office.15%29.aspx)**|Specifies whether a bin for values above the ChartGroup.BinsOverflowValue Property (PowerPoint) is enabled. Read/write  **Boolean**.|
|**[ChartGroup.BinsOverflowValue Property (PowerPoint)](http://msdn.microsoft.com/library/59f9b37a-8736-bd1e-9e71-0e324a10e646%28Office.15%29.aspx)**|If an [ChartGroup.BinsOverflowEnabled](http://msdn.microsoft.com/library/9d5e5296-b80c-f6dd-b418-1d0cd3a9adce%28Office.15%29.aspx) Property (PowerPoint) is **True**, specifies the value above which an overflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinsType Property (PowerPoint)](http://msdn.microsoft.com/library/f43cce63-dfad-aed3-2dfa-2359a9e5a728%28Office.15%29.aspx)**|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](http://msdn.microsoft.com/library/a9f49fcc-4c7c-5097-ab7f-0a233df415d0%28Office.15%29.aspx) Enumeration (PowerPoint).|
|**[ChartGroup.BinsUnderflowEnabled Property (PowerPoint)](http://msdn.microsoft.com/library/42b53b36-5a40-ac5d-cf2c-7658128006ca%28Office.15%29.aspx)**|Specifies whether a bin for values below the [ChartGroup.BinsUnderflowValue](http://msdn.microsoft.com/library/93a0ccff-c132-311a-7992-83d7adce3938%28Office.15%29.aspx) Property (PowerPoint) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsUnderflowValue Property (PowerPoint)](http://msdn.microsoft.com/library/93a0ccff-c132-311a-7992-83d7adce3938%28Office.15%29.aspx)**|If [ChartGroup.BinsUnderflowEnabled](http://msdn.microsoft.com/library/42b53b36-5a40-ac5d-cf2c-7658128006ca%28Office.15%29.aspx) Property (PowerPoint) is True, specifies the value below which an underflow bin is displayed. Read/write Double.|
|**[ChartGroup.BinWidthValue Property (PowerPoint)](http://msdn.microsoft.com/library/03224e89-65aa-76ff-68db-31be8465bd34%28Office.15%29.aspx)**|Specifies the number of points in each range. Read/write  **Double**.|
|**[DocumentWindow.ShowInsertAppDialog Method (PowerPoint)](http://msdn.microsoft.com/library/8ad060d1-ae80-6011-7fba-66f87d89d158%28Office.15%29.aspx)**||
|**[Point.IsTotal Property (PowerPoint)](http://msdn.microsoft.com/library/3692deb0-71cd-2cfc-163a-3ab4fe831a04%28Office.15%29.aspx)**|**True** if the point represents a total. Read/write **Boolean**.|
|**[Series.ParentDataLabelOption Property (PowerPoint)](http://msdn.microsoft.com/library/678ad97d-725b-5a4c-b3a4-294e9f905e5f%28Office.15%29.aspx)**|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XlParentDataLabelOptions](http://msdn.microsoft.com/library/566194d6-f4e3-53af-723c-025bf3909578%28Office.15%29.aspx) Enumeration (PowerPoint).|
|**[Series.QuartileCalculationInclusiveMedian Property (PowerPoint)](http://msdn.microsoft.com/library/0c6e80be-22f6-8e7e-437c-7c9066e0886d%28Office.15%29.aspx)**|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|**[Shape.HasInkXML Property (PowerPoint)](http://msdn.microsoft.com/library/3d985f9b-64e3-8712-fd5f-73d38ca56810%28Office.15%29.aspx)**|Returns an [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx) enumeration value that indicates whether the specified shape contains ink XML that can be retrieved via the[Shape.InkXML](http://msdn.microsoft.com/library/01e01d61-89a3-1314-fda5-6354d6590aa5%28Office.15%29.aspx) property. Read-only. An error is returned if the shape does not contain any ink XML.|
|**[Shape.InkXML Property (PowerPoint)](http://msdn.microsoft.com/library/01e01d61-89a3-1314-fda5-6354d6590aa5%28Office.15%29.aspx)**|Returns a  **String** that contains the InkActionML associated with the specified shape. Read-only. If the specified shape does not contain a ink object more than one ink object occurs , an error is returned.|
|**[Shape.IsNarration Property (PowerPoint)](http://msdn.microsoft.com/library/e07e42e3-149d-153f-6852-a41c0eae80e3%28Office.15%29.aspx)**|Specifies whether the specified shape range contains a narration. Read/write.|
|**[ShapeRange.HasInkXML Property (PowerPoint)](http://msdn.microsoft.com/library/1a7b7b8b-c0e8-8f62-1015-e99cb590fd50%28Office.15%29.aspx)**|Returns an [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx) enumeration value that indicates whether the specified shape range contains ink XML that can be retrieved via the[ShapeRange.InkXML](http://msdn.microsoft.com/library/faff227c-293a-58cf-fe49-eb8b5f5caac3%28Office.15%29.aspx) property. Read-only. An error is returned if the shape range does not contain any ink XML.|
|**[ShapeRange.InkXML Property (PowerPoint)](http://msdn.microsoft.com/library/faff227c-293a-58cf-fe49-eb8b5f5caac3%28Office.15%29.aspx)**|Returns a  **String** that contains the InkActionML associated with the specified shape range. Read-only. If the specified shape range does not contain a ink object more than one ink object occurs , an error is returned.|
|**[ShapeRange.IsNarration Property (PowerPoint)](http://msdn.microsoft.com/library/a82b4156-9025-aa7c-b132-b7f5cafa2b3b%28Office.15%29.aspx)**|Specifies whether the specified shape range contains a narration. Read/write. |
|**[Shapes.AddInkShapeFromXML Method (PowerPoint)](http://msdn.microsoft.com/library/88a395ac-b11e-d42e-f4b4-b41bf1d1347e%28Office.15%29.aspx)**|Creates an ink shape. Returns a [Shape](http://msdn.microsoft.com/library/1da93849-99e0-827e-ced3-c6cf7f8569f3%28Office.15%29.aspx) object that represents the new ink shape.|
|**[SlideShowView.LaserPointerEnabled Property (PowerPoint)](http://msdn.microsoft.com/library/9ba56542-a2bf-28d2-9609-50f9a4144c91%28Office.15%29.aspx)**|Returns  **true** if the current slide show pointer is a laser pointer. This property is applicable only while the slide show is running. Read/write. This property allows a user to programmatically query and set the state of the pointer shown during slide show. The property will return false for all other pointer types. Users can also change the state of the current pointer by setting this property to **true** to turn on the laser pointer or **false** to turn off the laser pointer.|
|**[XlBinsType Enumeration (PowerPoint)](http://msdn.microsoft.com/library/a9f49fcc-4c7c-5097-ab7f-0a233df415d0%28Office.15%29.aspx)**|Constants passed to and returned by the [ChartGroup.BinsType](http://msdn.microsoft.com/library/7230c44b-2e93-9790-2f27-d584688c6172%28Office.15%29.aspx) property.|
|**[XlParentDataLabelOptions Enumeration (PowerPoint)](http://msdn.microsoft.com/library/566194d6-f4e3-53af-723c-025bf3909578%28Office.15%29.aspx)**|Constants passed to and returned by the  **Series.ParentDataLabelOption** property.|

## Visio





|**Name**|**Description**|
|:-----|:-----|
|**[Document.Permission Property (Visio)](http://msdn.microsoft.com/library/944f11be-053c-7749-178c-5e8b79a32ea8%28Office.15%29.aspx)**||
|**[IVInvisibleApp.Application Property (Visio)](http://msdn.microsoft.com/library/5f84769f-404f-9766-4d25-41ef7dfed324%28Office.15%29.aspx)**||
|**[IVKeyboardEvent.Application Property (Visio)](http://msdn.microsoft.com/library/9eefdda1-02c9-e256-a57a-1862a59695cf%28Office.15%29.aspx)**||
|**[IVMouseEvent.Application Property (Visio)](http://msdn.microsoft.com/library/dc74f482-2807-3480-8bfc-e8b915f0dff8%28Office.15%29.aspx)**||
|**[Master.VisualBoundingBox Method (Visio)](http://msdn.microsoft.com/library/478d636f-e741-cf6b-3e16-b5faf70a9f14%28Office.15%29.aspx)**|Returns the bounding rectangle of the virtual container that has all the shapes of the given master.|
|**[Page.VisualBoundingBox Method (Visio)](http://msdn.microsoft.com/library/95e8a977-55c9-307a-bade-120cb8acdf9b%28Office.15%29.aspx)**|Returns the bounding rectangle of the virtual container that has all the shapes of the given page.|
|**[Selection.VisualBoundingBox Method (Visio)](http://msdn.microsoft.com/library/ae107bd8-ac99-6303-2820-a5afb19165a3%28Office.15%29.aspx)**|Returns the bounding rectangle of the virtual container that has all the shapes of the given selection.|
|**[Shape.VisualBoundingBox Method (Visio)](http://msdn.microsoft.com/library/6a7d4622-8ba5-c598-4aaa-c6297cb4c008%28Office.15%29.aspx)**|Returns the bounding rectangle of the given shape.|
|**[ValidationIssues.Stat Property (Visio)](http://msdn.microsoft.com/library/bf0731f1-fd5e-d2e3-489c-17efeab04291%28Office.15%29.aspx)**||
|**[VisColoringMethod Enumeration (Visio)](http://msdn.microsoft.com/library/d39f02dc-36ef-fdd4-62b1-0bfc4d7d2433%28Office.15%29.aspx)**||
|**[VisRecordsetFieldStatus Enumeration (Visio)](http://msdn.microsoft.com/library/532bc905-1227-eca4-ec7b-d87f7dfb8bb6%28Office.15%29.aspx)**||

## Word





|**Name**|**Description**|
|:-----|:-----|
|**[ChartGroup.BinsCountValue Property (Word)](http://msdn.microsoft.com/library/0a65eebd-7818-579d-4e4b-df50c0608cfa%28Office.15%29.aspx)**|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|**[ChartGroup.BinsOverflowEnabled Property (Word)](http://msdn.microsoft.com/library/a208e251-b2c8-04cc-f40f-97715f55fa33%28Office.15%29.aspx)**|Specifies whether a bin for values above the [BinsOverflowValue](http://msdn.microsoft.com/library/411856a7-ac17-e9eb-35bd-c851c0cfdfdc%28Office.15%29.aspx) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsOverflowValue Property (Word)](http://msdn.microsoft.com/library/288b119a-7a76-2b56-4181-9d39a5be397f%28Office.15%29.aspx)**|If an [BinsOverflowEnabled](http://msdn.microsoft.com/library/3af8d552-94e1-6f15-df2b-38fb7d3a0be1%28Office.15%29.aspx) is **True**, specifies the value above which an overflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinsType Property (Word)](http://msdn.microsoft.com/library/a403cac5-a397-e202-1dda-5b31e3815ef0%28Office.15%29.aspx)**|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](http://msdn.microsoft.com/library/945e729b-f0a0-fc0f-d198-c85aab081d7e%28Office.15%29.aspx).|
|**[ChartGroup.BinsUnderflowEnabled Property (Word)](http://msdn.microsoft.com/library/7ffe9878-2462-8d05-7158-24ba45107b31%28Office.15%29.aspx)**|Specifies whether a bin for values below the [BinsUnderflowValue](http://msdn.microsoft.com/library/40143963-c9a9-566e-e8aa-722cad0db0fc%28Office.15%29.aspx) is enabled. Read/write **Boolean**.|
|**[ChartGroup.BinsUnderflowValue Property (Word)](http://msdn.microsoft.com/library/40143963-c9a9-566e-e8aa-722cad0db0fc%28Office.15%29.aspx)**|If an [BinsUnderflowEnabled](http://msdn.microsoft.com/library/7ffe9878-2462-8d05-7158-24ba45107b31%28Office.15%29.aspx) is **True**, specifies the value below which an underflow bin is displayed. Read/write **Double**.|
|**[ChartGroup.BinWidthValue Property (Word)](http://msdn.microsoft.com/library/cda366d4-48ef-4ca9-62a8-f2d2f8843936%28Office.15%29.aspx)**|Specifies the number of points in each range. Read/write  **Double**.|
|**[CoAuthUpdates Object (Word)](http://msdn.microsoft.com/library/afd0abeb-276e-96f4-ee8a-01f263e69121%28Office.15%29.aspx)**|A collection of [CoAuthUpdate](http://msdn.microsoft.com/library/c00e5029-2e4b-97c0-33d3-86fdc53df535%28Office.15%29.aspx) objects that represent the updates that were merged into the document at the last explicit save.|
|**[Options.UseLocalUserInfo Property (Word)](http://msdn.microsoft.com/library/886bd7ce-8f3b-31f0-aacd-10f240b1bf88%28Office.15%29.aspx)**||
|**[Point.IsTotal Property (Word)](http://msdn.microsoft.com/library/58d203fd-1e7f-b14b-4eaa-f25a0494c5ea%28Office.15%29.aspx)**|**True** if the point represents a total. Read/write **Boolean**.|
|**[Series.ParentDataLabelOption Property (Word)](http://msdn.microsoft.com/library/e3b1e3a4-b775-2daa-56aa-094e8cc9a86b%28Office.15%29.aspx)**|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XLParentDataLabelOptions](http://msdn.microsoft.com/library/c83fe64d-5a14-74b5-5847-62cba83805b0%28Office.15%29.aspx).|
|**[Series.QuartileCalculationInclusiveMedian Property (Word)](http://msdn.microsoft.com/library/b539e619-1dc8-6419-28ba-3ab20b64c2b1%28Office.15%29.aspx)**|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|**[XlBinsType Enumeration (Word)](http://msdn.microsoft.com/library/945e729b-f0a0-fc0f-d198-c85aab081d7e%28Office.15%29.aspx)**|Constants passed to and returned by the [ChartGroup.BinsType](http://msdn.microsoft.com/library/a403cac5-a397-e202-1dda-5b31e3815ef0%28Office.15%29.aspx) property.|
|**[XlParentDataLabelOptions Enumeration (Word)](http://msdn.microsoft.com/library/c83fe64d-5a14-74b5-5847-62cba83805b0%28Office.15%29.aspx)**|Constants passed to and returned by the  **Series.ParentDataLabelOption** property.|

