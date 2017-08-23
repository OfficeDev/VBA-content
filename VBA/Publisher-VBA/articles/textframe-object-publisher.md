---
title: "Объект TextFrame (издатель)"
keywords: vbapb10.chm3932159
f1_keywords: vbapb10.chm3932159
ms.prod: publisher
api_name: Publisher.TextFrame
ms.assetid: 95e88f5a-b3dc-272e-7c1d-5282c97ae11e
ms.date: 06/08/2017
ms.openlocfilehash: b103fdc519da71f66764b27feb73cc4a7bc92de8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframe-object-publisher"></a>Объект TextFrame (издатель)

Представляет кадр текста в объект **[фигуры](http://msdn.microsoft.com/library/666cb7f0-62a8-f419-9838-007ef29506ee%28Office.15%29.aspx)** . Содержит текст в текстовой рамки и свойства, которые управляют полей и ориентации рамки.


## <a name="example"></a>Пример

Свойство **[TextFrame](http://msdn.microsoft.com/library/fc654905-d56b-9a6c-28fa-4b54bf2a8686%28Office.15%29.aspx)** используется для возврата объекта **TextFrame** для фигуры. Свойство **[TextRange](http://msdn.microsoft.com/library/44a8395e-81dc-7d06-f068-89f77a889f5e%28Office.15%29.aspx)** возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий диапазон текста в элементе frame указанный текст. В следующем примере добавляется текст надписи фигуры один активный публикации и форматирует новый текст.


```
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```


 **Примечание**  Некоторые фигуры не поддерживают вложенные текст (линии, полилинии, изображения и объекты OLE, например). Если для возвращения или задания свойства, которые управляют текстом в фрагмент текста для этих объектов, возникает ошибка.

Свойство **[HasTextFrame](http://msdn.microsoft.com/library/faf9514a-438b-ad12-a830-ed34cea8ba03%28Office.15%29.aspx)** определяет, является ли фигура имеет фрагмент текста и свойство **[HasText](http://msdn.microsoft.com/library/f8d1c660-c3f1-e835-adc3-114e6611de98%28Office.15%29.aspx)** позволяет определить, содержит ли рамки текста, как показано в следующем примере.




```
Sub GetTextFromTextFrame() 
 Dim shpText As Shape 
 
 For Each shpText In ActiveDocument.Pages(1).Shapes 
 If shpText.HasTextFrame = msoTrue Then 
 With shpText.TextFrame 
 If .HasText Then MsgBox .TextRange.Text 
 End With 
 End If 
 Next 
End Sub
```

Текстовые рамки могут быть связаны друг с другом, чтобы текст перетекал из рамки одной фигуры в текстовой рамке другую фигуру. Использование свойства **[NextLinkedTextFrame](http://msdn.microsoft.com/library/5ba08ab5-8515-4efe-59a3-79a11f6a7c4e%28Office.15%29.aspx)** и **[PreviousLinkedTextFrame](http://msdn.microsoft.com/library/00947ec3-fcff-4451-491b-5b7748ccb74e%28Office.15%29.aspx)** для связывания рамок текста. В следующем примере создается текстовое поле (прямоугольник рамке) и добавляет текст. Затем создается другое текстовое поле и связывает два текстовых рамок, чтобы текст будет помещен в второй из первой текстовой рамки.




```
Sub LinkTextBoxes() 
 Dim shpTextBox1 As Shape 
 Dim shpTextBox2 As Shape 
 
 Set shpTextBox1 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 72, 72, 36) 
 shpTextBox1.TextFrame.TextRange.Text = _ 
 "This is some text. This is some more text." 
 
 Set shpTextBox2 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 144, 72, 36) 
 shpTextBox1.TextFrame.NextLinkedTextFrame = shpTextBox2 _ 
 .TextFrame 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[BreakForwardLink](http://msdn.microsoft.com/library/60a7a798-ebd3-e00d-032d-685dd0d5a042%28Office.15%29.aspx)|
|[ValidLinkTarget](http://msdn.microsoft.com/library/ee946f58-669f-7150-0f40-2dd3b857e274%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/14b41c64-cdd3-f1ab-202c-49f18d72d035%28Office.15%29.aspx)|
|[AutoFitText](http://msdn.microsoft.com/library/468a9d3e-cb9d-8147-60ea-eb839d691e7a%28Office.15%29.aspx)|
|[Столбцы](http://msdn.microsoft.com/library/b025f208-3ca4-c0f1-e01e-023931c4c545%28Office.15%29.aspx)|
|[ColumnSpacing](http://msdn.microsoft.com/library/3b650d29-3716-e9b1-eaf0-92bdc0b77c5f%28Office.15%29.aspx)|
|[HasNextLink](http://msdn.microsoft.com/library/907ec470-e283-906a-e25f-f5a8548a18a4%28Office.15%29.aspx)|
|[HasPreviousLink](http://msdn.microsoft.com/library/85e0b497-55c9-d49f-2b65-e199361c121a%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/f8d1c660-c3f1-e835-adc3-114e6611de98%28Office.15%29.aspx)|
|[IncludeContinuedFromPage](http://msdn.microsoft.com/library/7c129bf2-60da-4170-1410-94961ccf3345%28Office.15%29.aspx)|
|[IncludeContinuedOnPage](http://msdn.microsoft.com/library/defa0bd7-abe7-ac2a-97a1-de5c5f0df790%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/55858bba-1103-48ba-64d6-5cc5ab677867%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/4e784b9f-9467-5a14-c211-589e69c3b8bc%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/bdbde217-6a51-7823-ac93-8bbffa583544%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/9709eefe-0857-f228-aa56-780c4789a413%28Office.15%29.aspx)|
|[NextLinkedTextFrame](http://msdn.microsoft.com/library/5ba08ab5-8515-4efe-59a3-79a11f6a7c4e%28Office.15%29.aspx)|
|[Ориентация](http://msdn.microsoft.com/library/f510e624-6322-4054-5e7f-8688c5ea817a%28Office.15%29.aspx)|
|[Переполнения](http://msdn.microsoft.com/library/5a0f053b-519a-1637-0d73-992c56cdd7f0%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/c4d2d0bd-7a6b-201c-4b1b-416490ab8023%28Office.15%29.aspx)|
|[PreviousLinkedTextFrame](http://msdn.microsoft.com/library/00947ec3-fcff-4451-491b-5b7748ccb74e%28Office.15%29.aspx)|
|[Статья](http://msdn.microsoft.com/library/7bbe0967-83aa-745b-ad13-8a7dfe61811c%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/44a8395e-81dc-7d06-f068-89f77a889f5e%28Office.15%29.aspx)|
|[VerticalTextAlignment](http://msdn.microsoft.com/library/cd809f00-b092-c483-fe99-2aa8043fb684%28Office.15%29.aspx)|

