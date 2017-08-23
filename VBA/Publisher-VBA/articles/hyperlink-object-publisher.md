---
title: "Объект гиперссылки (издатель)"
keywords: vbapb10.chm4653055
f1_keywords: vbapb10.chm4653055
ms.prod: publisher
api_name: Publisher.Hyperlink
ms.assetid: 1cc6d95b-357a-c169-a5d2-6850a1a3bbd6
ms.date: 06/08/2017
ms.openlocfilehash: a23cfa04844deceb2ef5dba7e76234744f081f7f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlink-object-publisher"></a>Объект гиперссылки (издатель)

Представляет гиперссылки. Объект **гиперссылки** , является участником коллекции **[гиперссылки](hyperlinks-object-publisher.md)** и **[фигуры](http://msdn.microsoft.com/library/666cb7f0-62a8-f419-9838-007ef29506ee%28Office.15%29.aspx)** и **[ShapeRange](shaperange-object-publisher.md)** объекты.


## <a name="example"></a>Пример

Используйте свойство **[гиперссылку](http://msdn.microsoft.com/library/0990ab32-b4a3-6c89-cb9f-8f8c64ef804f%28Office.15%29.aspx)** для возврата объекта **гиперссылок** , связанных с фигурой (фигура может иметь только одну гиперссылку). В следующем примере удаляется гиперссылки, связанной с первой фигуры в активный документ.


```
Sub DeleteHyperlink() 
 ActiveDocument.Pages(1).Shapes(1).Hyperlink.Delete 
End Sub
```

Использование **гиперссылок** (индекс), где индекс — номер индекса, для возврата объекта **гиперссылки** из документа, диапазон или выделить фрагмент. В следующем примере удаляется гиперссылки в выделение.




```
Sub DeleteSelectedHyperlink() 
 If Selection.TextRange.Hyperlinks.Count >= 1 Then 
 Selection.TextRange.Hyperlinks(1).Delete 
 End If 
End Sub
```

Используйте метод **[Add](http://msdn.microsoft.com/library/f5a8cc01-a571-623d-bfab-fe48e43a21b1%28Office.15%29.aspx)** для добавления гиперссылки. Следующий пример добавляет выбранный текст гиперссылки.




```
Sub AddHyperlinkToSelectedText() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="http://www.tailspintoys.com/" 
End Sub
```

Используйте свойство **[адрес](http://msdn.microsoft.com/library/784a9213-38bc-c5fd-f215-abeb174ec628%28Office.15%29.aspx)** добавить или изменить адрес гиперссылки. В следующем примере добавляется фигура active публикации и затем добавляет гиперссылки в фигуры.




```
Sub AddHyperlinkToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=200, Width:=300, Height:=300) 
 .Hyperlink.Address = "http://www.tailspintoys.com/" 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/28b7f351-c1a8-29f1-2114-ed6854fbd13a%28Office.15%29.aspx)|
|[SetPageRelative](http://msdn.microsoft.com/library/4b2f2e84-09ce-cef6-6f22-b82642cc71fe%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Адрес](http://msdn.microsoft.com/library/784a9213-38bc-c5fd-f215-abeb174ec628%28Office.15%29.aspx)|
|[Приложения](http://msdn.microsoft.com/library/dadf9b35-580e-c184-c439-38b3a4f1529f%28Office.15%29.aspx)|
|[EmailSubject](http://msdn.microsoft.com/library/16b60648-56fe-b8ba-3424-0dd6e88727e6%28Office.15%29.aspx)|
|[PageID](http://msdn.microsoft.com/library/1b5051eb-e6b4-a5a7-610a-5be03863a92b%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/a0e3ab66-cdc4-09ab-6995-8a5e0194d6e2%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/ff105ffe-cb48-0f6a-99ff-eaac0500938f%28Office.15%29.aspx)|
|[Фигура](http://msdn.microsoft.com/library/afd1dab7-472a-2aa5-f5da-1e2f783b5270%28Office.15%29.aspx)|
|[TargetType](http://msdn.microsoft.com/library/1cbc8c36-563c-4464-4f0d-2836682ce532%28Office.15%29.aspx)|
|[TextToDisplay](http://msdn.microsoft.com/library/26b5857c-3f94-0d33-f65e-9c34f2a4cc2b%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/6a9ac3c4-4f34-d759-af95-a3bdc510a56f%28Office.15%29.aspx)|

