---
title: "Объект поля (издатель)"
keywords: vbapb10.chm6094847
f1_keywords: vbapb10.chm6094847
ms.prod: publisher
api_name: Publisher.Fields
ms.assetid: fd7c95d9-bc34-95ee-180d-b99f3629eb33
ms.date: 06/08/2017
ms.openlocfilehash: a8c1d7639e7fde3763f26e4a4a4200b7ff4b7537
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fields-object-publisher"></a>Объект поля (издатель)

Коллекция объектов **[поля](field-object-publisher.md)** , которые представляют всех полей в диапазон текста.
 


## <a name="remarks"></a>Заметки

Свойство **[Count](fields-count-property-publisher.md)** для данного семейства сайтов в публикации возвращает число элементов в указанной фигуры или выделить фрагмент.
 

 

## <a name="example"></a>Пример

Используйте свойство **[полей](textrange-fields-property-publisher.md)** для возврата коллекции **полей** . Использование **полей** (индекс), где индекс — номер индекса, для возврата объекта **поля** . Номер индекса представляет позицию поля выбора, диапазон или публикации. В следующем примере отображается код и результатов первое поле в каждом текстовом поле в активной публикации.
 

 

```
Sub ShowFieldCodes() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = pbTextFrame Then 
 With shpShape.TextFrame.TextRange 
 If .Fields.Count > 0 Then 
 MsgBox "Code = " &amp; .Fields(1).Code &amp; vbLf _ 
 &amp; "Result = " &amp; .Fields(1).Result &amp; vbLf 
 End If 
 End With 
 End If 
 Next 
 Next 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[AddHorizontalInVertical](fields-addhorizontalinvertical-method-publisher.md)|
|[AddPhoneticGuide](fields-addphoneticguide-method-publisher.md)|
|[Элемент](fields-item-method-publisher.md)|
|[Разорвать связь](fields-unlink-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](fields-application-property-publisher.md)|
|[Count](fields-count-property-publisher.md)|
|[Родительский раздел](fields-parent-property-publisher.md)|

