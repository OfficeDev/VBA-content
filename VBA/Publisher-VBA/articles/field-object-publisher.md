---
title: "Объект поля (издатель)"
keywords: vbapb10.chm6160383
f1_keywords: vbapb10.chm6160383
ms.prod: publisher
api_name: Publisher.Field
ms.assetid: 93da311a-b834-f990-60e9-786d4f6a16f1
ms.date: 06/08/2017
ms.openlocfilehash: 22ea875efca4dd9077bb1c1ba9cd278f9386a6b7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="field-object-publisher"></a>Объект поля (издатель)

Представляет поле. Объект **поля** , является участником коллекции **[полей](fields-object-publisher.md)** . Коллекции **полей** представляет поля выбора, диапазон или публикации.
 


## <a name="remarks"></a>Заметки

Константа **pbFieldPageNumber** , принадлежит к группе **PbFieldType** константы, который включает различные типы полей.
 

 

## <a name="example"></a>Пример

Использование **[полей](textrange-fields-property-publisher.md)** (индекс), где индекс — номер индекса, для возврата объекта **поля** . Номер индекса представляет позицию поля выбора, диапазон или публикации. Следующие подсчитывает количество полей в активной публикации и отображает счетчик в сообщении.
 

 

```
Sub CountFields() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 Dim fldField As Field 
 Dim intFields As Integer 
 Dim intCount As Integer 
 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = pbTextFrame Then 
 intCount = intCount + shpShape.TextFrame.TextRange.Fields.Count 
 End If 
 Next 
 Next 
 If intCount > 0 Then 
 MsgBox "You have " &amp; intCount &amp; " fields in your publication." 
 Else 
 MsgBox "You have no fields in your publication." 
 End If 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Разорвать связь](field-unlink-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](field-application-property-publisher.md)|
|[Код](field-code-property-publisher.md)|
|[Далее](field-next-property-publisher.md)|
|[Родительский раздел](field-parent-property-publisher.md)|
|[PhoneticGuide](field-phoneticguide-property-publisher.md)|
|[Результат](field-result-property-publisher.md)|
|[TextRange](field-textrange-property-publisher.md)|
|[Type](field-type-property-publisher.md)|

