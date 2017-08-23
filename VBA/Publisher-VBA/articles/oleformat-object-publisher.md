---
title: "Объект OLEFormat (издатель)"
keywords: vbapb10.chm4521983
f1_keywords: vbapb10.chm4521983
ms.prod: publisher
api_name: Publisher.OLEFormat
ms.assetid: e5b72d6b-dff8-3882-549f-e376c1e4d372
ms.date: 06/08/2017
ms.openlocfilehash: 964324f58368f8d71115b5bbb5051887c8703cfd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="oleformat-object-publisher"></a>Объект OLEFormat (издатель)

Представляет характеристики OLE, отличный от ссылки (виден **[LinkFormat](linkformat-object-publisher.md)** ), для OLE-объект, элемент управления ActiveX или поля.
 


## <a name="remarks"></a>Заметки

Не все типы фигур и полей имеют возможности OLE. Используйте свойство **[Type](shape-type-property-publisher.md)** для объекта **[Shape](shape-object-publisher.md)** для определения, какая категория принадлежит указанный фигуры.
 

 
Используйте методы **[активировать](oleformat-activate-method-publisher.md)** и **[DoVerb](oleformat-doverb-method-publisher.md)** для автоматизации объекта OLE.
 

 

## <a name="example"></a>Пример

Используйте свойство **[OLEFormat](shape-oleformat-property-publisher.md)** для поля или фигуры для возврата объекта **OLEFormat** . Следующий пример активирует все объекты OLE в активной публикации.
 

 

```
Sub ActivateOLEObjects() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Pages(1).Shapes 
 If shpShape.Type = pbLinkedOLEObject Then 
 shpShape.OLEFormat.Activate 
 End If 
 Next 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Активация](oleformat-activate-method-publisher.md)|
|[DoVerb](oleformat-doverb-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](oleformat-application-property-publisher.md)|
|[Object](oleformat-object-property-publisher.md)|
|[ObjectVerbs](oleformat-objectverbs-property-publisher.md)|
|[Родительский раздел](oleformat-parent-property-publisher.md)|
|[ProgId](oleformat-progid-property-publisher.md)|

