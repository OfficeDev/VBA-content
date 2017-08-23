---
title: "Объект LinkFormat (издатель)"
keywords: vbapb10.chm4456447
f1_keywords: vbapb10.chm4456447
ms.prod: publisher
api_name: Publisher.LinkFormat
ms.assetid: 5b588edd-b026-cfc7-4acb-77290ae4d297
ms.date: 06/08/2017
ms.openlocfilehash: 18025fdd44a5512a97bf787852eda611638f4697
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="linkformat-object-publisher"></a>Объект LinkFormat (издатель)

Представляет связывания характеристики для объекта OLE или рисунок.
 


## <a name="remarks"></a>Заметки

Не все типы фигур и полей могут быть связаны с источника. Используйте свойство **[Type](shape-type-property-publisher.md)** для объекта **[Shape](shape-object-publisher.md)** для определения, могут быть связаны определенного фигуры.
 

 
Используйте метод **[обновления](linkformat-update-method-publisher.md)** для обновления ссылок. Для возвращения или задания полный путь для данной ссылки исходного файла, используйте свойство **[SourceFullName](linkformat-sourcefullname-property-publisher.md)** .
 

 

## <a name="example"></a>Пример

Используйте свойство **[LinkFormat](shape-linkformat-property-publisher.md)** для поля или фигуры для возврата объекта **LinkFormat** . В следующем примере обновляются ссылки на все связанные объекты OLE на первой странице active публикации.
 

 

```
Sub FindOLEObjects() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Pages(1).Shapes 
 If shpShape.Type = pbLinkedOLEObject Then 
 shpShape.LinkFormat.Update 
 End If 
 Next shpShape 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Update](linkformat-update-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](linkformat-application-property-publisher.md)|
|[Родительский раздел](linkformat-parent-property-publisher.md)|
|[SourceFullName](linkformat-sourcefullname-property-publisher.md)|

