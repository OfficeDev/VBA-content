---
title: "Объект тега (издатель)"
keywords: vbapb10.chm4784127
f1_keywords: vbapb10.chm4784127
ms.prod: publisher
api_name: Publisher.Tag
ms.assetid: f485d2cc-8e39-5aa3-d407-8c14401ec8bd
ms.date: 06/08/2017
ms.openlocfilehash: ec8e46c0445815f2fd7f9d68bf293386b614b6c4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tag-object-publisher"></a>Объект тега (издатель)

Представляет тег или настраиваемое свойство, которое можно создать для фигуры, диапазона фигуры, страницы или публикации. Каждый объект **тега** содержит имя настраиваемого свойства и значения для этого свойства. **Тег** объекты являются участниками семейства **[тегов](tags-object-publisher.md)** . Создайте тег, чтобы иметь возможность избирательно работать с определенными членами семейства сайтов, на основе атрибута, еще не представленные встроенных свойств.
 


## <a name="example"></a>Пример

Метод **[Item](tags-item-method-publisher.md)** коллекции **[теги](tags-object-publisher.md)** возвращает объект **тега** . В этом примере заполняет всех фигур на первой странице active публикации, если первый тег фигуры имеет значение овал.
 

 

```
Sub FormatTaggedShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.Tags.Count > 0 Then 
 If shp.Tags.Item(1).Value = "Oval" Then 
 shp.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 End If 
 End If 
 Next 
 End With 
End Sub
```

Используйте метод **[Add](tags-add-method-publisher.md)** для добавления объекта тега. В этом примере добавляется тег для всех фигур Овал в активной публикации.
 

 



```
Sub TagShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If InStr(1, shp.Name, "Oval") > 0 Then 
 shp.Tags.Add Name:="Oval", Value:="This is an oval shape." 
 End If 
 Next shp 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](tag-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](tag-application-property-publisher.md)|
|[Name](tag-name-property-publisher.md)|
|[Родительский раздел](tag-parent-property-publisher.md)|
|[Значение](tag-value-property-publisher.md)|

