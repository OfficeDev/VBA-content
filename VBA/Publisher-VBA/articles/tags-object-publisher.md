---
title: "Теги Object (издатель)"
keywords: vbapb10.chm4718591
f1_keywords: vbapb10.chm4718591
ms.prod: publisher
api_name: Publisher.Tags
ms.assetid: 76cccc1e-4623-af8b-f0f8-e6cc245b94fd
ms.date: 06/08/2017
ms.openlocfilehash: 9f875ffc7e4f5a8e419f47e823306ef875b1db21
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tags-object-publisher"></a>Теги Object (издатель)

Коллекция объектов **тег** , представляющий теги или настраиваемых свойств, применяемых к фигуры, диапазона фигуры, страницы или публикации.
 


## <a name="example"></a>Пример

Используйте свойство **[теги](shape-tags-property-publisher.md)** для доступа к коллекции **тегов** . Используйте метод **[Add](tags-add-method-publisher.md)** коллекции **теги** добавьте объект **тег** фигуры, диапазона фигуры, страницы или публикации. В этом примере добавляется тег для каждой фигуры овала на первой странице active публикации.
 

 

```
Sub AddNewTag() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If InStr(1, shp.Name, "Oval") > 0 Then 
 shp.Tags.Add Name:="Shape", Value:="Oval" 
 End If 
 Next shp 
 End With 
End Sub
```

Используйте свойство **[Count](tags-count-property-publisher.md)** для определения, если фигуры, диапазона фигуры, страницы или публикации содержит один или несколько объектов **тег** . В этом примере заполняет всех фигур на первой странице active публикации, если первый тег фигуры имеет значение овал.
 

 



```
Sub FormatTaggedShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.Tags.Count > 0 Then 
 If shp.Tags(1).Value = "Oval" Then 
 shp.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 End If 
 End If 
 Next shp 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](tags-add-method-publisher.md)|
|[Элемент](tags-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](tags-application-property-publisher.md)|
|[Count](tags-count-property-publisher.md)|
|[Родительский раздел](tags-parent-property-publisher.md)|

