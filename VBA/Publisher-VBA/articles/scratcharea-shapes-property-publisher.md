---
title: "Свойство ScratchArea.Shapes (издатель)"
keywords: vbapb10.chm1179651
f1_keywords: vbapb10.chm1179651
ms.prod: publisher
api_name: Publisher.ScratchArea.Shapes
ms.assetid: 0d867fec-42f4-fd61-c6c3-745be955e5d2
ms.date: 06/08/2017
ms.openlocfilehash: 7e7e1201e2d3277621c2687cafa555feaee0ce5a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="scratchareashapes-property-publisher"></a>Свойство ScratchArea.Shapes (издатель)

Возвращает коллекцию **[фигур](shapes-object-publisher.md)** , представляющий все объекты **фигур** в указанной публикации. Эта коллекция может содержать документы, фигур, рисунков, OLE объекты, ActiveX элементов управления, текстовые объекты и выноски.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Фигур**

 переменная _expression_A, представляет собой объект- **ScratchArea** .


## <a name="remarks"></a>Заметки

Сведения о возврате один элемент коллекции видеть **возврата объекта из коллекции**.


## <a name="example"></a>Пример

В этом примере добавляется прямоугольник для первой страницы в активной публикации.


```vb
Sub AddNewRectangle() 
 ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeRectangle, _ 
 Left:=5, Top:=25, Width:=100, Height:=50 
End Sub
```

В этом примере задается текстуры заливки для всех фигур в активной публикации. В этом примере предполагается, что имеется по крайней мере один фигуры в активной публикации.




```vb
Sub SetNewTextureForAllShapes() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 shp.Fill.PresetTextured PresetTexture:=msoTextureOak 
 Next shp 
End Sub
```

В этом примере добавляется тень для первой фигуры в активной публикации. В этом примере предполагается, что имеется по крайней мере один фигуры в активной публикации.




```vb
Sub SetShadowForFirstShape() 
 ActiveDocument.Pages(1).Shapes(1).Shadow.Type = msoShadow6 
End Sub
```

В этом примере отображается количество всех фигур на первой странице active публикации. В этом примере предполагается, что имеется по крайней мере один фигуры в активной публикации.




```vb
Sub CountShapesOnFirstPage() 
 MsgBox "You have " &; ActiveDocument.Pages(1) _ 
 .Shapes.Count &; " shapes on the first page." 
End Sub
```


