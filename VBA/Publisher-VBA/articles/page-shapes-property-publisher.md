---
title: "Свойство Page.Shapes (издатель)"
keywords: vbapb10.chm393219
f1_keywords: vbapb10.chm393219
ms.prod: publisher
api_name: Publisher.Page.Shapes
ms.assetid: 4e48d4cf-d7b6-9099-ddee-46a79e7eb7bf
ms.date: 06/08/2017
ms.openlocfilehash: 6317cb016e3c4f9d7e9899bdd46ea557333e9096
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageshapes-property-publisher"></a>Свойство Page.Shapes (издатель)

Возвращает коллекцию **[фигур](shapes-object-publisher.md)** , представляющий все объекты **фигур** в указанной публикации. Эта коллекция может содержать документы, фигур, рисунков, OLE объекты, ActiveX элементов управления, текстовые объекты и выноски.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Фигур**

 переменная _expression_A, представляющий объект **Page** .


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


