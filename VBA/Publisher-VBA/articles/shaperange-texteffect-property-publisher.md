---
title: "Свойство ShapeRange.TextEffect (издатель)"
keywords: vbapb10.chm2293833
f1_keywords: vbapb10.chm2293833
ms.prod: publisher
api_name: Publisher.ShapeRange.TextEffect
ms.assetid: 7bc822f2-4754-685d-fdd3-7479b5a3ac52
ms.date: 06/08/2017
ms.openlocfilehash: c50867c3f7a6f5737d5249f83fb77a0e9bd88a98
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangetexteffect-property-publisher"></a>Свойство ShapeRange.TextEffect (издатель)

Возвращает объект **[TextEffectFormat](texteffectformat-object-publisher.md)** , который представляет свойства объекта WordArt форматирования текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextEffect**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере добавляет объект WordArt active публикации и форматы и вставки дополнительных в нее.


```vb
Sub AddFormatNewWordArt() 
 With ActiveDocument.Pages(1).Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Snap ITC", FontSize:=30, FontBold:=msoTrue, _ 
 FontItalic:=msoFalse, Left:=150, Top:=130) 
 .Rotation = 90 
 With .TextEffect 
 .RotatedChars = msoTrue 
 .Text = "This is a " &; .Text 
 End With 
 .Width = 250 
 End With 
End Sub
```


