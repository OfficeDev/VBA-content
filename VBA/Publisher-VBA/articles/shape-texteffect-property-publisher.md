---
title: "Свойство Shape.TextEffect (издатель)"
keywords: vbapb10.chm2228297
f1_keywords: vbapb10.chm2228297
ms.prod: publisher
api_name: Publisher.Shape.TextEffect
ms.assetid: 187b55f8-9593-6a00-61e6-dbcf5c56b987
ms.date: 06/08/2017
ms.openlocfilehash: 148f3d198ea09b0ec280758b2db80b5346d4f347
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetexteffect-property-publisher"></a>Свойство Shape.TextEffect (издатель)

Возвращает объект **[TextEffectFormat](texteffectformat-object-publisher.md)** , который представляет свойства объекта WordArt форматирования текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextEffect**

 переменная _expression_A, представляющий объект **фигуры** .


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


