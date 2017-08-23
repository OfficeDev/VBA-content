---
title: "Свойство TextRange.InlineShapes (издатель)"
keywords: vbapb10.chm5308498
f1_keywords: vbapb10.chm5308498
ms.prod: publisher
api_name: Publisher.TextRange.InlineShapes
ms.assetid: ffe2d8f2-e1d7-44ea-00fd-3c6523c9fe44
ms.date: 06/08/2017
ms.openlocfilehash: 07b43f3679ae1637702d7a22dfb159c867cd5644
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinlineshapes-property-publisher"></a>Свойство TextRange.InlineShapes (издатель)

Возвращает коллекцию **[InlineShapes](inlineshapes-object-publisher.md)** , который представляет встроенных фигур, содержащихся в диапазон текста. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InlineShapes**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

InlineShapes


## <a name="remarks"></a>Заметки

С помощью **TextFrame.Story.TextRange.InlineShapes** возвращает всех встроенных фигур в рамке, включая те, которые находятся в переполнения. С помощью **TextFrame.TextRange.InlineShapes** возвращает только видимые встроенных фигур в фрагмент текста, а не указанные в переполнения.


## <a name="example"></a>Пример

Следующий пример находит первую фигуру (текстовое поле) на странице один из активных публикации. Свойство **InlineShapes** затем используется для определения, существует ли фигуры, встроенного в текстовом поле. Если обнаружены какие-либо, каждой фигуры встроенного отражается по вертикали и установленное красный цвет переднего плана.

Обратите внимание на то, с помощью **TextFrame.Story.TextRange.InlineShapes**встроенного фигуры, которые находятся в переполнения также обнаружения.




```vb
Dim theShape As Shape 
Dim i As Integer 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.Story.TextRange 
 If .InlineShapes.Count > 0 Then 
 For i = 1 To .InlineShapes.Count 
 .InlineShapes(i).Flip (msoFlipVertical) 
 .InlineShapes(i).Fill.ForeColor.RGB = vbRed 
 Next 
 End If 
End With
```


