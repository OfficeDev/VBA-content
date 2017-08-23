---
title: "Свойство TextRange.Text (издатель)"
keywords: vbapb10.chm5308416
f1_keywords: vbapb10.chm5308416
ms.prod: publisher
api_name: Publisher.TextRange.Text
ms.assetid: 13584812-307a-c32b-ca8f-27869728b64e
ms.date: 06/08/2017
ms.openlocfilehash: 7ddc6512093072f2c9178aa607f4bbb83c643573
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangetext-property-publisher"></a>Свойство TextRange.Text (издатель)

Возвращает или задает **строку** , представляющую текст в диапазон текста или WordArt фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Текст**

 переменная _expression_A, представляющий объект **TextRange** .


## <a name="example"></a>Пример

В следующем примере добавляет прямоугольник active публикации и добавляет текст.


```vb
Sub AddTextToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
 Left:=72, Top:=72, Width:=250, Height:=140) 
 .TextFrame.TextRange.Text = "Here is some test text" 
 End With 
End Sub
```


