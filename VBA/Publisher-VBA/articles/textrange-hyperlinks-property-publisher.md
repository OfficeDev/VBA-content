---
title: "Свойство TextRange.Hyperlinks (издатель)"
keywords: vbapb10.chm5308485
f1_keywords: vbapb10.chm5308485
ms.prod: publisher
api_name: Publisher.TextRange.Hyperlinks
ms.assetid: 0cf1f043-532c-3ffc-67cf-389adc5ac02f
ms.date: 06/08/2017
ms.openlocfilehash: 7bb0c98f9aefb768a3fb3d7bad7e8a4cfe83be20
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangehyperlinks-property-publisher"></a>Свойство TextRange.Hyperlinks (издатель)

Возвращает коллекцию **[гиперссылки](hyperlinks-object-publisher.md)** , представляющую все гиперссылки в диапазоне указанный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Гиперссылки**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Гиперссылки


## <a name="example"></a>Пример

В следующем примере выполняется поиск всех фигур на странице один из активных публикации, которые имеют текстовые рамки и сообщает, сколько гиперссылки каждой фигуры.


```vb
Dim hypAll As Hyperlinks 
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.HasTextFrame = msoTrue Then 
 Set hypAll = shpLoop.TextFrame.TextRange.Hyperlinks 
 Debug.Print "Shape " &; shpLoop.Name _ 
 &; " has " &; hypAll.Count &; " hyperlinks." 
 End If 
Next shpLoop
```


