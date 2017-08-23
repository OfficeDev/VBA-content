---
title: "Свойство Hyperlink.Shape (издатель)"
keywords: vbapb10.chm4587527
f1_keywords: vbapb10.chm4587527
ms.prod: publisher
api_name: Publisher.Hyperlink.Shape
ms.assetid: afd1dab7-472a-2aa5-f5da-1e2f783b5270
ms.date: 06/08/2017
ms.openlocfilehash: f2834d025b0e9de39bc4df762cb9ae7cf944f967
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinkshape-property-publisher"></a>Свойство Hyperlink.Shape (издатель)

Возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий фигуры, связанной с гиперссылки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Фигура**

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="example"></a>Пример

В этом примере добавляется гиперссылка на первую фигуру на первой странице active публикации и по вертикали зеркальное отражение фигуры. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub FormatHyperlinkShape() 
 With ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 .Address = "http://www.tailspintoys.com/" 
 .Shape.Flip FlipCmd:=msoFlipVertical 
 End With 
End Sub
```


