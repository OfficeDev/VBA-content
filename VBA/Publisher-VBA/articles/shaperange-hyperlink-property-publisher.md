---
title: "Свойство ShapeRange.Hyperlink (издатель)"
keywords: vbapb10.chm2293859
f1_keywords: vbapb10.chm2293859
ms.prod: publisher
api_name: Publisher.ShapeRange.Hyperlink
ms.assetid: 34ec968c-af66-7629-066f-80c8e1b40e84
ms.date: 06/08/2017
ms.openlocfilehash: ef6b2730fc46da768d6f73f166e8d0d61f399566
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangehyperlink-property-publisher"></a>Свойство ShapeRange.Hyperlink (издатель)

Возвращает объект **[гиперссылки](hyperlink-object-publisher.md)** , представляющий гиперссылки, связанной с указанной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Гиперссылки**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере задается фигуры одно по одному в активной публикации для перехода к указанного веб-сайта при щелчке фигуры.


```vb
Dim hypTemp As Hyperlink 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
hypTemp.Address = "http://www.tailspintoys.com/"
```


