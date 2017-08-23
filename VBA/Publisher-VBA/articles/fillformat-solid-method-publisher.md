---
title: "Метод FillFormat.Solid (издатель)"
keywords: vbapb10.chm2359317
f1_keywords: vbapb10.chm2359317
ms.prod: publisher
api_name: Publisher.FillFormat.Solid
ms.assetid: e34f6bc0-308b-4f86-5ce9-87e05c4a2089
ms.date: 06/08/2017
ms.openlocfilehash: a75e86303a4a901fd317eba8bf69ac9b75eb40ae
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatsolid-method-publisher"></a>Метод FillFormat.Solid (издатель)

Устанавливает указанный заливки единый цвет. Используйте этот метод для преобразования градиент, текстурой, узором или фона заливки обратно в сплошной заливке.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сплошной**

 переменная _expression_A, представляет собой объект- **FillFormat** .


## <a name="example"></a>Пример

В этом примере преобразует все заливки на первой странице активная публикация универсальный красной заливки.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 .Solid 
 .ForeColor.RGB = RGB(255, 0, 0) 
 End With 
Next shpLoop 

```


