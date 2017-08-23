---
title: "Свойство Shape.Hyperlink (издатель)"
keywords: vbapb10.chm2228323
f1_keywords: vbapb10.chm2228323
ms.prod: publisher
api_name: Publisher.Shape.Hyperlink
ms.assetid: 0990ab32-b4a3-6c89-cb9f-8f8c64ef804f
ms.date: 06/08/2017
ms.openlocfilehash: ec72d040b1a781aa130801d6baad05d48f6f258d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapehyperlink-property-publisher"></a>Свойство Shape.Hyperlink (издатель)

Возвращает объект **[гиперссылки](hyperlink-object-publisher.md)** , представляющий гиперссылки, связанной с указанной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Гиперссылки**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере задается фигуры одно по одному в активной публикации для перехода к указанного веб-сайта при щелчке фигуры.


```vb
Dim hypTemp As Hyperlink 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
hypTemp.Address = "http://www.tailspintoys.com/"
```


