---
title: "Свойство Shape.WebCommandButton (издатель)"
keywords: vbapb10.chm2228340
f1_keywords: vbapb10.chm2228340
ms.prod: publisher
api_name: Publisher.Shape.WebCommandButton
ms.assetid: c20b937b-6f53-fdc1-830a-4044831c351a
ms.date: 06/08/2017
ms.openlocfilehash: 3a7fbf9645a2fe14bb94ebd5d5a674cec51bf004
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapewebcommandbutton-property-publisher"></a>Свойство Shape.WebCommandButton (издатель)

Возвращает объект **[WebCommandButton](webcommandbutton-object-publisher.md)** , связанный с указанным фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebCommandButton**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

WebCommandButton


## <a name="example"></a>Пример

В этом примере создается кнопки Отправить форму Web и задает путь и имя скрипта для запуска при нажатии кнопки.


```vb
Dim shpNew As Shape 
Dim wcbTemp As WebCommandButton 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36) 
 
Set wcbTemp = shpNew.WebCommandButton 
 
With wcbTemp 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" _ 
 &; "scripts/ispscript.cgi" 
End With
```


