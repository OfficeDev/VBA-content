---
title: "Свойство OLEFormat.Object (издатель)"
keywords: vbapb10.chm4456451
f1_keywords: vbapb10.chm4456451
ms.prod: publisher
api_name: Publisher.OLEFormat.Object
ms.assetid: c6bc20e4-4578-7aa1-8cd8-8315b76b28c9
ms.date: 06/08/2017
ms.openlocfilehash: ac2a268d6843e30dcd41e53b5615a05917aada01
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="oleformatobject-property-publisher"></a>Свойство OLEFormat.Object (издатель)

Возвращает **объект** , представляющий интерфейс верхнего уровня на указанный объект OLE. Это свойство позволяет получить доступ к свойствам и методам элемента управления ActiveX или приложения, в котором был создан объект OLE. Объект OLE должна поддерживать OLE-автоматизации для этого свойства для работы. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Объект**

 переменная _expression_A, представляющий объект **OLEFormat** .


### <a name="return-value"></a>Возвращаемое значение

Object


## <a name="example"></a>Пример

Этот пример задает значение первой фигуры в активной публикации. Для обеспечения работы примера этой первой фигуры должен быть элемент управления ActiveX (например, флажок или переключатель).


```vb
Dim myObj As Object 
 
With ActiveDocument.Pages(1).Shapes(1).OLEFormat 
 .Activate 
 Set myObj = .Object 
End With 
 
myObj.Value = True
```


