---
title: "Свойство Document.ViewTwoPageSpread (издатель)"
keywords: vbapb10.chm196665
f1_keywords: vbapb10.chm196665
ms.prod: publisher
api_name: Publisher.Document.ViewTwoPageSpread
ms.assetid: b5e851ff-d5fc-a98d-02b3-7e14c1b957dc
ms.date: 06/08/2017
ms.openlocfilehash: beff1d85aa75606a3b2ffa700ea0bfcafab46625
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentviewtwopagespread-property-publisher"></a>Свойство Document.ViewTwoPageSpread (издатель)

Возвращает **значение True** , если указанный публикации следует рассматривать как двух страницах. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ViewTwoPageSpread**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере открывается окно сообщения и отображает, если текущей публикации должны быть отображены на странице два включен режим.


```vb
Sub ViewTwoPage() 
 
 MsgBox "View Two Page Spread = " &; _ 
 Application.ActiveDocument.ViewTwoPageSpread 
 
End Sub
```


