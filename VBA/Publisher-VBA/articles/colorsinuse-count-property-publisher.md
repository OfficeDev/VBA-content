---
title: "Свойство ColorsInUse.Count (издатель)"
keywords: vbapb10.chm2949122
f1_keywords: vbapb10.chm2949122
ms.prod: publisher
api_name: Publisher.ColorsInUse.Count
ms.assetid: 2f1cdf49-665a-63e9-d221-a1abf756b501
ms.date: 06/08/2017
ms.openlocfilehash: 16650d1f85789658c8a660b16c4984d0a23819da
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorsinusecount-property-publisher"></a>Свойство ColorsInUse.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **ColorsInUse** .


## <a name="example"></a>Пример

В этом примере отображается число страниц в активный документ.


```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " &; _ 
 ActiveDocument.Pages.Count &; " page(s)." 
End Sub
```

В этом примере отображается количество фигур в активном документе.




```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " &; intShapes &; " shape(s)." 
End Sub
```


