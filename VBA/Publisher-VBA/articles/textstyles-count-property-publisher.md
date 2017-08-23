---
title: "Свойство TextStyles.Count (издатель)"
keywords: vbapb10.chm5898243
f1_keywords: vbapb10.chm5898243
ms.prod: publisher
api_name: Publisher.TextStyles.Count
ms.assetid: c8620d07-d5ad-68f6-67c6-0179da441a4c
ms.date: 06/08/2017
ms.openlocfilehash: 2c187fdb19b5575815f987de9e4848de012f34c0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylescount-property-publisher"></a>Свойство TextStyles.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **TextStyles** .


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


