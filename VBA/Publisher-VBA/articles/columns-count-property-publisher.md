---
title: "Свойство Columns.Count (издатель)"
keywords: vbapb10.chm5046274
f1_keywords: vbapb10.chm5046274
ms.prod: publisher
api_name: Publisher.Columns.Count
ms.assetid: 2f7fdb6a-6cd0-2ede-bd34-6954ef23c1a0
ms.date: 06/08/2017
ms.openlocfilehash: 337fca83be37404c4b668eec60b67779c09b176d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="columnscount-property-publisher"></a>Свойство Columns.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **столбцов** .


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


