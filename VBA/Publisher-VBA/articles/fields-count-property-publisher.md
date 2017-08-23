---
title: "Свойство Fields.Count (издатель)"
keywords: vbapb10.chm6029315
f1_keywords: vbapb10.chm6029315
ms.prod: publisher
api_name: Publisher.Fields.Count
ms.assetid: a8a6b0d4-b029-0b45-6d76-6fb237c31c97
ms.date: 06/08/2017
ms.openlocfilehash: 72c05c8f5398c2e214e5fbc81bd49922c4e72144
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldscount-property-publisher"></a>Свойство Fields.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляющий объект **поля** .


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


