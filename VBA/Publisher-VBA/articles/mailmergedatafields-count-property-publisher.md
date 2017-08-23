---
title: "Свойство MailMergeDataFields.Count (издатель)"
keywords: vbapb10.chm6356993
f1_keywords: vbapb10.chm6356993
ms.prod: publisher
api_name: Publisher.MailMergeDataFields.Count
ms.assetid: f46da7b1-acd8-f2d2-a6aa-71cc3c8eca99
ms.date: 06/08/2017
ms.openlocfilehash: 938a0ce7107a45b97b71f975ae42a15f804b0071
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldscount-property-publisher"></a>Свойство MailMergeDataFields.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **MailMergeDataFields** .


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


