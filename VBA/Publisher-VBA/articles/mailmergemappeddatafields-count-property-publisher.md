---
title: "Свойство MailMergeMappedDataFields.Count (издатель)"
keywords: vbapb10.chm6488067
f1_keywords: vbapb10.chm6488067
ms.prod: publisher
api_name: Publisher.MailMergeMappedDataFields.Count
ms.assetid: 45bb34e6-3b6f-2daa-d782-2bbd02b1e7b4
ms.date: 06/08/2017
ms.openlocfilehash: d4a6599a724f761161614564ac6ac297aa7448bb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergemappeddatafieldscount-property-publisher"></a>Свойство MailMergeMappedDataFields.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **MailMergeMappedDataFields** .


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


