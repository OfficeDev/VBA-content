---
title: "Свойство BorderArt.Application (издатель)"
keywords: vbapb10.chm7667713
f1_keywords: vbapb10.chm7667713
ms.prod: publisher
api_name: Publisher.BorderArt.Application
ms.assetid: ecdd7a8a-9f3b-9cd3-9454-648e0be6f42e
ms.date: 06/08/2017
ms.openlocfilehash: 8b726ca0fc12840ab93f1bf39c7df28157340e0a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartapplication-property-publisher"></a>Свойство BorderArt.Application (издатель)

При использовании без квалификатор объекта, данное свойство возвращает объект **[приложения](application-object-publisher.md)** , который представляет текущего экземпляра Publisher. Используется квалификатор объекта, данное свойство возвращает объект **приложения** , представляющего создателя указанный объект. При использовании с помощью объекта OLE-автоматизации возвращает объект приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приложения**

 переменная _expression_A, представляет собой объект- **Узорные** .


## <a name="example"></a>Пример

В этом примере отображаются сведения о версии и построения для Publisher.


```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 &; .Version &; " build " &; .Build 
End With
```

В этом примере отображается имя приложения, создавшего каждого связанного объекта на странице один активный публикации.




```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Узорные объектов](borderart-object-publisher.md)

