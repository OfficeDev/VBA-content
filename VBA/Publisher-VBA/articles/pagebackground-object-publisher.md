---
title: "Объект PageBackground (издатель)"
keywords: vbapb10.chm8191999
f1_keywords: vbapb10.chm8191999
ms.prod: publisher
api_name: Publisher.PageBackground
ms.assetid: 647f5a84-0971-2f69-d281-c9ab402968a4
ms.date: 06/08/2017
ms.openlocfilehash: 9dabddae3222af0edf40e1f20cb70660ef342329
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagebackground-object-publisher"></a>Объект PageBackground (издатель)

Представляет фона страницы.
 


## <a name="example"></a>Пример

Свойство **фона** объекта **Page** для получения объекта **PageBackground** . В следующем примере создается объект **PageBackground** и задает фон первой страницы активных документов.
 

 

```
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
 
```

Использование **PageBackground.Exists** для определения, существует ли фон для указанного объекта **страницы** . Следующий пример основан на предыдущем примере. Сначала **PageBackground** объект создается и задать фон для первой страницы активных документов. Тест, выполняется проверка, если фон для страницы уже существует. В противном случае выберите один создается путем вызова метода **создания** объекта **PageBackground** .
 

 



```
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
If objPageBackground.Exists = False Then 
 objPageBackground.Create 
End If 
 
```

Используйте **PageBackground.Fill** для возврата объекта **FillFormat** . Следующий пример основан на предыдущем примере. Сначала **PageBackground** объект создается и задать фон для первой страницы активных документов. Тест, выполняется проверка, если фон для страницы уже существует. В противном случае выберите один создается путем вызова метода **создания** объекта **PageBackground** . С помощью свойства **заполнения** объекта **PageBackground** возвращается объект **FillFormat** . Установите несколько доступных свойств объекта **FillFormat** .
 

 



```
Dim objPageBackground As PageBackground 
Dim objFillFormat As FillFormat 
 
Set objPageBackground = ActiveDocument.Pages(1).Background 
If objPageBackground.Exists = False Then 
 objPageBackground.Create 
End If 
 
Set objFillFormat = objPageBackground.Fill 
With objFillFormat 
 .BackColor.RGB = RGB(Red:=0, GReen:=155, Blue:=99) 
 .ForeColor.RGB = RGB(Red:=155, GReen:=234, Blue:=0) 
 .TwoColorGradient msoGradientDiagonalDown, 4 
End With 
 
```

Удаление фона для указанного страницы с помощью **PageBackground.Delete** . В следующем примере удаляется фон для первой страницы в активном документе. (В следующем примере предполагается что указанной странице для существующего фона. Ошибка выполнения возникает, если страница не содержит фона.)
 

 



```
ActiveDocument.Pages(1).Background.Delete
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Create](pagebackground-create-method-publisher.md)|
|[Delete](pagebackground-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](pagebackground-application-property-publisher.md)|
|[Существует](pagebackground-exists-property-publisher.md)|
|[Заполните поля](pagebackground-fill-property-publisher.md)|
|[Родительский раздел](pagebackground-parent-property-publisher.md)|

