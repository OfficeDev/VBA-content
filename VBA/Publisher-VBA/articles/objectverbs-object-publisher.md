---
title: "Объект ObjectVerbs (издатель)"
keywords: vbapb10.chm4587519
f1_keywords: vbapb10.chm4587519
ms.prod: publisher
api_name: Publisher.ObjectVerbs
ms.assetid: e04cf7db-ee56-7d95-9f5c-7ecee1844866
ms.date: 06/08/2017
ms.openlocfilehash: 53f24fad17f735fb7b6ba484d12201fd083ba6a5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="objectverbs-object-publisher"></a>Объект ObjectVerbs (издатель)

Представляет коллекцию команды OLE для указанного объекта OLE. Команды OLE являются операции, поддерживаемые объекты OLE. Часто используемые команды OLE являются воспроизвести и изменить.
 


## <a name="example"></a>Пример

Свойство **[ObjectVerbs](oleformat-objectverbs-property-publisher.md)** возвращает объект **ObjectVerbs** . В следующем примере отображаются все доступные команды для объекта OLE, содержащихся в первую фигуру на первой странице в активной публикации. В данном примере для работы указанного фигуры должен содержать объекта OLE.
 

 

```
Sub GetVerbs() 
 Dim intCount As Integer 
 
 With ActiveDocument.Pages(1).Shapes(1).OLEFormat 
 For intCount = 1 To .ObjectVerbs.Count 
 MsgBox .ObjectVerbs(intCount) 
 Next 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](objectverbs-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](objectverbs-application-property-publisher.md)|
|[Count](objectverbs-count-property-publisher.md)|
|[Родительский раздел](objectverbs-parent-property-publisher.md)|

