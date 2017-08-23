---
title: "Объект PageSizes (издатель)"
keywords: vbapb10.chm8847359
f1_keywords: vbapb10.chm8847359
ms.prod: publisher
api_name: Publisher.PageSizes
ms.assetid: f31b08cc-2c76-e2d6-d1ae-6dcf2ac5824c
ms.date: 06/08/2017
ms.openlocfilehash: 2da1a8a55e9041454f0ec8cf899eed1cd3d430c8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesizes-object-publisher"></a>Объект PageSizes (издатель)

Представляет коллекцию всех объектов **PageSize** в родительский объект **документа** , где каждый объект **PageSize** представляет размер бумаги в текущем документе Microsoft Publisher.


## <a name="remarks"></a>Заметки

Размеры страниц, представленный в коллекцию **PageSizes** соответствуют значки, отображаемые в разделе **Пустая страница размеры** в диалоговом окне **Параметры страницы** в пользовательском интерфейсе Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать **PageSizes** коллекции для получения все страницы размеры доступно в текущем документе и Печать списка в окне **Интерпретация** .


```
Public Sub PageSizes_Example() 
 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim pubPageSize As Publisher.PageSize 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 For Each pubPageSize In pubPageSizes 
 Debug.Print pubPageSize.Name 
 Next 
 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/bce8ec2c-1a05-1e0b-8d75-7e4dd7084a19%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/10770705-e8b3-903c-bcfa-84ba26dc1478%28Office.15%29.aspx)|
|[Элемент](http://msdn.microsoft.com/library/7fc17907-7e3b-8415-23dc-7f1a64db731c%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/622d2bee-a7b7-6f5f-cb7c-39d69f432b27%28Office.15%29.aspx)|

