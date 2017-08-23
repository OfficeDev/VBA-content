---
title: "Объект PageSetup (издатель)"
keywords: vbapb10.chm7012351
f1_keywords: vbapb10.chm7012351
ms.prod: publisher
api_name: Publisher.PageSetup
ms.assetid: 23fe3235-88ea-ac93-6d7d-850298263046
ms.date: 06/08/2017
ms.openlocfilehash: 6ad48bd81658ca3c7177c9ac1a29ef838e073372
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetup-object-publisher"></a>Объект PageSetup (издатель)

Содержит информацию о настройке страницы для страниц в публикации.


## <a name="example"></a>Пример

Свойство **[PageSetup](http://msdn.microsoft.com/library/1dac39f0-2507-a85b-8c71-cd1980022fb3%28Office.15%29.aspx)** используется для возврата объекта **PageSetup** . В следующем примере задается всех страниц в активной публикации 8.5 дюймов ширину и высоту 11 дюймов.


```
Sub SetPageSetupOptions() 
 With ActiveDocument.PageSetup 
 .PageHeight = 11 * 72 
 .PageWidth = 8.5 * 72 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/fe7f0fc3-6449-63b8-21fc-d8ce8f7eb6cc%28Office.15%29.aspx)|
|[AvailablePageSizes](http://msdn.microsoft.com/library/5ad79ee6-6d32-6c46-c02e-a9ab252267cf%28Office.15%29.aspx)|
|[HorizontalGap](http://msdn.microsoft.com/library/e8ee51e0-59b3-8fb6-21f6-87d67a96dd66%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/19fbb72e-bb6e-18e9-28f3-c7e99b071bfb%28Office.15%29.aspx)|
|[PageHeight](http://msdn.microsoft.com/library/1ef153e2-5d13-d896-cd69-2066efa2f8ef%28Office.15%29.aspx)|
|[PageSize](http://msdn.microsoft.com/library/b0605215-5d91-e26e-d3c5-98254cf30044%28Office.15%29.aspx)|
|[PageWidth](http://msdn.microsoft.com/library/547f2881-d9fa-fa5f-1643-ab08084fb423%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/0aebdd7d-6ac6-77c1-1854-edab76ca0b10%28Office.15%29.aspx)|
|[PublicationLayout](http://msdn.microsoft.com/library/6c476789-577d-2088-37dc-bcaed25cd219%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/4eee9b1e-6c76-7831-13bc-25926c3c0f10%28Office.15%29.aspx)|
|[VerticalGap](http://msdn.microsoft.com/library/191d66c4-d168-625a-47b7-028167a98af9%28Office.15%29.aspx)|

