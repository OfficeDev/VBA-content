---
title: IBlogExtensibility Members (Office)
ms.prod: office
ms.assetid: 55f27978-9b18-f9a5-c276-298b2539ec3c
ms.date: 06/08/2017
---


# IBlogExtensibility Members (Office)
An object that provides the ability to manipulate blog entries.

An object that provides the ability to manipulate blog entries.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BlogProviderProperties](iblogextensibility-blogproviderproperties-method-office.md)|Contains information about the provider.|
|[GetCategories](iblogextensibility-getcategories-method-office.md)|This method returns the list of blog categories for an account so Microsoft Word can populate the categories dropdown list.|
|[GetRecentPosts](iblogextensibility-getrecentposts-method-office.md)|Returns the list of the user's last fifteen blog posts that Microsoft Word then displays in the  **Open Existing Post** dialog. This method does not actually return the blog post contents.|
|[GetUserBlogs](iblogextensibility-getuserblogs-method-office.md)|Returns the list and details of user blogs associated with the specified account.|
|[Open](iblogextensibility-open-method-office.md)|Opens the blog specified by the blog ID. It is called by the  **Open Existing Post** dialog based on the item selected by the user.|
|[PublishPost](iblogextensibility-publishpost-method-office.md)|Hands off the current post so it can be published by the provider.|
|[RepublishPost](iblogextensibility-republishpost-method-office.md)|Hands off the current post so it can be republished by the provider.|
|[SetupBlogAccount](iblogextensibility-setupblogaccount-method-office.md)|Called from the  **Choose Account** dialog when the provider's name is chosen in the **Blog Host** dropdown or when the user requests to change a provider's account in the **Blog Accounts** dialog box.|

