# Exchange Server Components Checker

Here’s the latest addition in Exchange 2013/2016 server quick tools. 
Sometimes, for maintenance or issues detected by Exchange Managed Availability, some components 
like Client Access, Autodiscover, Mail flow, etc... can be down on Exchange servers, which "exclude"
these servers from the pool of Exchange servers.
Here's a tool that helps you to check the state of every Exchange Server’s components and optionally start these if they are inactive.

These “Server Components” are new in Exchange 2013/2016, these are part of the Exchange Managed Availability features, and are there to ease server maintenance: inactive components are not targeted by the Client Access Services so that you can perform maintenance without having to affect the other Exchange services.



-	Choose your server version, Exchange 2013 or Exchange 2016, to check the components from
-	**[  ] CheckOnly** checkbox will just check and show each component from each of your organization’s servers: if checked, the tool will not attempt to start the inactive components
-	**[  ] HybridServer** checkbox will check two additional components that are used on Hybrid Exchange – O365 scenarios
-	**[  ] Show Inactive Only** : this checkbox will filter out all the Active components, so that you can check right away which servers and which components have Inactive services. You can check or uncheck it anytime, before or after the Server Components collection
-	Once you are happy with the chosen options, click the **[Run]** button to start gathering the information (at the bottom is the status bar along with a status label)

![image](https://user-images.githubusercontent.com/33433229/198128768-50ec47d0-0eac-4f86-8236-ad5c6d6761ef.png)
 

NOTE: once the components are gathered and show up, you can select all (click on one of the result cell, then CTRL+ A followed by CTRL + C) and paste it in a report, in Excel, or anywhere else:
 
![image](https://user-images.githubusercontent.com/33433229/198129402-1dd9a1f9-e983-437d-898e-460069d079fa.png)

And finally, you're able to start the Exchange Server Components with the requester of your choice (Maintenance, SideLine, Functional, Deployment,...) - just uncheck the **[ ] CheckOnly** check box, choose your requester, and click **[Run]** - **NOTE**: if you see a component is Inactive with a specific Requester, you can only bring it "Active" with the same Requester that brought it "Inactive" (the one that is listed next to the "Inactive" state of the component).

![image](https://user-images.githubusercontent.com/33433229/198129910-98513430-8071-4978-97c4-745befd30177.png)

More information on these Exchange Server Components:
https://blogs.technet.microsoft.com/exchange/2013/09/26/server-component-states-in-exchange-2013/
