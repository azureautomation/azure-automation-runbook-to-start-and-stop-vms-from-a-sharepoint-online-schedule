Azure Automation Runbook to start and stop VMs from a SharePoint Online Schedule
================================================================================

            

Please read this [blog post](https://blogs.msdn.microsoft.com/jrt/2017/05/15/using-azure-automation-and-sharepoint-online-lists-to-schedule-vm-start-up-and-shutdown) for full instructions on the usage of this script.


Process Flowchart --> AutoShutdown Flowchart.pdf


Extract: 
To make the most efficient use of Public cloud you need to ensure your servers are only running when they need to be.
Using SharePoint online as the source of virtual machine start-up and shutdown schedules is an option that allows for an easy, consolidated view of all your VMs and their schedules plus allows for the granular allocation of permissions for your staff, without
 having to grant them any Azure fabric access.


To integrate SharePoint Online with Azure Automation (and Powershell scripts), I've used a fantastic module written by [Tao Yang](https://www.powershellgallery.com/profiles/tao.yang/), SharepointSDK.
The details can be found [here](https://www.powershellgallery.com/packages/SharePointSDK/2.1.5).
So the first thing to do is import this module to your Azure Automation Modules by searching the Gallery for 'SharePointSDK' and/or install it for Powershell using Install-Module SharepointSDK.


 


 

 

 


        
    
TechNet gallery is retiring! This script was migrated from TechNet script center to GitHub by Microsoft Azure Automation product group. All the Script Center fields like Rating, RatingCount and DownloadCount have been carried over to Github as-is for the migrated scripts only. Note : The Script Center fields will not be applicable for the new repositories created in Github & hence those fields will not show up for new Github repositories.
