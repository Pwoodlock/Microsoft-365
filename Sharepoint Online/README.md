---
tags:
  - sharepoint
  - microsoft
  - powershell
SharePoint: Latest Scripts for Rockfield IT
---





> [! The following Scripts are here]
> - **SharePoint House Keeping Script**
> - **Online SharePoint Site to Online SharePoint Migration**









### ***- SharePoint House Keeping (Interactive)***
	- Versioning enable check.
	- Set two types Versioning
	- Search and Delete older versions after versioning number set
	- TODO: Add Rights to Global Admin to Purge Recycle Bin with "E-Discovery Manger Role / Security Compliance Roles" (You can still do this in the compliance centre portal)







#  - Online SharePoint Sub Site ==> *Online SharePoint Site* to New Tenant

> ***I have coded to a point where it does all the following below. I just didn't get back to it to finish the uploading side of things.  I will post in detail later regarding the different types modules and techniques used for the site to site migration.  This is a Non-Production at the moment!!! (Sorry)


	- Hook into Sharepoint Site address on Source Tenant
	- Download Associated Groups / Members / Users
	- Create csv file and folder directories of Sub Site
	- Download all types of meta data, media files etc...
	- Hook into Target Site and download Associated Groups / Members / Users 
	- Merge into source csv and check for conflicts 
	- That's it for now
	- [ ] Upload & Verify full contents to MS365 Tenant



