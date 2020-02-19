# MigrateToOffice365
This script allows you to automate your migration from Office 2016, to Office365.  The script will download the necessary installation files from Microsoft's Content Delivery Network (CDN) and go through a configuration procedure to install Office365.  

### FILES YOU WILL NEED ###

1.) Configure the XML files for Office365, Project, and Visio: 
- Instructions (https://docs.microsoft.com/en-us/deployoffice/overview-of-the-office-2016-deployment-tool)
- Office Customization Tool (https://config.office.com/deploymentsettings)
- Get the Office Deployment Tool (https://www.microsoft.com/en-us/download/details.aspx?id=49117)

2.) Create the Deployment Folder
- Name the folder 'OfficeLocal'

3.) Put in the XMLs and Office Deployment Tool (setup.exe) in this folder:
- Office365: Office365ProPlus.xml
- Office365 w/Visio: Office365VisioProPlus.xml
- Office365 w/Project: Office365ProjectProPlus.xml
- Office365 w/Project and Visio: Office365VisioProjectProPlus.xml
