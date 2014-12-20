# SPO Trigger Re-index scripts by [@mikaelsvenson]
- reindex-tenant.ps1 - Script to mark SharePoint online content to be picked up at the next crawl.

- reindex-users.ps1 - Script to mark User Profiles to be picked up at the next crawl    


## Re-indexing of SPO content
This script which will iterate all site collections, sites and sub-sites on your tenant to force a trigger of re-indexing / full crawl of all items on the next crawl.

In SharePoint Online you cannot force a full crawl via the admin user interface and thus have to iterate all sites and set a property in the property bag to accomplish the same task.

This is needed for several search schema modifications, for example when you map crawled properties to managed properties.

The script also has a function to enable population of the ManagedProperties managed property for use with the SharePoint 2013 Search Query Tool v2.

The script is executed like this:

    .\reindex-tenant.ps1 -url https://tenant-admin.sharepoint.com -username someuser -password somepassword

> Be sure to check that **$csomPath** is pointing to the right folder for your CSOM library install

You may also add the -enableAllManagedProperties parameter which by default is set to $false. This is used in conjenction with the feature mentioned above with the
ManagedProperties property.

> For an explanation of the ManagedProperties property read [http://techmikael.blogspot.com/2014/03/debugging-managed-properties-using.html]

More info: [How to trigger a full re-index in SharePoint On-line].

## Re-indexing of user profiles
See [How to trigger re-indexing of user profiles in SharePoint On-line] for an explanation of user profile indexing

**Requires [CSOM package] from September 3rd 2014 or newer**
**DO NOT run from a SharePoint Online Management Powershell. Use a regular Powershell to avoid CSOM DLL conflicts**

This script will iterate all user profiles in SPO to force a trigger of re-indexing of the profiles on the next crawl.

Monitoring has shown that indexing of user profiles happens
at approximately a **4 hour interval** (per December 10th 2014). I have seen as low as **2h** (night time) and as high as **8h** (daytime).

The script is executes like this:

    .\reindex-users.ps1 -url https://techmikael-admin.sharepoint.com -username <yourusernameORprompt> -password <yourpwORprompt>

[How to trigger a full re-index in SharePoint On-line]:http://techmikael.blogspot.com/2014/02/how-to-trigger-full-re-index-in.html
[CSOM package]:http://aka.ms/spocsom
[@mikaelsvenson]:https://twitter.com/mikaelsvenson
[http://techmikael.blogspot.com/2014/03/debugging-managed-properties-using.html]:http://techmikael.blogspot.com/2014/03/debugging-managed-properties-using.html
[How to trigger re-indexing of user profiles in SharePoint On-line]:http://techmikael.blogspot.com/2014/12/how-to-trigger-re-indexing-of-user.html
