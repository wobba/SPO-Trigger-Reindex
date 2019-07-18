# SPO Trigger Re-index scripts by [@mikaelsvenson]
## Scripts
- reindex-tenant.ps1 - Script to mark SharePoint online content to be picked up at the next crawl.

## Re-indexing of user profiles
See [How to trigger re-indexing of user profiles in SharePoint On-line] for an explanation of user profile indexing

This script will iterate all user profiles in SPO to force a trigger of re-indexing of the profiles on the next crawl.

Monitoring has shown that indexing of user profiles happens
at approximately a **4 hour interval** (per December 10th 2014). I have seen as low as **2h** (night time) and as high as **8h** (daytime).

The script is executes like this:

    .\reindex-users.ps1 -url https://techmikael-admin.sharepoint.com
0
You can also add the switch *-changeProperty* to choose if you want to use *SPS-Birthday* or *Department* (now default) as your change property.

[How to trigger re-indexing of user profiles in SharePoint On-line]:http://techmikael.blogspot.com/2014/12/how-to-trigger-re-indexing-of-user.html
