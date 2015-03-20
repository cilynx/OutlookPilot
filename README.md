# OutlookPilot #

A productivity plugin for Outlook, inspired by [MailPilot](http://mindsense.co/mailpilot/)

## Overview ##

While working at [AWS](http://www.aws.com), I discovered [MailPilot](http://mindsense.co/mailpilot/) and quickly embraced its model of scheduling emails to be dealt with at some point in the future.  I especially like that the schedule is managed via a standard folder structure as opposed to a local metadata store as this allowed me to interact with the scheduling hierarchy using any mail client from any device.

Now that I'm working at [Azure](http://www.azure.com), having a Mac as my primary machine suddenly became less of an option.  Outlook supports some rudimentary reminder scheduling, but does not enable the granularity that MailPilot does and does not surface the schedule in any way that can be accessed from other clients accessing the account.  Thus, OutlookPilot was born.

### Basic Functions ###

OutlookPilot adds a collection of buttons to the Outlook Ribbon which enable you to schedule emails for processing at a later date.  The _1-9_ buttons schedule the email for 1-9 days out, respectively.  _Today_ is self-explanatory.

![OutlookPilot Screenshot](screenshots/OutlookPilot-0.0.1.0.png)

### The Joy of Whenever ###

The _Whenever_ button looks over your upcoming schedule and schedules the active item for the first workday that you're not busy.  As of today, _workday_ is defined as Monday-Friday and _busy_ is defined as <5 actions already scheduled on that day.  In the future, these could be configurable settings.  _Whenever_ is arguably OutlookPilot's best feature and one that is sadly absent from MailPilot.

(To be fair, MailPilot has a _Set Aside_ feature, but this puts the active message into a folder that you need to actively look at in the future when you have nothing better to do.  How often do you have nothing better to do?  When I was using MailPilot, I found that using _Set Aside_ was no better for me than archiving or deleting the item.)

### Schedule Folder Hierarchy ###

Your schedule is stored in a collection of folders under a top-level _Pilot_ folder.  For example, right now my schedule hierarchy looks like this:

Pilot
- 20150320
- 20150323
- 20150324
- 20150325
- 20150326
- 20150327
- 20150330
- 20150331
- 20150401
- ...
- 20150410
- 20150413
- 20150414

OutlookPilot will automagically create folders as it needs them and remove folders when they're empty.  OutlookPilot does not maintain any external metadata store, so you can create and delete folders manually as well as move around the messages inside them and OutlookPilot will continue to perform exactly as you would expect.

### Weekends and Busy Days ###

OutlookPilot has a smidge of intelligence that will prompt you for confirmation if you try to schedule an item on a weekend or on a day that you're already busy.  It will obviously still let you do whatever you want to do, but it'll be there to remind you not to set yourself up for failure.

### Blocked Days ###

You can manually prevent OutlookPilot from scheduling on a specific day by creating a folder that starts with the date you care about and has anything else after it.  For example "20150401 - OOTO", "20150402: At Training", or even "20150403 " would all be valid _blocked_ folders.  OutlookPilot will pass right over them when autoscheduling similar to how it handles weekends.  _Blocked_ ranges are not currently supported because it's difficult to make elegant when viewing the folders on other clients.

## FAQ ##

### 1. Is OutlookPilot compatible with MailPilot? ###

Not today, no.  I don't like MailPilot's folder naming as it doesn't follow digit significance and thus doesn't sort well across multiple months.  I'm not completely married to my implementation if others prefer otherwise.
