# OutlookPilot #

A productivity plugin for Outlook, inspired by [MailPilot](http://mindsense.co/mailpilot/)

## Overview ##

While working with [AWS](http://www.aws.com), I discovered [MailPilot](http://mindsense.co/mailpilot/) and quickly embraced its model of scheduling emails to be dealt with at some point in the future.  I especially like that the schedule is managed via a standard folder structure as opposed to a local metadata store as this allowed me to interact with the scheduling hierarchy using any mail client from any device.

Now that I'm working with [Azure](http://www.azure.com), having a Mac as my primary machine suddenly became less of an option.  Outlook supports some rudimentary reminder scheduling, but does not enable the granularity that MailPilot does and does not surface the schedule in any way that can be accessed from other clients accessing the account.  Thus, OutlookPilot was born.

## FAQ ##

### 1. Is OutlookPilot compatible with MailPilot? ###

Not today, no.  I don't like MailPilot's folder naming as it doesn't follow digit significance and thus doesn't sort well across multiple months.  I'm not completely married to my implementation if others prefer otherwise.

### 2. What does OutlookPilot do? ###

OutlookPilot adds a collection of buttons to the Outlook Ribbon which enable you to schedule emails for processing at a later date.  The _1-9_ buttons schedule the email for 1-9 days out, respectively.  _Today_ is self-explanatory.

![OutlookPilot Screenshot](/cilynx/OutlookPilot/master/screenshots/OutlookPilot-0.0.1.0.png)

The _Whenever_ button looks over your upcoming schedule and schedules the active item for the first workday that you're not busy.  As of today, _workday_ is defined as Monday-Friday and _busy_ is defined as <5 actions already scheduled on that day.  It the future, these could be configurable settings.
