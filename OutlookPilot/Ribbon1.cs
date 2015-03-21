using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookPilot
{
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Diagnostics; // For Debug.writeLine
    using System.Windows.Forms; // For MessageBox
    
    public partial class Ribbon1
    {
       
        /* TODO:
         * Calendar-based Busy Calculations
         */

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) 
        {
            removeEmptyFolders();
        }

        private void removeEmptyFolders()
        {
            /* TODO: Make this configurable */
            Outlook.Folder pilotFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Pilot"] as Outlook.Folder;
            
            /* Remove any empty Pilot folders */
            Outlook.Folders childFolders = pilotFolder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    if (childFolder.Items.Count == 0)
                    {
                        if (childFolder.Name.Length == 8) // Don't delete "blocked" folders even if they're empty
                        { 
                            Debug.WriteLine(childFolder.Name + " is empty and we are deleting it");
                            childFolder.Delete(); 
                        } 
                    }
                }
            }
        }

        private string folderStatus(string folderName)
        {
            /* TODO: Make this configurable */
            Outlook.Folder pilotFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Pilot"] as Outlook.Folder;

            /* Look through all Pilot folders and see if we have one that blocks this date */
            Outlook.Folders childFolders = pilotFolder.Folders;
            if(childFolders.Count > 0)
            {
                foreach(Outlook.Folder childFolder in childFolders)
                {
                    if (childFolder.Name.Equals(folderName)) { return ("present"); }
                    if (childFolder.Name.StartsWith(folderName)) { return ("blocked"); }
                }
            }
            return ("absent");
        }

        private void button1_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(1)); }
        private void button2_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(2)); }
        private void button3_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(3)); }
        private void button4_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(4)); }
        private void button5_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(5)); }
        private void button6_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(6)); }
        private void button7_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(7)); }
        private void button8_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(8)); }
        private void button9_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(9)); }
        private void today_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now); }
        private void tomorrow_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(1)); }
        private void Whenever_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Now.AddDays(-1)); }


        /* TODO: Find the least busy day this week and make it happen */
        private void eow_Click(object sender, RibbonControlEventArgs e) { defer(DateTime.Today);  }
       
        private void defer(DateTime date)
        {
            /* Dialog Results */
            DialogResult skipBusy = DialogResult.Yes;
            DialogResult skipWeekends = DialogResult.Yes;

            /* Prompt for Busy? */
            bool promptBusy = true;
            /* Prompt for Weekend? */
            bool promptWeekend = true;

            Outlook.Folder targetFolder;

            /* "Whenever" passes yesterday.  Start scheduling options with tomorrow. */
            if(date < DateTime.Today)
            {
                date = DateTime.Today.AddDays(1);
                promptBusy = false;     // Don't prompt when we run into a busy day.  Just skip over it.
                promptWeekend = false;  // Don't prompt when we come to a weekend.  Just skip over it.
            }

            /* Define the main Pilot folder
            * TODO: Make this configurable */
            Outlook.Folder pilotFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Pilot"] as Outlook.Folder;  

            while(true)
            {
                /* Skip if the date is blocked */
                if (folderStatus(date.ToString("yyyyMMdd")).Equals("blocked")) { date = date.AddDays(1); continue; }

                /* If we're scheduling on a weekend, let's make sure that's what we want */
                if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                {
                    if (promptWeekend)
                    {
                        promptWeekend = false; // Only prompt once
                        skipWeekends = MessageBox.Show("The day you selected is a weekend.  Would you like to reschedule for the next available week day?", "Weekend Work?", MessageBoxButtons.YesNo);
                    }

                    if (skipWeekends == DialogResult.Yes) // We want to reschedule
                    {
                        while (date.DayOfWeek != DayOfWeek.Monday)
                        {
                            date = date.AddDays(1);
                        }
                        Debug.WriteLine("Pushing out to Monday: " + date.ToString("yyyyMMdd"));
                        continue;
                    }
                    else // We do not want to reschedule
                    {
                        break;
                    }
                }

                /* Create our target folder if it is absent */
                if (folderStatus(date.ToString("yyyyMMdd")).Equals("absent")) { pilotFolder.Folders.Add(date.ToString("yyyyMMdd")); }

                /* Grab our targetFolder */
                targetFolder = pilotFolder.Folders[date.ToString("yyyyMMdd")] as Outlook.Folder;

                /* If we're already busy on the target day, see if we want to push things out
                 * TODO: Make this configurable */
                if (targetFolder.Items.Count >= 5)
                {
                    if (promptBusy)
                    {
                        promptBusy = false;
                        skipBusy = MessageBox.Show("You're already pretty busy on " + date.ToString("yyyyMMdd") + ". Would you like to reschedule for your new available day?", "Getting Busy!", MessageBoxButtons.YesNo);
                    }

                    if (skipBusy == DialogResult.Yes) // We want to reschedule
                    {
                        date = date.AddDays(1);
                        continue;
                    }
                    else // We do not want to reschedule
                    {
                        break;
                    }
                }

                /* If we make it this far, we have a date we can use */
                break;
            }

            /* Define targetFolder in case it didn't get defined above */
            targetFolder = pilotFolder.Folders[date.ToString("yyyyMMdd")] as Outlook.Folder;

            /* Get our Selected message / conversation */
            Outlook.Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            Debug.WriteLine("Selection.Count = " + selection.Count);
            Outlook.Selection convHeaders = selection.GetSelection(Outlook.OlSelectionContents.olConversationHeaders) as Outlook.Selection;
            Debug.WriteLine("Selection.Count (ConversationHeaders) = " + convHeaders.Count);

            /* If we have a Conversation selected.  A Conversation is either the collapsed header 
             * of a conversation or a single message that is not part of a conversation */
            if (convHeaders.Count >= 1)
            {
                foreach (Outlook.ConversationHeader convHeader in convHeaders)
                {
                    Outlook.SimpleItems items = convHeader.GetItems();
                    for (int i = 1; i <= items.Count; i++)
                    {
                        Outlook.MailItem mail = items[i] as Outlook.MailItem;
                        mail.UnRead = false;
                      
                        if (mail.Parent.FolderPath != targetFolder.FolderPath)
                        {
                            Debug.WriteLine("Moving " + mail.Subject + " to " + targetFolder.FolderPath);
                            mail.Move(targetFolder);
                        }
                        else
                        {
                            Debug.WriteLine(mail.Subject + " is already in " + targetFolder.FolderPath + ".  Doing nothing.");
                        }
                    }
                }
            }
            else // Otherwise, we have selected a single message that is part of a larger Conversation.
            {
                Outlook.MailItem mail = selection[1] as Outlook.MailItem;
                mail.UnRead = false;
                Debug.WriteLine("Moving " + mail.Subject + " to " + targetFolder.FolderPath);
                if(mail.Parent != targetFolder.Name)
                {
                    mail.Move(targetFolder);
                }
                else
                {
                    Debug.WriteLine(mail.Subject + " is already in " + targetFolder.FolderPath + ".  Doing nothing.");
                }
            }

            removeEmptyFolders();
        }       
    }
}
