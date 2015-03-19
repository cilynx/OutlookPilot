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
         * Weekend Detection / Warning / Rescheduling
         * Busy Day Detection / Warning / Rescheduling
         * Calendar-based Busy Calculations
         */

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) 
        {
            removeEmptyFolders();
        }

        private void removeEmptyFolders()
        {
            /* TODO: Make this configurable */
            Outlook.MAPIFolder pilotFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Pilot"];
            
            /* Remove any empty Pilot folders */
            for (int i=1; i <= pilotFolder.Folders.Count; i++)
            {
                if(pilotFolder.Folders[i].Items.Count == 0)
                {
                    Debug.WriteLine(pilotFolder.Folders[i].Name + " is empty and we are deleting it");
                    pilotFolder.Folders[i].Delete();
                }
            }
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
            DialogResult dr;
            /* Prompt for Busy? */
            bool promptBusy = true;
            /* Prompt for Weekend? */
            bool promptWeekend = true;

            /* "Whenever" passes yesterday.  Start scheduling options with tomorrow. */
            if(date < DateTime.Today)
            {
                date = DateTime.Today.AddDays(1);
                promptBusy = false;
                promptWeekend = false;
            }

            /* If we're deferring to a weekend, let's make sure that's what we want */
            if(date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
            {
                dr = DialogResult.Yes;
                if (promptWeekend)
                {
                    dr = MessageBox.Show("The day you selected is a weekend.  Would you like to reschedule for the next available week day?", "Weekend Work?", MessageBoxButtons.YesNo);
                }
                if(dr == DialogResult.Yes)
                {
                    promptBusy = false;                      
                    while(date.DayOfWeek != DayOfWeek.Monday)
                    {
                        date = date.AddDays(1);
                    }
                    Debug.WriteLine("Pushing out to Monday: " + date.ToString("yyyyMMdd"));
                }
            }

            /* Define the main Pilot folder
             * TODO: Make this configurable */
            Outlook.MAPIFolder pilotFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Pilot"];            

            /* I like this format better than MailPilot's, but it means I should implement a MailPilot compatibility mode 
             * such that MailPilot and OutlookPilot can be used on the same set of folders and play nice together */
            String folderName = date.ToString("yyyyMMdd");

            /* Create our targetFolder if it does not already exist then grab the object */
            try { pilotFolder.Folders.Add(folderName); }
            catch { Debug.WriteLine("Folder already exists: " + folderName); }
            Outlook.MAPIFolder targetFolder = pilotFolder.Folders[folderName];

            /* If we're already really busy on the target day, see if we want to push things out
             * TODO: Make this configurable */
            while(targetFolder.Items.Count >= 5)
            {
                dr = DialogResult.Yes; // Set to yes so we reschedule by default if the prompt is disabled
                if(promptBusy)
                {
                    dr = MessageBox.Show("You're already pretty busy on " + folderName + ". Would you like to reschedule for your new available day?", "Getting Busy!", MessageBoxButtons.YesNo);
                }
                if (dr == DialogResult.Yes) // We want to reschedule
                {
                    promptBusy = false;
                    date = date.AddDays(1);
                    if (date.DayOfWeek == DayOfWeek.Saturday) { date = date.AddDays(2); } // Don't automagically schedule things on weekends
                    folderName = date.ToString("yyyyMMdd");
                    try { pilotFolder.Folders.Add(folderName); }
                    catch { Debug.WriteLine("Folder already exists: " + folderName); }
                    targetFolder = pilotFolder.Folders[folderName];
                }
                else // We do not want to reschedule
                {
                    break;
                }
            }

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

        private void NextWeek_Click(object sender, RibbonControlEventArgs e)
        {

        }

       
    }
}
