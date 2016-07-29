using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace MailSortOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer = null;

        private void ThisAddIn_Startup
            (object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (CurrentExplorer_Event);
        }

        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder =
                this.Application.ActiveExplorer().CurrentFolder;
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selectedObject = this.Application.ActiveExplorer().Selection[1];
                    if (selectedObject is Outlook.MailItem)
                    {
                        //Outlook.MailItem mailItem = (selectedObject as Outlook.MailItem);
                        //string senderEmail = mailItem.SenderEmailAddress;
                        //string senderName = mailItem.SenderName;
                        //Outlook.Folders folders = GetAllFolders();

                        //Outlook.MAPIFolder matchingFolder = FindMatchingFolder(senderName, folders);
                        //if (matchingFolder != null)
                        //{
                        //    MessageBox.Show("Matching folder was found " + matchingFolder.Name);
                        //    mailItem.Move(matchingFolder);
                        //}
                        //else
                        //{
                        //    MessageBox.Show("Matching folder wasn't found");
                        //}
                        MoveItems();
                    }
                }
            }
            catch (Exception exception)
            {
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        private Outlook.Folders GetAllFolders()
        {
            Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            Outlook.Folders childFolders = root.Folders;
            string message = "";
            foreach (Outlook.MAPIFolder folder in childFolders)
            {
                message += folder.Name + " | ";
            }
            MessageBox.Show(message);
            return childFolders;
        }

        private Outlook.MAPIFolder FindMatchingFolder(string senderName, Outlook.Folders folders)
        {
            Outlook.MAPIFolder matchingFolder = null;
            foreach (Outlook.MAPIFolder folder in folders)
            {
                if (folder.Name.Equals(senderName))
                {
                    matchingFolder = folder;
                }
            }
            return matchingFolder;
        }


        private void MoveItems()
        {
            Outlook.Folders allFolders = GetAllFolders();
            Outlook.NameSpace myNameSpace = Application.GetNamespace("MAPI");
            Outlook.MAPIFolder InboxFolder = allFolders[32];
            MessageBox.Show(InboxFolder.Name);
            //Outlook.MAPIFolder InboxFolder = myNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items mailItems = InboxFolder.Items;



            foreach (Outlook.MailItem mailItem in mailItems)
            {
                MessageBox.Show(!mailItem.UnRead + "");
                if (!mailItem.UnRead)
                {
                    string senderName = mailItem.SenderName;
                    MessageBox.Show(senderName);
                    Outlook.MAPIFolder matchingFolder = FindMatchingFolder(senderName, allFolders);
                    if (matchingFolder != null)
                    {
                        MessageBox.Show("Matching folder was found " + matchingFolder.Name);
                        mailItem.Move(matchingFolder);
                        mailItem.p
                        Outlook.MAPIFolder myDestFolder = InboxFolder.Folders["verplaatsmap"];
                        MessageBox.Show(myDestFolder.FolderPath);
                        //mailItem.Move(myDestFolder);
                        //Outlook.MailItem copyOfMailItem = mailItem.Copy();
                        Console.WriteLine("HOIHOIHIO");
                        try
                        {
                            MessageBox.Show("Cool!!!!");
                            //copyOfMailItem.Move(myDestFolder);
                            MessageBox.Show("coolio!!!!");
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show("Failure!!!!");
                        }
                    }
                    else
                    {
                        //MessageBox.Show("Matching folder wasn't found");
                    }
                }
                else {
                    // mail wasn't read yet
                }
            }
        }
    }
}
