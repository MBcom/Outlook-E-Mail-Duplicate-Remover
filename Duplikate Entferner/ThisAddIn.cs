using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;

namespace Duplikate_Entferner
{
    public partial class ThisAddIn
    {
        Outlook.Application app;
        Outlook.MAPIFolder deletedMailsFolder;
        Outlook.MAPIFolder mainFolder;
        bool completeSearchRunning = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;

            deletedMailsFolder = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);

            mainFolder = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            AddEventListenerForFolder(mainFolder as Outlook.Folder);

            Thread t = new Thread(() =>
            {
                while (true)
                {
                    //remove duplicates in new emails every 2 minutes
                    Thread.Sleep(120000);
                    if (Properties.Settings.Default.auto_delete)
                    {
                        RemoveDuplikates(true);
                    } 
                }
            });
            t.Start();
        }

        /// <summary>
        /// Adds event listener to specified folder and subfolders.
        /// </summary>
        /// <param name="folder"></param>
        private void AddEventListenerForFolder(Outlook.Folder folder)
        {
            folder.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
            folder.Items.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);

            foreach (Outlook.Folder f in folder.Folders)
            {
                AddEventListenerForFolder(f);
            }
        }

        /// <summary>
        /// Function called, when a new item is changed.
        /// </summary>
        /// <param name="Item"></param>
        private void Items_ItemChange(object Item)
        {
            if (!completeSearchRunning)
            {
                Items_ItemAdd(Item);
            }
        }


        /// <summary>
        /// Function called, when a new item is added to Inbox.
        /// </summary>
        /// <param name="Item"></param>
        private void Items_ItemAdd(object Item)
        {
            if (!Properties.Settings.Default.auto_delete) return;

            if (!(Item is Outlook.MailItem)) return;

            Task t = Task.Factory.StartNew(() =>
            {
                try
                {
                    Outlook.MailItem mail = Item as Outlook.MailItem;

                    return search(mail).Count();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
                return -1;
            }).ContinueWith(r =>
            {
                Debug.WriteLine("Email check completed" + r.Result);
            });


        }

        /// <summary>
        /// Returns true if mail items are probably equal
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        private bool MailItemEquals(Outlook.MailItem a, Outlook.MailItem b)
        {
            if (a.Subject != b.Subject || a.SenderEmailAddress != b.SenderEmailAddress || a.BodyFormat != b.BodyFormat || a.ReceivedTime != b.ReceivedTime)
            {
                return false;
            }
            switch (a.BodyFormat)
            {
                case Outlook.OlBodyFormat.olFormatPlain:
                    return a.Body == b.Body;
                case Outlook.OlBodyFormat.olFormatHTML:
                    return a.HTMLBody == b.HTMLBody;
                case Outlook.OlBodyFormat.olFormatRichText:
                    return a.RTFBody == b.RTFBody;
            }
            return true;
        }

        private void EnumerateConversation(object item, Outlook.Conversation conversation, Outlook.MailItem filter, ref Dictionary<string, Outlook.MailItem> mailsFound)
        {
            Outlook.SimpleItems items = conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    //  enumerate only MailItem type.
                    if (myItem is Outlook.MailItem)
                    {
                        Outlook.MailItem m = myItem as Outlook.MailItem;
                        Outlook.Folder inFolder = m.Parent as Outlook.Folder;
                        string msg = m.Subject + " in folder " + inFolder.Name;

                        Debug.WriteLine(msg);

                        if (!mailsFound.ContainsKey(m.EntryID) && MailItemEquals(m, filter) && inFolder != deletedMailsFolder)
                        {
                            mailsFound.Add(m.EntryID, m);
                        }
                    }
                    // Continue recursion.
                    EnumerateConversation(myItem, conversation, filter, ref mailsFound);
                }
            }
        }

        /// <summary>
        /// Function searches for email duplicates from a specific email
        /// </summary>
        /// <param name="filter"></param>
        /// <returns></returns>
        private IEnumerable<Outlook.MailItem> search(Outlook.MailItem filter)
        {
            if (filter.Parent == deletedMailsFolder) yield break;

            Dictionary<string, Outlook.MailItem> mailsFound = new Dictionary<string, Outlook.MailItem>() { { filter.EntryID, filter } };

            // Obtain a Conversation object.
            Outlook.Conversation conv = filter.GetConversation();

            // Obtain Table that contains rows 
            // for each item in Conversation.
            Outlook.Table table = conv.GetTable();

            //break if there just one
            if (table.GetRowCount() == 1) yield break;

            Debug.WriteLine("Conversation Items Count: " + table.GetRowCount().ToString());

            // Obtain root items and enumerate Conversation.
            Outlook.SimpleItems simpleItems = conv.GetRootItems();
            foreach (object item in simpleItems)
            {
                // enumerate only MailItem type.
                if (item is Outlook.MailItem)
                {
                    Outlook.MailItem m = item as Outlook.MailItem;
                    Outlook.Folder inFolder = m.Parent as Outlook.Folder;
                    string msg = m.Subject + " in folder " + inFolder.Name;

                    Debug.WriteLine(msg);

                    if (!mailsFound.ContainsKey(m.EntryID) && MailItemEquals(m, filter) && inFolder != deletedMailsFolder)
                    {
                        mailsFound.Add(m.EntryID, m);
                    }
                }
                // Call EnumerateConversation 
                // to access child nodes of root items.
                EnumerateConversation(item, conv, filter, ref mailsFound);
            }

            if (mailsFound.Count > 1)
            {
                //delete duplicates //O(((n-1)n)/2)
                var mf = mailsFound.ToArray();
                for (int i = 0; i < mf.Length - 1; i++)
                {
                    for (int j = i + 1; j < mf.Length; j++)
                    {
                        if (app.Session.CompareEntryIDs(mf[i].Key, mf[j].Key))
                        {
                            mailsFound.Remove(mf[j].Key);
                        }
                    }
                }

                List<Outlook.MailItem> mGelesen = new List<Outlook.MailItem>(); //List for readed emails
                List<Outlook.MailItem> mNichtGelesen = new List<Outlook.MailItem>(); //List for unreaded
                foreach (Outlook.MailItem m in mailsFound.Values)
                {
                    if (m.UnRead)
                    {
                        mNichtGelesen.Add(m);
                    }
                    else
                    {
                        mGelesen.Add(m);
                    }
                }
                if (mGelesen.Count > 0) // if there are readed emails, move the new unreaded to trash
                {
                    foreach (Outlook.MailItem m in mNichtGelesen)
                    {
                        m.Move(deletedMailsFolder); //move to deleted folder
                        yield return m;
                    }
                }
                else
                {//no email were already readed
                    bool first = false;
                    foreach (Outlook.MailItem m in mNichtGelesen) //delete just n-1
                    {
                        if (first)
                        {
                            m.Move(deletedMailsFolder); //move to deleted folder  
                        }
                        yield return m;
                        first = true;
                    }
                }
                //clear readed mails too
                bool first2 = false;
                foreach (Outlook.MailItem m in mGelesen) //delete just n-1
                {
                    if (first2)
                    {
                        m.Move(deletedMailsFolder); //move to deleted folder
                    }
                    yield return m;
                    first2 = true;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }


        /// <summary>
        /// Function gets all emails of the inbox and subfolders.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="already_deleted"></param>
        /// <param name="unreadedOnly">Search only in unread messages</param>
        private void GetMails(Outlook.MAPIFolder folder, ref HashSet<string> already_deleted, bool unreadedOnly)
        {
            Outlook.Items items;
            if (unreadedOnly)
            {
                //take only unreaded messages
                items = folder.Items.Restrict("[Unread]=true");
            }
            else
            {
                items = folder.Items;
            }

            foreach (object mm in items)
            {
                if (!(mm is Outlook.MailItem)) continue;

                try
                {
                    Outlook.MailItem mail = mm as Outlook.MailItem;

                    //jump if already deleted
                    if (mail.Parent == deletedMailsFolder) continue;

                    foreach (var item in search(mail))
                    {
                        already_deleted.Add(item.EntryID);
                    }
                }
                catch (Exception)
                {
                }
            }

            Debug.WriteLine($"{items.Count} items processed in {folder.Name}");

            if (folder.Folders.Count > 0)
            {
                foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                {
                    GetMails(subFolder, ref already_deleted, unreadedOnly);
                }
            }
        }

        /// <summary>
        /// Function deletes all duplicates found in mails dictionary created by 'GetMails'
        /// </summary>
        /// <param name="searchInUnreadOnly">Searches only in unread messages.</param>
        /// <returns></returns>
        public int RemoveDuplikates(bool searchInUnreadOnly = false)
        {
            completeSearchRunning = true;
            HashSet<string> already_deleted = new HashSet<string>();
            GetMails(mainFolder, ref already_deleted, searchInUnreadOnly);
            completeSearchRunning = false;

            return already_deleted.Count;
        }



        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
