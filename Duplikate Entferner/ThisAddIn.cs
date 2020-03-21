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

namespace Duplikate_Entferner
{
    public partial class ThisAddIn
    {
        Outlook.Application app;
        Outlook.MAPIFolder deletedMailsFolder;
        Outlook.MAPIFolder mainFolder;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;

            deletedMailsFolder = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);

            mainFolder = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            mainFolder.Items.ItemAdd += Items_ItemAdd;
        }


        /***
         * Function called, when a new item is added to Inbox.
         * 
         */
        private void Items_ItemAdd(object Item)
        {
            if (!Properties.Settings.Default.auto_delete) return;
            try
            {
                Outlook.MailItem mail = Item as Outlook.MailItem;

                search(mail);
            }
            catch (Exception)
            {
            }
        }

        /**
         * Function searches for email duplicates from a specific email
         * 
         */
        private IEnumerable<Outlook.MailItem> search(Outlook.MailItem filter)
        {
            List<Outlook.MailItem> mailsFound = new List<Outlook.MailItem>();

            foreach (Outlook.MailItem m in filter.GetConversation().GetRootItems())
            {
                if (m.EntryID != filter.EntryID && m.Subject == filter.Subject && m.SenderEmailAddress == filter.SenderEmailAddress && m.ReceivedTime == filter.ReceivedTime && m.Body == filter.Body && m.Parent != deletedMailsFolder)
                {
                    mailsFound.Add(m);
                }
            }

            if (mailsFound.Count > 1)
            {
                List<Outlook.MailItem> mGelesen = new List<Outlook.MailItem>(); //List for readed emails
                List<Outlook.MailItem> mNichtGelesen = new List<Outlook.MailItem>(); //List for unreaded
                foreach (Outlook.MailItem m in mailsFound)
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
                    filter.Move(deletedMailsFolder);
                    yield return filter;
                }
                else
                {//no email were already readed
                    bool first = false;
                    foreach (Outlook.MailItem m in mNichtGelesen) //delete just n-1
                    {
                        if (first)
                        {
                            m.Move(deletedMailsFolder); //move to deleted folder
                            yield return m;
                        }
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
                        yield return m;
                    }
                    first2 = true;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }


        /**
         * Function gets all emails of the inbox and subfolders.
         * 
         */
        private void GetMails(Outlook.MAPIFolder folder, ref HashSet<string> already_deleted)
        {
            if (folder.Folders.Count > 0)
            {
                foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                {
                    GetMails(subFolder, ref already_deleted);
                }
            }

            

            Outlook.Items items = folder.Items;
            foreach (object mm in items)
            {
                try
                {
                    Outlook.MailItem mail = mm as Outlook.MailItem;

                    //jump if already deleted
                    if (already_deleted.Contains(mail.EntryID)) continue;

                    foreach (var item in search(mail))
                    {
                        already_deleted.Add(item.EntryID);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        /**
         * Function deletes all duplicates found in mails dictionary created by 'GetMails'
         * 
         */
        public int removeDuplikates()
        {
            HashSet<string> already_deleted = new HashSet<string>();
            GetMails(mainFolder, ref already_deleted);

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
