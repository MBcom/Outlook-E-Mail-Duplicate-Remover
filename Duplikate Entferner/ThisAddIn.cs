﻿using System;
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


        /// <summary>
        /// Function called, when a new item is added to Inbox.
        /// </summary>
        /// <param name="Item"></param>
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

        /// <summary>
        /// Function searches for email duplicates from a specific email
        /// </summary>
        /// <param name="filter"></param>
        /// <returns></returns>
        private IEnumerable<Outlook.MailItem> search(Outlook.MailItem filter)
        {
            if (filter.Parent == deletedMailsFolder) yield break;

            HashSet<Outlook.MailItem> mailsFound = new HashSet<Outlook.MailItem>() { filter };

            foreach (Outlook.MailItem m in filter.GetConversation().GetRootItems())
            {
                if (MailItemEquals(m, filter) && m.Parent != deletedMailsFolder)
                {
                    mailsFound.Add(m);
                }
            }

            if (mailsFound.Count > 0)
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


        /// <summary>
        /// Function gets all emails of the inbox and subfolders.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="already_deleted"></param>
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
        }

        /// <summary>
        /// Function deletes all duplicates found in mails dictionary created by 'GetMails'
        /// </summary>
        /// <returns></returns>
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
