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
        Outlook.MAPIFolder mailsfolder;
        Outlook.MAPIFolder mainFolder;
        Outlook._NameSpace oNS;
        Dictionary<string, List<Outlook.MailItem>> mails = new Dictionary<string, List<Outlook.MailItem>>();
        List<Outlook.MailItem> mailsFound;
        bool done = false;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;
            oNS = (Outlook._NameSpace)app.GetNamespace("MAPI");

            mainFolder = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            mainFolder.Items.ItemAdd += Items_ItemAdd;
        }

 
        /***
         * Function called, when a new item is added to Inbox.
         * 
         */
        private void Items_ItemAdd(object Item)
        {
            try
            {
                Outlook.MailItem mail = Item as Outlook.MailItem;
                mailsFound = new List<Outlook.MailItem>();
                search(mail, mainFolder);
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
                            m.Move(oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)); //move to deleted folder
                        }
                    }
                    else
                    {
                        bool first = false;
                        foreach (Outlook.MailItem m in mNichtGelesen) //delete just n-1
                        {
                            if (first)
                            {
                                m.Move(oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)); //move to deleted folder
                            }
                            first = true;
                        }
                    }
                }

            }
            catch (Exception)
            {
            }
        }
        
        /**
         * Function searches for email duplicates from a specific email
         * 
         */
        private void search(Outlook.MailItem filter, Outlook.MAPIFolder folder)
        {
            Outlook.Items items = folder.Items;
            foreach(Outlook.MailItem mitem in items.OfType<Outlook.MailItem>().Where(m => m.Subject == filter.Subject && m.SenderEmailAddress == filter.SenderEmailAddress && m.ReceivedTime == filter.ReceivedTime && m.Body == filter.Body).Select(m => m))
            {
                mailsFound.Add(mitem);
            }
            if (folder.Folders.Count > 0)
            {
                foreach (Outlook.MAPIFolder f in folder.Folders)
                {
                    search(filter, f);
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
        private void GetMails(Outlook.MAPIFolder folder)
        {
            Task t = Task.Factory.StartNew(f =>
            {
                if (folder.Folders.Count > 0)
                {
                    foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                    {
                        GetMails(subFolder);
                    }
                }

                Outlook.Items items = folder.Items;
                foreach (object m in items)
                {
                    try
                    {
                        Outlook.MailItem mail = m as Outlook.MailItem;
                        string hash = getMailProps(mail);
                        if (mails.ContainsKey(hash))
                        {
                            mails[hash].Add(mail);
                        }
                        else
                        {
                            List<Outlook.MailItem> l = new List<Outlook.MailItem>();
                            l.Add(mail);
                            mails.Add(hash, l);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }, folder, TaskCreationOptions.AttachedToParent);
            t.Wait();
        }

        /**
         * Function deletes all duplicates found in mails dictionary created by 'GetMails'
         * 
         */
        public int removeDuplikates()
        {
            if (!done) //checks if GetMails is run once
            {
                GetMails(oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox));
                done = true;
            }
            int i = 0;
            foreach (string s in mails.Keys)
            {
                if (mails[s].Count > 1)
                {
                    List<Outlook.MailItem> mGelesen = new List<Outlook.MailItem>();
                    List<Outlook.MailItem> mNichtGelesen = new List<Outlook.MailItem>();
                    foreach (Outlook.MailItem m in mails[s])
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
                            m.Move(oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)); //move to deleted folder
                            i++;
                        }
                    }
                    else
                    {
                        bool first = false;
                        foreach (Outlook.MailItem m in mNichtGelesen) 
                        {
                            if (first) //delete just n-1
                            {
                                m.Move(oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)); //move to deleted folder
                                i++;
                            }
                            first = true;
                        }
                    }
                }
            }

            return i;
        }

        static string getMailProps(Outlook.MailItem mail)
        {
            return mail.ReceivedTime.ToLongTimeString() + mail.SenderEmailAddress + mail.Subject + mail.Recipients.ToString();
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
