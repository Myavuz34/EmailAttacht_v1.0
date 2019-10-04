using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using EmailAttacht_v1._0.Models.Context;
using EmailAttacht_v1._0.Models;
using System.Windows.Forms;

namespace EmailAttacht_v1._0
{
    public partial class ThisAddIn
    {
        DatabaseContext db = new DatabaseContext();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.NewMail += new Microsoft.Office.Interop.Outlook
               .ApplicationEvents_11_NewMailEventHandler(ThisApplication_NewMail);
        }

        private void ThisApplication_NewMail()
        {
            Outlook.MAPIFolder inBox = this.Application.ActiveExplorer()
                .Session.GetDefaultFolder(Outlook
                    .OlDefaultFolders.olFolderInbox);
            Outlook.Items inBoxItems = inBox.Items;
            Outlook.MailItem newEmail = null;
            inBoxItems = inBoxItems.Restrict("[Unread] = true");

            try
            {
                foreach (object collectionItem in inBoxItems)
                {
                    newEmail = collectionItem as Outlook.MailItem;

                    if (newEmail != null && newEmail.Subject == "Report")
                    {
                        if (newEmail.Attachments.Count > 0)
                        {
                            newEmail.UnRead = false;
                            newEmail.Save();

                            for (int i = 1; i <= newEmail
                                                .Attachments.Count; i++)
                            {
                                EmailAttachtLog log = new EmailAttachtLog();
                                var path = @"C:\Users\Artpro\Desktop\excelData\";
                                var fileInfo = new FileInfo(newEmail.Attachments[i].FileName);
                                var fileName = DateTime.Now.ToString("dd-MM-yyyy HH.mm") + "_" + fileInfo.Name;
                                newEmail.Attachments[i].SaveAsFile(path + fileName);


                                var excelPath = path + fileName;
                                if (excelPath != null)
                                {
                                    var pathlog = "Dosya " + "' " + excelPath + " '" + " dizinine kayıt edildi.";
                                    log.Date = DateTime.Now;
                                    log.Level = "Info";
                                    log.Message = pathlog;
                                    db.EmailAttachtsLog.Add(log);
                                    db.SaveChanges();

                                    ImportData(excelPath);
                                    var importlog = "Veriler veritabanına kayıt edildi.";
                                    log.Date = DateTime.Now;
                                    log.Level = "Info";
                                    log.Message = importlog;
                                    db.EmailAttachtsLog.Add(log);

                                    db.SaveChanges();
                                }
                                else
                                {
                                    var errorpath = "Kayıt için belirtilen dizin bulunamadı.";
                                    log.Date = DateTime.Now;
                                    log.Level = "Error";
                                    log.Message = errorpath;
                                    db.EmailAttachtsLog.Add(log);
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                EmailAttachtLog log = new EmailAttachtLog();
                log.Date = DateTime.Now;
                log.Level = "Error";
                log.Exception = ex.ToString();
                log.Message = ex.Message;
                db.EmailAttachtsLog.Add(log);
                db.SaveChanges();
            }
        }


        private void ImportData(string filePath)
        {
            int count = 0;

            try
            {
                var package = new ExcelPackage(new FileInfo(filePath));
                int startColumn = 2;
                int startRow = 6;

                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];


                DatabaseContext db = new DatabaseContext();
                User tblUser = new User();

                var rowCnt = worksheet.Dimension.End.Row - 6;


                for (int i = 0; i <= rowCnt; i++)
                {

                    tblUser.Name = Convert.ToString(worksheet.Cells[startRow, startColumn].Value);
                    tblUser.SurName = Convert.ToString(worksheet.Cells[startRow, startColumn + 1].Value);


                    var isSuccess = SaveClass(tblUser, db);
                    if (isSuccess)
                        count++;
                    startRow++;

                }
            }
            catch (Exception exception)
            {

                EmailAttachtLog log = new EmailAttachtLog();
                log.Date = DateTime.Now;
                log.Level = "Error";
                log.Exception = exception.ToString();
                log.Message = exception.Message;
                db.EmailAttachtsLog.Add(log);
            }
        }
        public bool IsNullValue(string nullValue)
        {
            string pattern = "#DIV/0";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            return regex.IsMatch(nullValue);
        }
        public double ToDouble(string s)
        {
            if (IsNullValue(s) == true)
                return 0;
            else
                return Convert.ToDouble(s);
        }
        public bool SaveClass(User className, DatabaseContext db)
        {
            var result = false;
            try
            {
                var item = new User
                {
                    Name = className.Name,
                    SurName = className.SurName,

                };
                db.Users.Add(item);
                db.SaveChanges();
            }
            catch (Exception exception)
            {

                EmailAttachtLog log = new EmailAttachtLog();
                log.Date = DateTime.Now;
                log.Level = "Error";
                log.Exception = exception.ToString();
                log.Message = exception.Message;
                db.EmailAttachtsLog.Add(log);
                db.SaveChanges();
            }

            return result;
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
