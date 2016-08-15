using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

//Include namespace for StringBuilder class
using System.Text;

//Include namespace for sending email
using System.Net;
using System.Net.Mail;

//Include namespaces needed for reading data from excel
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using System.Data;
namespace Test
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_click(object sender, EventArgs e)
        {
            //Get path from web.config file to upload
            string FilePath = ConfigurationManager.AppSettings["FilePath"].ToString();
            string filename = string.Empty;
            //To check whether file is selected or not to uplaod
            if (FileUploadtoServer.HasFile)
            {
                try
                {
                    string[] allowdFile = { ".xls", ".xlsx" };
                    //Here we are allowing only excel file so verifying selected file is in excel format or not
                    string FileExt = System.IO.Path.GetExtension(FileUploadtoServer.PostedFile.FileName);
                    //Check whether selected file is valid extension or not
                    bool isValidFile = allowdFile.Contains(FileExt);
                    if (!isValidFile)
                    {
                        lblMsg.ForeColor = System.Drawing.Color.Red;
                        lblMsg.Text = "Please upload only Excel";
                    }
                    else
                    {
                        lblMsg.Text = "";
                        // Get size of uploaded file, here restricting size of file
                        int FileSize = FileUploadtoServer.PostedFile.ContentLength;
                        if (FileSize <= 1048576)//1048576 byte = 1MB
                        {
                            //Get file name of selected file
                            filename = Path.GetFileName(Server.MapPath(FileUploadtoServer.FileName));

                            //Save selected file into server location
                            FileUploadtoServer.SaveAs(Server.MapPath(FilePath) + filename);
                            //Get file path
                            string filePath = Server.MapPath(FilePath) + filename;
                            //Open the connection with excel file based on excel version
                            OleDbConnection con = null;
                            if (FileExt == ".xls")
                            {
                                con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=Excel 8.0;");

                            }
                            else if (FileExt == ".xlsx")
                            {
                                con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0;");
                            }
                            con.Open();
                            //Get the list of sheet available in excel sheet
                            DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            //Get first sheet name
                            string getExcelSheetName = dt.Rows[0]["Table_Name"].ToString();
                            //Select rows from first sheet in excel sheet and fill into dataset
                            OleDbCommand ExcelCommand = new OleDbCommand(@"SELECT * FROM [" + getExcelSheetName + @"]", con);
                            OleDbDataAdapter ExcelAdapter = new OleDbDataAdapter(ExcelCommand);


                            DataTable ExcelDataSet = new DataTable();

                            //Storing data into ViewState variable so that can access this into Sending mail button block
                            ViewState["ExcellFileData"] = ExcelDataSet;
                            ExcelAdapter.Fill(ExcelDataSet);
                            con.Close();
                            //Bind the dataset into gridview to display excel contents
                            GridView1.DataSource = ExcelDataSet;
                            GridView1.DataBind();
                            //Enabling the visiblity of Send email button
                            btnIncrementLetter.Visible = true;
                        }
                        else
                        {
                            lblMsg.Text = "Attachment file size should not be greater then 1 MB!";
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMsg.Text = "Error occurred while uploading a file: " + ex.Message;
                }
            }
            else
            {
                lblMsg.Text = "Please select a file to upload.";
            }
        }

        protected void btnIncrementLetter_Click(object sender, EventArgs e)
        {

            //Check ViewState variable has value or not

            if (ViewState["ExcellFileData"] != null)
            {

                DataTable dt = (DataTable)ViewState["ExcellFileData"];
                
                if (dt.Rows.Count > 0)
                {

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        //do something here to iterate each rows

                        string Emails = dt.Rows[i]["Email"].ToString();
                        string FullName = dt.Rows[i]["FullName"].ToString();
                        // Specify the from and to email address
                        MailMessage mailMessage = new MailMessage
                            ("sendmailtoemployee@gmail.com", Emails);

                        //Creating subject of the email using StringBuilder class
                        StringBuilder mailBody = new StringBuilder();
                        mailBody.AppendFormat("<h1>Congratulations! We have awarded you salary increment for the years 2016</h1>");
                        mailBody.AppendFormat("Dear {0}," ,FullName);
                        mailBody.AppendFormat("<br />");
                        mailBody.AppendFormat("<p>Keeping in view your best performance during the year 2014 and 2015. The company management is very happy to notify you that your salary has been increased.Please read the reference letter for more details</p>");

                        // Specify the email body
                        mailMessage.Body = mailBody.ToString();
                        // Specify the email Subject
                        mailMessage.Subject = "Salary increment letter for the years 2016";

                        //Attaching Increment letter to be sent to all employees

                        Attachment at = new Attachment(Server.MapPath("~/UploadedFiles/SalaryIncreaseLetter.docx"));
                        mailMessage.Attachments.Add(at);

                        // No need to specify the SMTP settings as these 
                        // are already specified in web.config
                        SmtpClient smtpClient = new SmtpClient();
                        // Finall send the email message using Send() method
                        smtpClient.Send(mailMessage);
                    }
                }




                lblEmailAlertMsg.ForeColor = System.Drawing.Color.Green;
                lblEmailAlertMsg.Text = "Emails with increment letter sent to above fetched employees!!";


            }



        }
    }
}
