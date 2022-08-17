using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Printing;
using ceTe.DynamicPDF.Printing;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Configuration;

namespace CopyFiles
{
    //Service Request/Project number for this project is TFS24196/SRQ00479913.
    //Devloper of this application is Sanjay Kumar
    //Manager of this Project is Pradeep Kumar Singh
    //GO LIVE date is 17-May-2022
    //CAB# CHN-458
    /*
     This project is targeted to automate the BackupOffice(AHS) manual printing process.
     */
    class Program
    {
        #region Email 
       static string toEmail = ConfigurationManager.AppSettings["ToEmail"];
       static string FromEmail = ConfigurationManager.AppSettings["FromEmail"];
       static string BccEmail = ConfigurationManager.AppSettings["BccEmail"];
       static string Strsmtp = ConfigurationManager.AppSettings["SMTP"];
        #endregion

        #region source ,destination and logs
        static string sourceFilePath = ConfigurationManager.AppSettings["source"];
        
        static string currentDateTm = DateTime.Now.Date.ToString("yyyyMMdd");
        //static string logPath = ConfigurationManager.AppSettings["logPath"] + currentDateTm;
        static string[] allFiles = Directory.GetFiles(sourceFilePath, "*.pdf", SearchOption.TopDirectoryOnly);
        static int totalCount = allFiles.Length;
        #endregion

        public static string GetPrinterFullName(string printerIP, string printerName)
        {
            var server = new PrintServer();
            var queues = server.GetPrintQueues(new[]
            { EnumeratedPrintQueueTypes.Shared, EnumeratedPrintQueueTypes.Connections });
             string fulllName = queues.Where(q => q.Name == printerName &&
             q.QueuePort.Name == printerIP).Select(q => q.FullName).FirstOrDefault();
             return fulllName;
        }             
      
        static void GetPrinterQueue()
        {
           
            string logPath = ConfigurationManager.AppSettings["logPath"] + currentDateTm;
            string PrntName = ConfigurationManager.AppSettings["PrinterName"];
            string toErrorEmail = ConfigurationManager.AppSettings["ToErrorEmail"];
            string printerName = string.Empty;
            string printerPort = string.Empty;

            var server = new PrintServer();                     
            var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            var queu = queues.Where(x => x.Name.Equals(PrntName)).FirstOrDefault();
            //full name:HP LaserJet Enterprise 600 M601 M602 M603 PCL6 Class Driver (Copy 1)(US Ops Printer name) & port=10.136.8.11_2 
            // QueueDriver: HP LaserJet Enterprise 600 M601 M602 M603 PCL6 Class Driver
            // QueuePrintProcessor : winprint and hostingPrintServername: \\QORCWVSIFTU01
            //HP LaserJet Pro MFP M126nw
            //queu.IsTonerLow != true && 
            if (queu.IsOffline != true && queu.IsInError != true && queu.IsOutOfPaper != true && queu.IsPaperJammed != true)
            {
                if (queu.IsTonerLow==true)
                {
                    Logger.SendEmail("Printer Toner is showing low. \n\n <br><br> Thanks\n", "Printer Toner is low !! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);
                }
                Print(PrntName);
            }
            else
            {
                
                string status=PrinterStatus();
                Logger.WriteLog("Some issue in printer \n " + "Printer is offline \n :" + queu.IsOffline + "\n Ink Low \n :" + queu.IsTonerLow, logPath, "_ErroLogs.txt");
                
                Console.WriteLine("Some issue in printer please check! \n");

            }
            // string strEmlBody = Logger.EmailSuccessBody();
            // Logger.SendEmail(strEmlBody, "Print Successfully!", Strsmtp, FromEmail, toEmail, BccEmail);
            #region
            //foreach (var pq in queues)
            //{
            //    Console.WriteLine("{2}\t{0}\t{1}\t{3}\t{4}", "Status :" + pq.IsTonerLow, "processing :" + pq.HasPaperProblem, "Printer Nmae : " + pq.Name, "printer is busy :" + pq.IsBusy, "Printer is offline :" + pq.IsOffline, "Number of queues : " + pq.NumberOfJobs, DateTime.Now.ToString("HH:mm:ss.fff"));
            //}


            //foreach (var queue in queues)
            //{
            //    printerName = queue.Name;//"HP LaserJet Pro MFP M126nw"//NOI2PR01//BETPR06
            //    printerPort = queue.QueuePort.Name;//10.73.26.10//10.134.15.6

            //}

            //PrintQueue queu = queues.Where(x => x.Name.Equals(printerName)).FirstOrDefault();

            //string printerFullName = GetPrinterFullName("10.73.26.10", "NOI2PR01");


            //Print("HP LaserJet Pro MFP M126nw");//NOI2PR01 on QORCWVSIPRINT01.IQOR.QOR.COM
            #endregion
        }
        #region In this function we are checking what is the status of Printer, then we are printing the PDF.

        static string PrinterStatus()
        {
            string printerName = string.Empty;
            string printerPort = string.Empty;
            string status = string.Empty;

            string toErrorEmail = ConfigurationManager.AppSettings["ToErrorEmail"];
            string FromEmail = ConfigurationManager.AppSettings["FromEmail"];
            string BccEmail = ConfigurationManager.AppSettings["BccEmail"];
            string Strsmtp = ConfigurationManager.AppSettings["SMTP"];

            var server = new PrintServer();
            var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            var queu = queues.Where(x => x.Name.Equals("HP LaserJet Enterprise 600 M601 M602 M603 PCL6 Class Driver (Copy 1)")).FirstOrDefault();

            
            if (queu.IsOffline == true)
            {
                status = "fail";
               Logger.SendEmail("There is Some issue in Printer, showing Offline. \n\n <br> <br>Thanks\n", "Printer is offline ! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);
            }
            //else if (queu.IsTonerLow == true)
            //{
            //    status = "pass";
            //    Logger.SendEmail("Printer Toner is showing low. \n\n <br><br> Thanks\n", "Printer Toner is low !! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);

            //}
            else if (queu.IsInError == true)
            {
                status = "fail";
                Logger.SendEmail("There is some error in Printer. \n\n <br> <br> Thanks\n", "Error in Printer!! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);
            }
            else if (queu.IsOutOfPaper == true)
            {
                status = "fail";
                Logger.SendEmail("Printer is OutOfPaper, please check.\n\n <br><br>Thanks\n", "Printer is OutOfPaper!! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);
                //return status;
            }
            else if (queu.IsPaperJammed == true)
            {
                status = "fail";
                Logger.SendEmail("Printer is Paper Jammed, please check.\n\n <br><br>Thanks\n", "Printer is PaperJammed!! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);
                //return status;
            }
            else
            {
                status = "pass";
            }
            return status;
        }
        #endregion

        #region Here checking PDF files and sending Print CMD to Printer, incase there is any issue in printer notifying to OPS
        public static void Print(string PrinterName)
        {
            string logPath = ConfigurationManager.AppSettings["logPath"] + currentDateTm;
                     

            // Print the files.
            foreach (string filename in allFiles)
            {
                string Status = string.Empty;
                try
                {
                    Status = PrinterStatus();
                    if (Status == "pass")
                    {

                        
                        Logger.WriteLog("Printing: " + filename, logPath, "_PrintingFile.txt");

                        FileInfo file_into = new FileInfo(filename);
                        string short_name = file_into.Name;

                        // Read the file's contents

                        StringBuilder text = new StringBuilder();
                        //var FileContents = File.ReadAllText(filename).Trim();
                        using (PdfReader reader = new PdfReader(filename))
                        {
                            for (int pageNo = 1; pageNo <= reader.NumberOfPages; pageNo++)
                            {
                                text.Append(PdfTextExtractor.GetTextFromPage(reader, pageNo));

                            }
                            reader.Close();
                        }

                        // Printing the file
                        PrintJob printJob = new PrintJob(PrinterName, filename);
                        printJob.Print();//1236pm
                        printJob.Dispose();

                        Logger.WriteLog(short_name, logPath, "_Print_Success.txt");
                        FileMove(filename);

                    }
                    else
                    { break; }
                }
                catch (Exception ex)
                {
                    Logger.WriteLog("Error during file reading and printing" + ex.Message + filename, logPath, "_ErroLogs.txt");
                    Logger.SendEmail("There is Some issue for \n" + ex.Message, "Error in printing the file! ", Strsmtp, FromEmail, toEmail, BccEmail);
                    
                }
               
            }
            
            
            // FileMove();  
        }
        #endregion

        #region Here we are moving printed files in Archive folder
        public static void FileMove(string path)
        {
            string logPath = ConfigurationManager.AppSettings["logPath"] + currentDateTm;
            string destination2 = ConfigurationManager.AppSettings["destination"]+ currentDateTm+"\\";
            //string[] files = Directory.GetFiles(sourceFilePath);
            try
            {
                if (!Directory.Exists(destination2))
                {
                    Directory.CreateDirectory(destination2);
                }
                //foreach (string file in allFiles)
                //{
                try
                {
                    string filname = System.IO.Path.GetFileName(path);
                    destination2 = destination2 + filname;
                    Console.WriteLine(path + " = " + destination2);
                    // Ensure that the target does not exist.
                    if (File.Exists(destination2))
                    {
                        File.Delete(destination2);
                        File.Move(path, destination2);

                    }
                    else
                    { File.Move(path, destination2); }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Logger.WriteLog("Some issue to moving file" + ex.Message, logPath, "_FileMove_ErrorLogs.txt");
                }
                //}
            }
            catch (Exception ex)
            {
                Logger.WriteLog("Some issue to moving file"+ex.Message,logPath,"_FileMove_ErrorLogs.txt");
            }
        }
        #endregion

        static void Main(string[] args)
        {
            string toErrorEmail = ConfigurationManager.AppSettings["ToErrorEmail"];
            string currentHour = DateTime.Now.ToString("HH");
            string ReportHour = ConfigurationManager.AppSettings["ReportHour"];

            GetPrinterQueue();
            string logPath = ConfigurationManager.AppSettings["logPath"] + currentDateTm;
           // Logger.WriteLog("App is runing!!!", logPath, "_RuningStatusLogs.txt");
            if (currentHour == ReportHour)
            {
                
                string strEmlBody = Logger.EmailSuccessBody();
                Logger.SendEmail(strEmlBody, "Letter Printing Report!!! ", Strsmtp, FromEmail, toErrorEmail, BccEmail);
            }
            Logger.WriteLog("App is runing!!!", logPath, "_RuningStatusLogs.txt");
            Console.WriteLine("Test");
            //Console.ReadLine();
        }
    }
}
