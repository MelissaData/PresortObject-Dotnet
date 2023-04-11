using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.ComponentModel.Design;
using MelissaData;
using System.Threading;
using System.Linq;

namespace MelissaPresortObjectWindowsDotnet
{
  class Program
  {
    static void Main(string[] args)
    {
      // Variables
      string license = "";
      string testPresortFile = "";
      string dataPath = "";

      ParseArguments(ref license, ref testPresortFile, ref dataPath, args);
      RunAsConsole(license, testPresortFile, dataPath);
    }

    static void ParseArguments(ref string license, ref string testPresortFile, ref string dataPath, string[] args)
    {
      for (int i = 0; i < args.Length; i++)
      {
        if (args[i].Equals("--license") || args[i].Equals("-l"))
        {
          if (args[i + 1] != null)
          {
            license = args[i + 1];
          }
        }
        if (args[i].Equals("--dataPath") || args[i].Equals("-d"))
        {
          if (args[i + 1] != null)
          {
            dataPath = args[i + 1];
          }
        }
        if (args[i].Equals("--file") || args[i].Equals("-f"))

        {
          if (args[i + 1] != null)
          {
            testPresortFile = args[i + 1];
          }
        }
      }
    }

    static void RunAsConsole(string license, string testPresortFile, string dataPath)
    {
      Console.WriteLine("\n\n========= WELCOME TO MELISSA PRESORT OBJECT WINDOWS DOTNET =========\n");

      PresortObject presortObject = new PresortObject(license, dataPath);

      bool shouldContinueRunning = true;

      if (presortObject.mdPresortObj.GetInitializeErrorString() != "No Errors")
      {
        shouldContinueRunning = false;
      }

      while (shouldContinueRunning)
      {
        DataContainer dataContainer = new DataContainer();

        if (string.IsNullOrEmpty(testPresortFile))
        {
          Console.WriteLine("\nFill in each value to see the Presort Object results");
          Console.Write("Presort file path:");

          dataContainer.PresortFile = Console.ReadLine();
          dataContainer.FormatPresortOutputFile();
        }
        else
        {
          dataContainer.PresortFile = testPresortFile;
          dataContainer.FormatPresortOutputFile();
        }

        // Print user input
        Console.WriteLine("\n============================== INPUTS ==============================\n");
        Console.WriteLine($"\t         Presort File: {dataContainer.PresortFile}");
       
        // Execute Presort Object
        presortObject.ExecuteObjectAndResultCodes(ref dataContainer);

        // Print output
        Console.WriteLine("\n============================== OUTPUT ==============================\n");
        Console.WriteLine("\n\tPresort Object Information:");

        List<string> sections = dataContainer.GetWrapped();

        Console.WriteLine($"\t          Output File: {sections[0]}");

        for (int i = 1; i < sections.Count; i++) 
        {
          if((i == sections.Count - 1) && sections[i].EndsWith("\\"))
          {
            sections[i] = sections[i].Substring(0, sections[i].Length - 1);
          }
          Console.WriteLine($"\t                       {sections[i]}");
        }

        bool isValid = false;

        if (!string.IsNullOrEmpty(testPresortFile))
        {
          isValid = true;
          shouldContinueRunning = false;
        }

        while (!isValid)
        {
          Console.WriteLine("\nTest another file? (Y/N)");
          string testAnotherResponse = Console.ReadLine();

          if (!string.IsNullOrEmpty(testAnotherResponse))
          {
            testAnotherResponse = testAnotherResponse.ToLower();
            if (testAnotherResponse == "y")
            {
              isValid = true;
            }
            else if (testAnotherResponse == "n")
            {
              isValid = true;
              shouldContinueRunning = false;
            }
            else
            {
              Console.Write("Invalid Response, please respond 'Y' or 'N'");
            }
          }
        }
      }
      Console.WriteLine("\n============= THANK YOU FOR USING MELISSA DOTNET OBJECT ============\n");
    }
  }

  class PresortObject
  {
    // Path to Presort Object data files (.dat, etc)
    string dataFilePath;

    // Create instance of Melissa Presort Object
    //public PresortObjLib.PresortCheckClass presortObj = new PresortObjLib.PresortCheckClass();
    public mdPresort mdPresortObj = new mdPresort();

    public PresortObject(string license, string dataPath)
    {
      // Set license string and set path to data files  (.dat, etc)
      mdPresortObj.SetLicenseString(license);
      dataFilePath = dataPath;
      mdPresortObj.SetPathToPresortDataFiles(dataFilePath);

      // If you see a different date than expected, check your license string and either download the new data files or use the Melissa Updater program to update your data files.  
      //PresortObjLib.ProgramStatus pStatus = presortObj.InitializeDataFiles();
      mdPresort.ProgramStatus pStatus = mdPresortObj.InitializeDataFiles();

      // If an issue occurred while initializing the data files, this will throw
      if (pStatus != mdPresort.ProgramStatus.ErrorNone)
      {
        Console.WriteLine("Failed to Initialize Object.");
        Console.WriteLine(pStatus);
        return;
      }

      Console.WriteLine($"                DataBase Date: {mdPresortObj.GetDatabaseDate()}");
      Console.WriteLine($"              Expiration Date: {mdPresortObj.GetLicenseStringExpirationDate()}");

      /**
       * This number should match with the file properties of the Melissa Object binary file.
       * If TEST appears with the build number, there may be a license key issue.
       */
      Console.WriteLine($"               Object Version: {mdPresortObj.GetBuildNumber()}\n");
    }

    // This will call the functions to process the input name as well as generate the result codes
    public void ExecuteObjectAndResultCodes(ref DataContainer data)
    {
      /* PreSortSettings to the desired type of Presort 
       * The sample defaults to First Class Mailing, if you'd like to use Standard, comment out the current line and uncomment the one containing the enumeration: STD_LTR_AUTO
       * STD_LTR_AUTO is for Standard Mail, Letter, Automation pieces
       * FCM_LTR_AUTO = First Class Mail, Letter, Automation pieces
       */
      mdPresortObj.SetPreSortSettings((int)mdPresort.SortationCode.FCM_LTR_AUTO);
      //presortObj.PreSortSettings(PresortObjLib.SortationCode.STD_LTR_AUTO);

      /* Sack and Parcel Dimensions */
      mdPresortObj.SetSackWeight("30");
      mdPresortObj.SetPieceLength("9");
      mdPresortObj.SetPieceHeight("4.5");
      mdPresortObj.SetPieceThickness("0.042");
      mdPresortObj.SetPieceWeight("1.5");

      /* Mailers ID 
       * Insert a valid 6-9 digit MailersID number
       * If you do not have a valid Mailers ID, you can visit the USPS to apply for one. Go to usps.com and click on the 'Business Customer Gateway' link at the bottom of the page.
       */
      mdPresortObj.SetMailersID("123456");

      /* Postage Payment Methods 
       * If both of these functions are set to false (Default setting), then Metered Mail will be used.
       * Permit Imprint - Mailing without affixing postage. When set to true, current mailing will use permit imprint as method of postage payment
       */
      // presortObj.PSPermitImprint = false; // way the mailer pays for his mailing

      /*	Precanceled Stamp - Cancellation of adhesive postage, stamped envelopes, or stamped cards before mailing.
       *	When set to true, PSPrecanceledStampValue must be set to the postage value in cents
       */
      //presortObj.PSPrecanceledStamp = true;
      //presortObj.PSPrecanceledStampValue(0.10);

      /* Sorting Options 
       * PSNCOA - Uses NCOALink for move updates
       */
      mdPresortObj.SetPSNCOA(true);

      /* Post Office of Mailing Information - This is the Post office from where it will be mailed from */
      mdPresortObj.SetPSPostOfficeOfMailingCity("RSM");
      mdPresortObj.SetPSPostOfficeOfMailingState("CA");
      mdPresortObj.SetPSPostOfficeOfMailingZIP("92688");

      /* Update the Parameters 
       * Verify and validate that the following properties were set to correct ranges: SetPieceHeight, SetPieceLength, SetPieceThickness,and SetPieceWeight
       */
      if (mdPresortObj.UpdateParameters() == false)
      {
        Console.WriteLine("Parameter Error: " + mdPresortObj.GetParametersErrorString());
        Console.WriteLine("\nAborting the program...\n");

        Environment.Exit(1);
      }

      /* Add Input Records to Presort Object
       * Parse through the input file and add each record to the Presort Object
       */
      StreamReader reader = null;
      Dictionary<string, string> addressDict = new Dictionary<string, string>();
 
      try
      {
        reader = new StreamReader(data.PresortFile);
        string line = reader.ReadLine(); //Header row
        line = reader.ReadLine();
        
        while (line != null)
        {
          // Pre-determined format for test file
          string[] split = line.Split(new char[] { ',' });
          mdPresortObj.SetRecordID(split[0]);
          string addressInfo = split[1] + "," + split[2] + "," + split[3] + "," + split[4] + "," + split[5] + "," + split[6] + "," + split[7] + "," + split[8] + "," + split[9];
          addressDict.Add(split[0], addressInfo);
          mdPresortObj.SetZip(split[6]);
          mdPresortObj.SetPlus4(split[7]);
          mdPresortObj.SetDeliveryPointCode(split[8]);
          mdPresortObj.SetCarrierRoute(split[9]);
       
          //Add record
          if (mdPresortObj.AddRecord() == false)
          {
            Console.WriteLine("Error Adding Record " + split[0] + ": " + mdPresortObj.GetParametersErrorString());
          }

          line = reader.ReadLine();
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("\nError Reading File: " + ex.Message);
        Console.WriteLine("\nAborting the program...\n");

        Environment.Exit(1);
      }
      finally
      {
        if (reader != null)
        {
          reader.Close();
        }
      }

      /* Do Presort - Do and check if Presort was successful, otherwise return the error.*/
      if (mdPresortObj.DoPresort() == false)
      {
        Console.WriteLine("Error During Presort: " + mdPresortObj.GetParametersErrorString());
        Console.WriteLine("\nAborting the program...\n");

        Environment.Exit(1);
      }

      /* Write Results */
      try
      {
        StreamWriter sw = new StreamWriter(data.OutputFile);

        // Write the headers on the first line
        sw.WriteLine("RecID,MAK,Address,Suite,City,State,Zip,Plus4,DeliveryPointCode,CarrierRoute,TrayNumber,SequenceNumber,Endorsement,BundleNumber");

        // Grab the first record and continue through the file
        bool writeback = mdPresortObj.GetFirstRecord();
        while (writeback)
        {
          sw.WriteLine(mdPresortObj.GetRecordID() + "," + addressDict[mdPresortObj.GetRecordID()] + ","
              + mdPresortObj.GetTrayNumber() + "," + mdPresortObj.GetSequenceNumber() + "," + mdPresortObj.GetEndorsementLine() + "," + mdPresortObj.GetBundleNumber());
          writeback = mdPresortObj.GetNextRecord();
        }

        sw.Flush();
        sw.Close();
      }
      catch (Exception ex)
      {
        Console.WriteLine("\nError Writing File: " + ex.Message);
      }
    }
  }
  public class DataContainer
  {
    public string PresortFile { get; set; } = "";
    public string OutputFile { get; set; } = "";

    public void FormatPresortOutputFile()
    { 
      // change the name of the output file
      int location = PresortFile.IndexOf(".csv");
      OutputFile = PresortFile.Substring(0, location) + "_output.csv";
    }

    public List<string> GetWrapped()
    {
      FileInfo file = new FileInfo(OutputFile);
      string filePath = file.FullName;

      int maxLineLength = 50;
      string[] lines = filePath.Split(new string[] { "\\" }, StringSplitOptions.None);
      string currentLine = "";
      List<string> wrappedString = new List<string>();

      foreach (string section in lines)
      {
        if ((currentLine + section).Length > maxLineLength)
        {
          wrappedString.Add(currentLine.Trim());
          currentLine = "";
        }

        if(section.Contains(OutputFile))
        {
          currentLine += section;
        }
        else
        {
          currentLine += section + "\\";
        }    
      }

      if (currentLine.Length > 0)
      {
        wrappedString.Add(currentLine.Trim());
      }

      return wrappedString;
    }
  }
}