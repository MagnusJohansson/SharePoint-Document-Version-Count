using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.SharePoint;
using Utility;
using System.IO;

namespace CountFileVersions
{
  public class Program
  {
    public static void Main(string[] args)
    {
      string siteUrl = string.Empty;
      string webUrl = string.Empty;
      string docLibName = string.Empty;
      string outFilename = string.Empty;
      string outFileSepChar = "\t";

      try
      {
        CommandArgs cmdArgs = CommandLine.Parse(args);
        Dictionary<string, string> dictArgs = cmdArgs.ArgPairs;

        if (dictArgs.ContainsKey("url"))
        {
          siteUrl = dictArgs["url"];
        }
        else
        {
          Usage();
          return;
        }

        if (dictArgs.ContainsKey("web"))
        {
          webUrl = dictArgs["web"];
        }

        if (dictArgs.ContainsKey("doclib"))
        {
          docLibName = dictArgs["doclib"];
        }
        else
        {
          Usage();
          return;
        }

        if (dictArgs.ContainsKey("outfile"))
        {
          outFilename = dictArgs["outfile"];
        }

        if (dictArgs.ContainsKey("sepchar"))
        {
          outFileSepChar = dictArgs["sepchar"];
        }


        CountAllFileVersions(siteUrl,
            webUrl,
            docLibName,
            outFilename,
            outFileSepChar);

        Console.WriteLine("Press return to continue.");
        Console.ReadLine();
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
      }

    }

    private static void Usage()
    {
      Console.WriteLine("Usage:");
      Console.WriteLine("CountFileVersions.exe -url <website> <-web <webname>> -doclib <document library name> -outfile <filename> -sepchar <sepchar>");
    }

    private static void CountAllFileVersions(
      string siteUrl,
      string webUrl,
      string docLibName,
      string outFilename,
      string outFileSepChar)
    {
      int totalFileCount = 0;
      int totalCount = 0;
      bool writeToFile = outFilename.Length > 0;
      StreamWriter sw = null;

      if (writeToFile)
      {
        sw = new StreamWriter(outFilename);
      }

      using (SPSite site = new SPSite(siteUrl))
      {
        using (SPWeb web = site.OpenWeb(webUrl))
        {
          SPList list = web.Lists[docLibName];

          foreach (SPListItem item in list.Items)
          {
            if (item.File != null)
            {
              Console.WriteLine("Filename: {0}, version count: {1}", item.File.Name, item.File.Versions.Count + 1);
              if (writeToFile)
              {
                sw.WriteLine("{0}{1}{2}", item.File.Name, outFileSepChar, item.File.Versions.Count + 1);
              }
              totalCount += item.File.Versions.Count + 1;

              totalFileCount++;
            }
          }
        }
      }

      if (sw != null)
      {
        sw.Close();
        sw.Dispose();
      }

      Console.WriteLine("Total number of files: {0}", totalFileCount);
      Console.WriteLine("Total number of versions: {0}", totalCount);
    }
  }
}
