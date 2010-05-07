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
      bool purge = false;

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

        if (dictArgs.ContainsKey("purge"))
        {
          if (dictArgs["purge"].ToString().ToLower() == "y")
          {
            purge = true;
          }
          else
          {
            Console.Write("Are you sure you want to purge your items [y/n]?");
            ConsoleKeyInfo keyInfo = Console.ReadKey();
            if (keyInfo.Key == ConsoleKey.Y)
            {
              purge = true;
            }
            else
            {
              return;
            }
            Console.WriteLine();
          }
        }

        CountAllFileVersions(siteUrl,
            webUrl,
            docLibName,
            outFilename,
            outFileSepChar,
            purge);

        Console.WriteLine("Done.");
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
      }

    }

    private static void Usage()
    {
      Console.WriteLine("Usage:");
      Console.WriteLine("CountFileVersions.exe -url <website> [-web <webname>] -doclib <document library name> -outfile <filename> -sepchar <sepchar> [-purge [y]]");
      Console.WriteLine();
      Console.WriteLine("Examples:");
      Console.WriteLine();
      Console.WriteLine("CountFileVersions -url http://mysite -web subsite -doclib mylib");
      Console.WriteLine("\tDisplays all files and versions in the mylib library.");
      Console.WriteLine();
      Console.WriteLine("CountFileVersions -url http://mysite -doclib mylib -purge y");
      Console.WriteLine("\tDisplays all files and versions in the mylib library.");
      Console.WriteLine("\tAnd purges all versions.");
    }

    private static void CountAllFileVersions(
      string siteUrl,
      string webUrl,
      string docLibName,
      string outFilename,
      string outFileSepChar,
      bool purge)
    {
      int totalFileCount = 0;
      int totalCount = 0;
      int purgedCount = 0;
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
          SPList list = null;
          try
          {
            list = web.Lists[docLibName];
          }
          catch (Exception ex)
          {
            Console.WriteLine("Can't open the list " + docLibName);
            Console.WriteLine("Reason: " + ex.Message);
            return;
          }

          if (purge)
          {
            Console.WriteLine("Scanning and purging...");
          }
          else
          {
            Console.WriteLine("Scanning...");
          }
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
              if (purge && item.File.Versions.Count > 0)
              {
                purgedCount += item.File.Versions.Count;
                item.File.Versions.DeleteAll();
              }

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
      Console.WriteLine("Total number of versions purged: {0}", purgedCount);
    }
  }
}
