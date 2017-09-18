using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;

namespace DWRCalReport
{
    class Program
    {
        private string filePath;// = @"c:\tools\dwrlist.csv";
        static void Main(string[] args)
        {
            Program p = new Program();
            if (args.Length > 0)
            {
                p.filePath = args[0];
            }
            else
            {
                p.filePath = Properties.Settings.Default.filePath;
            }
            p.GetAllCalendarItems();
            Console.WriteLine("done");
            Console.ReadLine();
        }

        public void WriteDWRItem(DWRItemClass item)
        {
            Console.WriteLine(String.Format("{0}", item.Categories));
            Console.WriteLine(String.Format("{0}", item.Subject));
            Console.WriteLine(String.Format("{0}", item.startDate));
            Console.WriteLine(String.Format("{0}", item.endDate));
            Console.WriteLine(String.Format("{0}", item.duration));
            Console.WriteLine(String.Format("{0}", item.isAllDay));
        }

        public void WriteAllItems(List<DWRItemClass> items)
        {
            //items.Sort(delegate (DWRItemClass o1, DWRItemClass o2) { return o1.startDate.CompareTo(o2.startDate); });
            foreach (DWRItemClass item in items)
            {
                WriteDWRItem(item);
            }
        }

        public void WriteToCSV(List<DWRItemClass> items, string filePath)
        {
            List<string> catList = new List<string>();

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (var file = File.CreateText(filePath))
            {
                file.WriteLine("StartDate,Subject,Categories,EndDate,Duration,AllDay");
                foreach (var item in items)
                {
                    var cList = item.Categories.Split(',');
                    catList = new List<string>(cList);
                    catList.RemoveAll(delegate (string s1) 
                    {
                        return !s1.Trim().StartsWith("DWR:");
                    });
                    //catList.Sort(delegate (string o1, string o2)
                    //    {
                    //        if (o1.StartsWith("DWR:") && o2.StartsWith("DWR:")) return 0;
                    //        if (o1.StartsWith("DWR:") && !o2.StartsWith("DWR:")) return 0;
                    //        else return 1;
                    //    });
                    var cats = String.Join(",", catList.ToArray());
                    file.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\"", item.startDate, item.Subject, cats, item.endDate, item.duration, item.isAllDay));
                }
                file.WriteLine();
            }
        }

        public void GetAllCalendarItems()
        {
            List<DWRItemClass> dwrItems = new List<DWRItemClass>();

            Microsoft.Office.Interop.Outlook.Application oApp = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNamespace = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder CalendarFolder = null;
            Microsoft.Office.Interop.Outlook.Items outlookCalendarItems = null;

            oApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = oApp.GetNamespace("MAPI"); ;
            CalendarFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);
            outlookCalendarItems = CalendarFolder.Items;
            outlookCalendarItems.IncludeRecurrences = true;

            foreach (Microsoft.Office.Interop.Outlook.AppointmentItem item in outlookCalendarItems)
            {
                if (item.IsRecurring)
                {
                    Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                    DateTime first = new DateTime(2017, 9, 17, item.Start.Hour, item.Start.Minute, 0);
                    DateTime last = new DateTime(2017, 9, 17);
                    Microsoft.Office.Interop.Outlook.AppointmentItem recur = null;

                    for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                    {
                        try
                        {
                            recur = rp.GetOccurrence(cur);
                            //MessageBox.Show(recur.Subject + " -> " + cur.ToLongDateString());
                            //Console.WriteLine(recur.Subject + " -> " + cur.ToLongDateString());
                        }
                        catch
                        { }
                    }
                }
                else
                {
                    if (item.Start >= DateTime.Now)
                    {
                        //MessageBox.Show(item.Subject + " -> " + item.Start.ToLongDateString());
                        //Console.WriteLine(item.Subject + " -> " + item.Start.ToLongDateString());
                        if (item.Categories.Contains("DWR:"))
                        {
                            dwrItems.Add(new DWRItemClass() { Categories = item.Categories, duration = item.Duration, endDate = item.End, isAllDay = item.AllDayEvent, startDate = item.Start, Subject = item.Subject });
                        }
                    }
                }
            }

            dwrItems.Sort(delegate (DWRItemClass o1, DWRItemClass o2) { return o1.startDate.CompareTo(o2.startDate); });
            WriteAllItems(dwrItems);
            WriteToCSV(dwrItems, filePath);
        }
    }

    public class DWRItemClass
    {
        public string Categories { get; set; }
        public string Subject { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public int duration { get; set; }
        public bool isAllDay { get; set; }
    }
}
