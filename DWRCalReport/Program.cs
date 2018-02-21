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
            Console.WriteLine(String.Format("{0}", item.DWRDays));
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
                file.WriteLine("Project,DWRActivity,Categories,Subject,StartDate,EndDate,DWRDays,AllDay,Duration");
                foreach (var item in items)
                {
                    //Get Project info as well from raw object
                    //DWRDetails details = ParseDetails(item);
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
                    file.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\"", 
                        item.Project, item.DWRActivity, cats, item.Subject, item.startDate, item.endDate, item.DWRDays, item.isAllDay, item.duration));
                }
                file.WriteLine();
            }
        }

        private DWRDetails ParseDetails(DWRItemClass item)
        {
            DWRDetails details = new DWRDetails();

            var clist = item.Categories.Split(',').ToList<string>();

            List<string> catList = clist.FindAll(delegate (string s1) 
            {
                return s1.Trim().StartsWith("DWR:");
            });

            List<string> projectInfo = clist.FindAll(delegate (string s1)
            {
                return s1.Trim().StartsWith("Project:");
            });

            if (projectInfo != null && projectInfo.Count == 1)
            {
                details.project = projectInfo[0].Trim();
                details.DWRProcess = String.Join(",", catList.ToArray());
                //var pandc = projectInfo[0].Trim().Split('-');
                //if (pandc.Length == 2)
                //{
                //    details.DWRProcess = String.Join(",", catList.ToArray());
                //    details.Company = pandc[0];
                //    details.project = pandc[1];
                //}
            }
            else
            {
                details.DWRProcess = String.Join(",", catList.ToArray());
                details.project = "Unspecified";
            }

            return details;
        }

        private Decimal GetMinDWRDays(DateTime startDate, DateTime endDate)
        {
            if (startDate == null || endDate == null) { return 0.0M; }  //Defensive

            Decimal _mMinTime = Properties.Settings.Default.mintime;

            TimeSpan total = endDate - startDate;
            Decimal ts = total.Days;
            if (Decimal.Compare(ts, _mMinTime) < 0) { ts = _mMinTime; }

            return ts;
        }

        public void GetAllCalendarItems()
        {
            List<DWRItemClass> dwrItems = new List<DWRItemClass>();
            DateTime startDate;

            try
            {
                string startDateText = Properties.Settings.Default.startDate;
                string[] di = startDateText.Split('/');
                int year = Convert.ToInt32(di[2]);
                int month = Convert.ToInt32(di[1]);
                int day = Convert.ToInt32(di[0]);
                startDate = new DateTime(year, month, day);
            }
            catch
            {
                startDate = DateTime.Now;
            }

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
                    if (item.Start >= startDate)
                    {
                        //MessageBox.Show(item.Subject + " -> " + item.Start.ToLongDateString());
                        //Console.WriteLine(item.Subject + " -> " + item.Start.ToLongDateString());
                        if (item.Categories != null && item.Categories.Trim().Contains("DWR:"))
                        {
                            //dwrItems.Add(new DWRItemClass() { Categories = item.Categories, duration = item.Duration, endDate = item.End, isAllDay = item.AllDayEvent, startDate = item.Start,
                            //Subject = item.Subject, DWRDays = GetMinDWRDays(item.Start, item.End) });
                            dwrItems.Add(ParseItemToObject(item));
                        }
                    }
                }
            }

            dwrItems.Sort(delegate (DWRItemClass o1, DWRItemClass o2) { return o1.startDate.CompareTo(o2.startDate); });
            WriteAllItems(dwrItems);
            WriteToCSV(dwrItems, filePath);
        }

        private DWRItemClass ParseItemToObject(Microsoft.Office.Interop.Outlook.AppointmentItem item)
        {
            DWRItemClass di = new DWRItemClass();
            di.Categories = item.Categories;
            di.duration = item.Duration;
            di.endDate = item.End;
            di.isAllDay = item.AllDayEvent;
            di.startDate = item.Start;
            di.Subject = item.Subject;
            di.DWRDays = GetMinDWRDays(item.Start, item.End);

            DWRDetails details = ParseDetails(di);
            di.Project = details.project;
            di.DWRActivity = details.DWRProcess;

            return di;
        }
    }

    public class DWRItemClass
    {
        public string Categories { get; set; }
        public string Project { get; set; }
        public string DWRActivity { get; set; }
        public string Subject { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public int duration { get; set; }
        public bool isAllDay { get; set; }
        public Decimal DWRDays { get; set; }
    }

    public class DWRDetails
    {
        public string DWRProcess { get; set; }
        public string Company { get; set; }
        public string project { get; set; }
    }
}
