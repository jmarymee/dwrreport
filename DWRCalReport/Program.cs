using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DWRCalReport
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
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
