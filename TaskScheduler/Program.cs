using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskScheduler
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("<<< TASK SCHEDULER >>>");

            while (true)
            {
                Console.WriteLine("\ntype 'exit' to exit");
                Console.WriteLine("type 'add' to add task");
                string input = Console.ReadLine().ToLower();

                if (input == "add")
                {
                    Add();
                }
                else if (input == "exit")
                {
                    Console.WriteLine("Exit the app? y - yes / n - no:");
                    string exit = Console.ReadLine().ToLower();
                    if (exit == "y")
                    {
                        break;
                    }
                }
            }
        }
        static void Add()
        {
            // Get the service on the remote machine
            using (TaskService ts = new TaskService())
            {
                TaskDefinition td = ts.NewTask();
                
                // Description
                Console.WriteLine("\nType description:");
                string description = Console.ReadLine();
                td.RegistrationInfo.Description = description;

                // Add Trigger
                Trigger taskTrigger = GetTriggerFromConsole();
                td.Triggers.Add(taskTrigger);

                // Create an action that will be launched
                Console.WriteLine("\nType aplication path:");
                string path = Console.ReadLine();
                td.Actions.Add(new ExecAction(path));

                // Register the task in the root folder
                Console.WriteLine("\nType name:");
                string name = Console.ReadLine();
                ts.RootFolder.RegisterTaskDefinition(name, td);
            }
        }
        static Trigger GetTriggerFromConsole()
        {
            // Trigger hours
            Console.WriteLine("Type trigger hours:");
            int.TryParse(Console.ReadLine(), out int hours);
            hours = hours > 23 ? 23 : (hours < 0 ? 0 : hours);
            
            // Trigger minutes
            Console.WriteLine("Type trigger minutes:");
            int.TryParse(Console.ReadLine(), out int minutes);
            minutes = minutes > 59 ? 59 : (minutes < 0 ? 0 : minutes);

            // Trigger type
            while (true)
            {
                Console.WriteLine("Trigger type:\nd - daily, w - weekly, m - monthly");
                string input = Console.ReadLine().ToLower();

                if (input == "d") // Daily
                {
                    // Interval
                    Console.WriteLine("Interval between days in days:");
                    short.TryParse(Console.ReadLine(), out short interval);
                    interval = interval < 0 ? (short)0 : interval;

                    // Return trigger
                    return new DailyTrigger
                    {
                        DaysInterval = interval,
                        StartBoundary = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, hours, minutes, 0)
                    };
                }
                else if (input == "w") // Weekly
                {
                    // Interval
                    Console.WriteLine("Interval between weeks in weeks:");
                    short.TryParse(Console.ReadLine(), out short interval);
                    interval = interval < 0 ? (short)0 : interval;

                    // Day of week
                    Console.WriteLine("Which day in week?\n0 - all days, 1 - Monday, 2 - Tuesday, ..., 7 - Sunday:");
                    int.TryParse(Console.ReadLine(), out int weekDay);
                    weekDay = weekDay > 7 ? 7 : (weekDay < 0 ? 0 : weekDay);

                    DaysOfTheWeek[] daysOfTheWeek = new DaysOfTheWeek[]
                    {
                            DaysOfTheWeek.AllDays, DaysOfTheWeek.Monday, DaysOfTheWeek.Tuesday,
                            DaysOfTheWeek.Wednesday, DaysOfTheWeek.Thursday, DaysOfTheWeek.Friday,
                            DaysOfTheWeek.Saturday, DaysOfTheWeek.Sunday
                    };

                    // Return trigger
                    return new WeeklyTrigger
                    {
                        WeeksInterval = interval,
                        DaysOfWeek = daysOfTheWeek[weekDay],
                        StartBoundary = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, hours, minutes, 0)
                    };
                }
                else if (input == "m") // Monthly
                {
                    // Month of year
                    Console.WriteLine("Which month in year?\n0 - All months, 1 - January, 2 - February, ..., 12 - December:");
                    int.TryParse(Console.ReadLine(), out int yearMonth);
                    yearMonth = yearMonth > 12 ? 12 : (yearMonth < 0 ? 0 : yearMonth);

                    MonthsOfTheYear[] monthsOfTheYear = new MonthsOfTheYear[]
                    {
                            MonthsOfTheYear.AllMonths,
                            MonthsOfTheYear.January, MonthsOfTheYear.February, MonthsOfTheYear.March, MonthsOfTheYear.April,
                            MonthsOfTheYear.May, MonthsOfTheYear.June, MonthsOfTheYear.July, MonthsOfTheYear.August,
                            MonthsOfTheYear.September, MonthsOfTheYear.October, MonthsOfTheYear.November, MonthsOfTheYear.December
                    };

                    Console.WriteLine("Which days in month?\n type numbers between 1 - 31\ntype 'last' to include last day of month\ntype 'end' to stop typing:");
                    List<int> daysOfMonth = new List<int>();
                    bool lastDay = false;
                    while (true)
                    {
                        string console = Console.ReadLine().ToLower();
                        int.TryParse(console, out int monthDay);
                        if (monthDay >= 1 && monthDay <= 31)
                            daysOfMonth.Add(monthDay);
                        if (console == "last")
                            lastDay = true;
                        if (console == "end")
                            break;
                    }

                    // Return trigger
                    return new MonthlyTrigger
                    {
                        DaysOfMonth = daysOfMonth.ToArray(),
                        MonthsOfYear = monthsOfTheYear[yearMonth],
                        RunOnLastDayOfMonth = lastDay,
                        StartBoundary = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, hours, minutes, 0)
                    };
                }
            }
        }
    }
}
