
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Text;

namespace Course_Schedule
{
    class Program
    {
        static void Main()
        {
            Excel excel = new Excel(@"C:\Users\clair\Downloads\View_My_Courses.xlsx");

            excel.CreateCalender();
            Console.WriteLine("Done! :D");
        }
    }

    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ScheduleWS;
        int colour = 35;
        Dictionary<string, int> courseColours = new Dictionary<string, int>();

        public Excel(string path)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ScheduleWS = wb.Worksheets[1];

            wb.Worksheets.Add(After: ScheduleWS, Count: 2);

            CalenderSetup();

        }

        public void CreateCalender()
        {
            //Console.WriteLine("Create Calender");

            //wb.Worksheets[1].Cells[1, 1].Interior.Color = colour;
            //Console.WriteLine(ScheduleWS.Cells[1, 1].Text);

            for (int row = 4; ScheduleWS.Cells[row, 11].Value2 != null; row++)
            {
                List<int> terms = new List<int>();

                //StringBuilder startDate = ScheduleWS.Cells[row, 11].Value2;

                //Console.WriteLine(ScheduleWS.Cells[row, 11].text);

                if (ScheduleWS.Cells[row, 11].Text.Contains("2024"))
                    terms.Add(2);

                if (ScheduleWS.Cells[row, 12].Text.Contains("2025"))
                    terms.Add(3);

                TermCalender(row, terms);
            }

            Save();

        }

        public void TermCalender(int row, List<int> terms)
        {
            //Console.WriteLine("Term Calender");

            StringBuilder meetingPatterns = new StringBuilder(ScheduleWS.Cells[row, 8].Text);
            //Console.WriteLine(meetingPatterns);

            string courseSection = ScheduleWS.Cells[row, 5].Text;

            if (!meetingPatterns.Equals(""))
            {
                List<string> days = FindDays(meetingPatterns);
                List<double> times = FindTimes(meetingPatterns);
                int currentColour = colour;

                double timeStart = times[0] * 2;
                double timeEnd = times[1] * 2;

                string courseName = GetCourseName(courseSection);

                if (courseColours.ContainsKey(courseName))
                    currentColour = courseColours[courseName];
                else
                {
                    courseColours.Add(courseName, currentColour);
                    colour++;
                }

                for (int j = 0; j < terms.Count; j++)
                {
                    Worksheet termSheet = wb.Worksheets[terms[j]];

                    for (int i = 0; i < days.Count; i++)
                    {
                        int calenderColumn = CalenderColumn(days[i]);
                        termSheet.Cells[timeStart, calenderColumn + 1].Value2 = courseSection;
                        termSheet.Cells[timeStart + 1, calenderColumn + 1].Value2 = meetingPatterns.ToString();

                        for (double k = timeStart; k < timeEnd; k++)
                        {
                            if (termSheet.Cells[k, calenderColumn + 1].Interior.ColorIndex != -4142)
                                termSheet.Cells[k, calenderColumn + 1].Interior.ColorIndex = 3;

                            else 
                                termSheet.Cells[k, calenderColumn + 1].Interior.ColorIndex = currentColour;
                        }
                    }
                }
            }

            else
            {
                for (int j = 0; j < terms.Count; j++)
                {
                    Worksheet termSheet = wb.Worksheets[terms[j]];
                    int blankRow = 16;

                    while (termSheet.Cells[blankRow, 8].Value2 != null)
                    {
                        blankRow++;
                    }

                    termSheet.Cells[blankRow, 8] = courseSection;
                }
            }
        }

        void CalenderSetup()
        {
            //int MORNING_COLOUR = 19;
            //int AFTERNOON_COLOUR = 15;

            int TIME_LABEL_COLUMN_WIDTH = 10;

            Worksheet term1 = wb.Worksheets[2];
            Worksheet term2 = wb.Worksheets[3];

            List<int> timeLabel = new List<int> { 8, 0, 0 };
            StringBuilder timeOfDay = new StringBuilder("a.m.");

            //int timeColour = MORNING_COLOUR;

            for (int i = 1; i <= 24; i++)
            {
                if (timeLabel[0] == 12)
                {
                    timeOfDay[0] = 'p';
                    //timeColour = AFTERNOON_COLOUR;
                }

                term1.Cells[i + 15, 1].Value2 = timeLabel[0] + ":" + timeLabel[1] + timeLabel[2] + " " + timeOfDay;
                term2.Cells[i + 15, 1].Value2 = timeLabel[0] + ":" + timeLabel[1] + timeLabel[2] + " " + timeOfDay;

                //term1.Cells[i+15, 1].Interior.ColorIndex = timeColour;
                //term2.Cells[i + 15, 1].Interior.ColorIndex = timeColour;

                term1.Cells[i + 15, 1].ColumnWidth = TIME_LABEL_COLUMN_WIDTH;
                term2.Cells[i + 15, 1].ColumnWidth = TIME_LABEL_COLUMN_WIDTH;

                if (timeLabel[1] == 3)
                {
                    if (timeLabel[0] != 12)
                    {
                        timeLabel[0] = timeLabel[0] + 1;
                        timeLabel[1] = 0;
                    }

                    else
                    {
                        timeLabel[0] = 1;
                        timeLabel[1] = 0;
                    }
                }

                else
                {
                    timeLabel[1] = 3;

                    term1.Range[term1.Cells[i + 15, 1], term1.Cells[i + 15, 6]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    term1.Range[term1.Cells[i + 15, 1], term1.Cells[i + 15, 6]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

                    term2.Range[term2.Cells[i + 15, 1], term2.Cells[i + 15, 6]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    term2.Range[term2.Cells[i + 15, 1], term2.Cells[i + 15, 6]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
                    
                }

                
            }


            for (int i = 2; i <= 3; i++)
            {
                wb.Worksheets[i].Cells[15, 2] = "Mon";
                wb.Worksheets[i].Cells[15, 3] = "Tue";
                wb.Worksheets[i].Cells[15, 4] = "Wed";
                wb.Worksheets[i].Cells[15, 5] = "Thu";
                wb.Worksheets[i].Cells[15, 6] = "Fri";

                wb.Worksheets[i].Cells[15, 8] = "Asyncronous/Unscheduled Sections:";

                wb.Worksheets[i].Range["B1:F1"].ColumnWidth = 25;
                wb.Worksheets[i].Range["A15:A39"].RowHeight = 25;
                wb.Worksheets[i].Range["A1:F39"].WrapText = true;

                wb.Worksheets[i].Range["A1:A39"].VerticalAlignment = XlVAlign.xlVAlignTop;
            }

            term1.Name = "Term 1";
            term2.Name = "Term 2";

        }

        List<string> FindDays(StringBuilder meetingPatterns)
        {

            bool dayFound = false;

            for (int i = 0; !dayFound;)
            {
                //Console.WriteLine("Find Days:" + meetingPatterns);
                //Console.WriteLine("Find Days:" + meetingPatterns[0]);
                if (!meetingPatterns[i].Equals('|'))
                {
                    meetingPatterns.Remove(i, 1);

                }

                else
                {
                    meetingPatterns.Remove(i, 2);
                    dayFound = true;
                }
            }

            List<string> days = new List<string>();

            while (!meetingPatterns[0].Equals('|'))
            {
                StringBuilder day = new StringBuilder();

                while (!meetingPatterns[0].Equals(' '))
                {
                    day.Append(meetingPatterns[0]);
                    meetingPatterns.Remove(0, 1);
                }

                days.Add(day.ToString());
                meetingPatterns.Remove(0, 1);

            }

            meetingPatterns.Remove(0, 2);
            return days;
        }

        List<double> FindTimes(StringBuilder meetingPatterns)
        {
            //Console.WriteLine("Find Times");

            List<double> times = new List<double>();
            StringBuilder hour = new StringBuilder();
            double foundTime;
            int timesFound = 0;

            while (timesFound < 2)
            {
                if (!meetingPatterns[0].Equals(':'))
                {
                    hour.Append(meetingPatterns[0]);
                }

                else
                {
                    foundTime = double.Parse(hour.ToString());
                    meetingPatterns.Remove(0, 1);

                    if (meetingPatterns[0].Equals('3'))
                    {
                        foundTime += 0.5;
                    }

                    meetingPatterns.Remove(0, 3);

                    if (meetingPatterns[0].Equals('p') && foundTime - 12 < 0)
                    {
                        foundTime += 12;
                    }

                    times.Add(foundTime);
                    hour.Clear();
                    meetingPatterns.Remove(0, 6);

                    timesFound++;
                }

                meetingPatterns.Remove(0, 1);
            }

            return times;
        }

        int CalenderColumn(string day)
        {

            if (day.Equals("Mon"))
                return 1;

            else if (day.Equals("Tue"))
                return 2;

            else if (day.Equals("Wed"))
                return 3;

            else if (day.Equals("Thu"))
                return 4;

            else
                return 5;
        }

        string GetCourseName(string courseSection)
        {
            StringBuilder courseName = new StringBuilder();

            for (int i = 0; !courseSection[i].Equals('-'); i++)
            {
                courseName.Append(courseSection[i]);
            }

            return courseName.ToString();
        }

        public void Save()
        {
            wb.SaveAs(@"C:\Users\clair\Downloads\View_My_Courses_Calendar.xlsx");
            wb.Close();
            excel.Quit();
        }
    }
}
