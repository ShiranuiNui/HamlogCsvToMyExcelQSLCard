using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MyLiblary = ShiranuiSayaka.Libraries;

namespace HamlogCsvToMyExcelQSLCard
{
    class Program
    {
        static void Main(string[] args)
        {
            var Worksheet = new ExcelControlWithPrint();
            var AllData = new List<LogData>();
            using (var TextFile = new System.IO.StreamReader(@"C:\Users\Daiti Murota\Documents\Visual Studio 2015\Projects\HamlogCsvToMyExcelQSLCard\QSL-0729.csv", Encoding.Default))
            {
                while (TextFile.EndOfStream == false)
                {
                    AllData.Add(new LogData(TextFile.ReadLine()));
                }
            }
            AllData = AllData.OrderBy(x => x.Callsign.Substring(2, 1)).ThenBy(x => x.Callsign.Substring(1, 1)).ThenBy(x => x.Callsign.Substring(3)).Select(x => x).ToList();
            AllData.ForEach(x => Console.WriteLine(x.Callsign));
            var ExitCsvStream = new System.IO.StreamWriter(@"C:\Users\Daiti Murota\Documents\Visual Studio 2015\Projects\HamlogCsvToMyExcelQSLCard\QSL-0729New.csv",false,Encoding.Default);
            AllData.ForEach(x => x.WriteToCsv(ExitCsvStream));
            ExitCsvStream.Close();
            foreach (LogData Data in AllData)
            {
                Data.WriteToExcel(Worksheet);
            }
        }
    }
    class LogData
    {
        public string Callsign { get; private set; }
        public string Year { get; private set; }
        public string Month { get; private set; }
        public string Day { get; private set; }
        public string Jst { get; private set; }
        public string SignalReport { get; private set; }
        public string Frequency { get; private set; }
        public string Mode { get; private set; }
        public string Output { get; private set; }
        public bool IsGetQSL { get; private set; }
        public bool IsContest { get; private set; }
        private string OldTextData { get; }
        public LogData(string Data)
        {
            this.OldTextData = Data;
            string[] SeparetedData = Data.Split(',');
            this.Callsign = SeparetedData[0];
            string[] YearMonthDay = SeparetedData[1].Split('/');
            this.Year = "20" + YearMonthDay[0];
            this.Month = YearMonthDay[1];
            this.Day = YearMonthDay[2];
            this.Jst = SeparetedData[2].Replace("J", "");
            this.SignalReport = SeparetedData[3];
            this.Frequency = SeparetedData[5];
            this.Mode = SeparetedData[6];
            if (SeparetedData[13].IndexOf("MyOut") >= 0)
            {
                this.Output = "5W";
            }
            else
            {
                switch (this.Mode)
                {
                    case "JT65":
                        this.Output = "10W";
                        break;
                    case "FM":
                        this.Output = "18W";
                        break;
                    case "SSB":
                        this.Output = "20W";
                        break;
                }
            }
            this.IsGetQSL = (SeparetedData[9].IndexOf('*') >= 0)
                ? true : false;
            this.IsContest = (SeparetedData[12].IndexOf("59MUROTA") >= 0)
                ? true : false;
        }
        public void WriteToExcel(ExcelControlWithPrint WorkSheet)
        {
            WorkSheet.WriteCell("B1", this.Callsign);
            WorkSheet.WriteCell("A7", this.Year);
            WorkSheet.WriteCell("B7", this.Month);
            WorkSheet.WriteCell("C7", this.Day);
            WorkSheet.WriteCell("D7", this.Jst);
            WorkSheet.WriteCell("E7", this.SignalReport);
            WorkSheet.WriteCell("G7", this.Frequency);
            WorkSheet.WriteCell("G7", this.Mode);
            WorkSheet.WriteCell("C11", this.Output);
            if (this.IsGetQSL == true)
            {
                WorkSheet.WriteCell("F8", "TNX");
            }
            else
            {
                WorkSheet.WriteCell("F8", "PSE");
            }
            if (this.IsContest == true)
            {
                WorkSheet.WriteCell("B13", "2016 NYP,TXH FB QSO!");
            }
            else
            {
                WorkSheet.WriteCell("B13", "TNX FB QSO!");
            }
            WorkSheet.PreviewAndPrint();
        }
        public void WriteToCsv(System.IO.StreamWriter TargetFile)
        {
            string[] SeparetedData = this.OldTextData.Split(',');
            string OldSentData = SeparetedData[9];
            if (this.IsGetQSL == false)
            {
                SeparetedData[9] = "\"JS\"";
            }
            else
            {
                SeparetedData[9] = OldSentData.Replace(" ", "S");
            }
            string NewTextData = string.Join(",", SeparetedData);
            TargetFile.WriteLine(NewTextData);
        }
    }
    class ExcelControlWithPrint : MyLiblary.ExcelControl
    {
        public void PreviewAndPrint()
        {
            this.WorkSheet.PrintOutEx(1, 1, 1, true);
        }
    }
}
