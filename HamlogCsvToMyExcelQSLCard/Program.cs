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
            var Worksheet = new MyLiblary.ExcelControl();
            var AllData = new List<LogData>();
            const string PATHNAME = "160814";
            using (var TextFile = new System.IO.StreamReader($@"C:\Users\Daiti Murota\Documents\QSLカード\CSV\{PATHNAME}\{PATHNAME}.csv", Encoding.Default))
            {
                while (TextFile.EndOfStream == false)
                {
                    AllData.Add(new LogData(TextFile.ReadLine()));
                }
            }
            //AllData = AllData.OrderBy(x => x.Callsign.Substring(2, 1)).ThenBy(x => x.Callsign.Substring(1, 1)).ThenBy(x => x.Callsign.Substring(3)).Select(x => x).ToList();
            //AllData.ForEach(x => Console.WriteLine(x.Callsign));
            using (var ExitCsvStream = new System.IO.StreamWriter($@"C:\Users\Daiti Murota\Documents\QSLカード\CSV\{PATHNAME}\{PATHNAME}new.csv", false, Encoding.Default))
            {
                AllData.ForEach(x => x.WriteToCsv(ExitCsvStream));
                ExitCsvStream.Close();
            }
            foreach (LogData Data in AllData)
            {
                Data.WriteToExcel(Worksheet);
            }
            AllData = AllData.OrderBy(x => x.Callsign.Substring(2, 1)).ThenBy(x => x.Callsign.Substring(1, 1)).ThenBy(x => x.Callsign.Substring(3)).Select(x => x).ToList();
            using (var ExitOrderCallsign = new System.IO.StreamWriter($@"C:\Users\Daiti Murota\Documents\QSLカード\CSV\{PATHNAME}\{PATHNAME}Order.txt", false, Encoding.Default))
            {
                AllData.ForEach(x => ExitOrderCallsign.WriteLine(x.Callsign));
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
        public string MovingArea { get; private set; } = "";
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
            if (SeparetedData[12].IndexOf("MyOut", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var Output = SeparetedData[12].Split(' ').Where(x => x.IndexOf("MyOut", StringComparison.OrdinalIgnoreCase) >= 0)
                    .Select(x => x).First().Split(':')[1];
                Output = Output.Replace("\"", "");
                this.Output = Output;
            }
            else
            {
                switch (this.Mode)
                {
                    case "JT65":
                        this.Output = "10W";
                        break;
                    case "FM":
                        this.Output = "20W";
                        break;
                    case "SSB":
                        this.Output = "50W";
                        break;
                }
            }
            if (SeparetedData[12].IndexOf("Moving", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (SeparetedData[12].IndexOf("：") >= 0) { SeparetedData[12] = SeparetedData[12].Replace("：", ":"); }
                var MovingArea = SeparetedData[12].Split(' ').Where(x => x.IndexOf("Moving", StringComparison.OrdinalIgnoreCase) >= 0)
                    .Select(x => x).First().Split(':')[1];
                MovingArea = MovingArea.Replace("\"", "");
                this.MovingArea = MovingArea;
            }
            this.IsGetQSL = (SeparetedData[9].IndexOf('*') >= 0)
                ? true : false;
            this.IsContest = (SeparetedData[12].IndexOf("59MUROTA") >= 0)
                ? true : false;
        }
        public void WriteToExcel(MyLiblary.ExcelControl WorkSheet)
        {
            this.WriteCallSign(WorkSheet);
            WorkSheet.WriteCell("A9", this.Year.Substring(2));
            WorkSheet.WriteCell("C9", this.Month);
            WorkSheet.WriteCell("E9", this.Day);
            WorkSheet.WriteCell("F9", this.Jst);
            WorkSheet.WriteCell("I9", this.SignalReport);
            WorkSheet.WriteCell("K9", this.Frequency);
            WorkSheet.WriteCell("O9", this.Mode);
            WorkSheet.WriteCell("A14", $"【OUTPUT】 {this.Output}");
            /*if (this.IsGetQSL == true)
            {
                WorkSheet.WriteCell("O10", "TNX");
            }
            else
            {*/
            WorkSheet.WriteCell("O10", "PSE");
            //}
            string Remarks = "【Remarks】 ";
            if (this.IsContest == true)
            {
                Remarks += "2016 NYP,TNX FB QSO!";
            }
            if (this.MovingArea != "")
            {
                Remarks += $"貴局移動地:{this.MovingArea}";
            }
            if (Remarks == "【Remarks】 ")
            { Remarks += "TNX FB QSO!"; }
            WorkSheet.WriteCell("A15", Remarks);
            WorkSheet.PrintSheet(true);
        }
        private void WriteCallSign(MyLiblary.ExcelControl WorkSheet)
        {
            WorkSheet.WriteCell("A3", this.Callsign);
            char[] CallsignSepareted = this.Callsign.ToCharArray();
            WorkSheet.WriteCell("G2", CallsignSepareted[0].ToString());
            WorkSheet.WriteCell("I2", CallsignSepareted[1].ToString());
            WorkSheet.WriteCell("K2", CallsignSepareted[2].ToString());
            WorkSheet.WriteCell("M2", CallsignSepareted[3].ToString());
            WorkSheet.WriteCell("O2", CallsignSepareted[4].ToString());
            WorkSheet.WriteCell("Q2", CallsignSepareted[5].ToString());
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
}
