using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
//using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.OleDb;
using Telerik.Windows.Controls;
using Telerik.Windows.Controls.Chart;
//using Telerik.Charting;
using Telerik.Windows.Data;
using Microsoft.Win32;
using Telerik.Windows.Controls.ChartView;
using System.IO;
using System.Drawing;
using Telerik.Windows.Controls.Charting;
using System.Windows.Controls;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Excel_Retriever
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<string> names = new List<string>();
        public List<string> AllNames = new List<string>();
        public List<BubblePoints> AllPoints;
        public DataSet beginDataSet, endDataSet, combinedDataSet;
        public DataSet originalBeginDataSet, originalEndDataSet;
        public DataSet beginDSBubble, endDSBubble;
        public int SubtractedNamesBegin, SubtractedNamesEnd, ColNumBegin, ColNumEnd, LeftToExportBegin, LeftToExportEnd, startGraphBegin, startGraphEnd, NumNamesRowBegin, NumNamesRowEnd;
        public bool onBeginSheet, importTwoGraphs = false, SecondChartExported = false, BoolDisableNameBox = false;
        public string textFilePath = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
        }

        public DataSet GetDataTableFromExcel(string FilePath)
        {
            try
            {
                OleDbConnection con = new OleDbConnection("Provider= Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + "; Extended Properties=\"Excel 12.0;HDR=YES;\"");
                OleDbDataAdapter da = new OleDbDataAdapter("select * from [Sheet1$]", con);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }

#region Export Bubble Graphs

        private void ExportBbubbleCharts_Click(object sender, RoutedEventArgs e)
        {
            if (importTwoGraphs)
            {
                Ignore_Names_2Sheets();
                onBeginSheet = true;
                GetBubblePoints(beginDSBubble);
                ExportBubbleGraph();
            }

            else
            {
                if (onBeginSheet)
                {
                    Ignore_Names(beginDSBubble);
                    GetBubblePoints(beginDSBubble);
                    ExportBubbleGraph();
                }

                else
                {
                    Ignore_Names(endDSBubble);
                    GetBubblePoints(endDSBubble);
                    ExportBubbleGraph();
                }
            }
            ExportBubbleText.Text = "Bubble Chart Successfully Exported!";
        }

        public void ExportSecondChart()
        {
            onBeginSheet = false;
            SecondChartExported = true;
            GetBubblePoints(endDSBubble);
            ExportBubbleGraph();
        }

        public void Ignore_Names(DataSet sheet)
        {
            int num, ColNum = onBeginSheet ? ColNumBegin : ColNumEnd;
            var table = sheet.Tables[0];
            var columns = table.Columns;
            var nameColumn = columns["Name"];

            foreach (DataRow row in nameColumn.Table.Rows)
                names.Add(row[0].ToString().ToLower());
            for (int i = ColNum; i > 0; i--)
            {
                string test = columns[i].ColumnName.ToLower();
                foreach (char c in test)
                {
                    if (int.TryParse(c.ToString(), out num))
                        test = test.Remove(test.IndexOf(c));
                }
                test = test.Trim();
                if (!names.Contains(test))
                {
                    if (!BoolDisableNameBox)
                    {
                        if (MessageBox.Show(test + " is not in the 'Name' column. Continue?", "Name does not exist.",
                            MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                            return;
                    }
                    columns.RemoveAt(i);
                    if (onBeginSheet)
                        SubtractedNamesBegin++;
                    else
                        SubtractedNamesEnd++;
                }
            }
            if (onBeginSheet)
            {
                SubtractedNamesBegin = Math.Abs(ColNum - SubtractedNamesBegin);
                for (int k = SubtractedNamesBegin; k > 0; k--)
                    sheet.Tables[0].Rows[k - 1][k] = 0;
            }
            else
            {
                SubtractedNamesEnd = Math.Abs(ColNum - SubtractedNamesEnd);
                for (int k = SubtractedNamesEnd; k > 0; k--)
                    sheet.Tables[0].Rows[k - 1][k] = 0;
            }
        }

        public void GetBubblePoints(DataSet sheet)
        {
            AllPoints = new List<BubblePoints>();

            int NumRows = onBeginSheet ? NumNamesRowBegin : NumNamesRowEnd;
            int SubtractedNames = onBeginSheet ? SubtractedNamesBegin : SubtractedNamesEnd;

            for (int i = 0; i <= 5; i++)
                for (int j = 0; j <= 5; j++)
                    AllPoints.Add(new BubblePoints(i, j));

            for (int i = 0; i < NumRows-1; i++)
            {
                for (int j = 1; j <= SubtractedNames; j++)
                {
                    if (i == (j - 1))
                        continue;
                    int x = int.Parse(sheet.Tables[0].Rows[i][j].ToString());
                    int y = int.Parse(sheet.Tables[0].Rows[j - 1][i + 1].ToString());
                    foreach (BubblePoints b in AllPoints)
                    {
                        if (b.X == x && b.Y == y)
                        {
                            b.total++;
                            break;
                        }
                    }
                }     
            }
        }

        public void Ignore_Names_2Sheets()
        {
            int num;
            List<string> BeginList = new List<string>(), EndList = new List<string>(), CombList = new List<string>();
            for (int i = 0; i < NumNamesRowBegin; i++)
                BeginList.Add(beginDSBubble.Tables[0].Rows[i]["Name"].ToString().ToLower().Trim());

            for (int j = 0; j < NumNamesRowEnd; j++)
                EndList.Add(endDSBubble.Tables[0].Rows[j]["Name"].ToString().ToLower().Trim());

            if (BeginList.Count <= EndList.Count)
            {
                for (int k = 0; k < BeginList.Count; k++)
                {
                    if (EndList.Contains(BeginList.ElementAt(k)))
                        CombList.Add(BeginList.ElementAt(k));
                }
            }

            else
            {
                for (int l = 0; l < EndList.Count; l++)
                {
                    if (BeginList.Contains(EndList.ElementAt(l)))
                        CombList.Add(EndList.ElementAt(l));
                }
            }

            for (int i = 0; i < NumNamesRowBegin; i++)
            {
                string name = beginDSBubble.Tables[0].Rows[i]["Name"].ToString().ToLower().Trim();
                if (!CombList.Contains(name))
                {
                    beginDSBubble.Tables[0].Rows.RemoveAt(i);
                    NumNamesRowBegin--;
                    i--;
                }
            }

            for (int i = 0; i < NumNamesRowEnd; i++)
            {
                string name = endDSBubble.Tables[0].Rows[i]["Name"].ToString().ToLower().Trim();
                if (!CombList.Contains(name))
                {
                    endDSBubble.Tables[0].Rows.RemoveAt(i);
                    NumNamesRowEnd--;
                    i--;
                }
            }

            for (int i = ColNumBegin; i > 0; i--)
            {
                var table = beginDSBubble.Tables[0];
                var columns = table.Columns;
                var nameColumn = columns["Name"];

                string test = columns[i].ColumnName.ToLower();
                foreach (char c in test)
                {
                    if (int.TryParse(c.ToString(), out num))
                        test = test.Remove(test.IndexOf(c));
                }
                test = test.Trim();
                if (!CombList.Contains(test))
                {
                    if (!BoolDisableNameBox)
                    {
                        if (MessageBox.Show(test + " is not in the 'Name' column. Continue?", "Name does not exist.",
                            MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                            return;
                    }
                    columns.RemoveAt(i);
                    SubtractedNamesBegin++;
                }
            }

            for (int i = ColNumEnd; i > 0; i--)
            {
                var table = endDSBubble.Tables[0];
                var columns = table.Columns;
                var nameColumn = columns["Name"];

                string test = columns[i].ColumnName.ToLower();
                foreach (char c in test)
                {
                    if (int.TryParse(c.ToString(), out num))
                        test = test.Remove(test.IndexOf(c));
                }
                test = test.Trim();
                if (!CombList.Contains(test))
                {
                    if (!BoolDisableNameBox)
                    {
                        if (MessageBox.Show(test + " is not in the 'Name' column. Continue?", "Name does not exist.",
                            MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                            return;
                    }
                    columns.RemoveAt(i);
                    SubtractedNamesEnd++;
                }
            }

            SubtractedNamesEnd = Math.Abs(ColNumEnd - SubtractedNamesEnd);
            for (int k = SubtractedNamesEnd; k > 0; k--)
                endDSBubble.Tables[0].Rows[k - 1][k] = 0;

            SubtractedNamesBegin = Math.Abs(ColNumBegin - SubtractedNamesBegin);
            for (int k = SubtractedNamesBegin; k > 0; k--)
                beginDSBubble.Tables[0].Rows[k - 1][k] = 0;
        }

        public void ExportBubbleGraph()
        {
            BubbleGraph.Width = 800;
            BubbleGraph.Height = 600;
            BubbleGraph.Measure(new System.Windows.Size(BubbleGraph.Width, BubbleGraph.Height));
            BubbleGraph.Arrange(new System.Windows.Rect(BubbleGraph.DesiredSize));
            BubbleGraph.DefaultView.ChartArea.EnableAnimations = false;

            BubbleGraph.DefaultView.ChartTitle.Content = onBeginSheet ? "Beginning of Semester Evaluation" : "End of Semester Evaluation";

            BubbleGraph.DefaultView.ChartTitle.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            BubbleGraph.DefaultView.ChartLegend.UseAutoGeneratedItems = true;
            DataSeries bubbleSeries = new DataSeries();
            bubbleSeries.LegendLabel = "Evaluation";
            bubbleSeries.Definition = new BubbleSeriesDefinition();

            foreach (BubblePoints b in AllPoints)
            {
                bubbleSeries.Add(new DataPoint() { XValue = b.X, YValue = b.Y, BubbleSize = b.total * .25 });
            }

            BubbleGraph.DefaultView.ChartArea.DataSeries.Add(bubbleSeries);

            Dispatcher.BeginInvoke((Action)(() =>
            {
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Bubble Graphs\";
                if (!Directory.Exists(filePath))
                    Directory.CreateDirectory(filePath);
                using (FileStream fs = File.Create(filePath + BubbleGraph.DefaultView.ChartTitle.Content.ToString() + ".png"))
                {
                    if (importTwoGraphs && !SecondChartExported)
                    {
                        ExportSecondChart();
                        BubbleGraph.ExportToImage(fs, new PngBitmapEncoder());
                    }
                    else
                        BubbleGraph.ExportToImage(fs, new PngBitmapEncoder());
                }
            }), DispatcherPriority.ApplicationIdle, null);

        }

#endregion Export Bubble Graphs

#region Export Bar Graphs
        private void ExportBarGraphs_Click(object sender, RoutedEventArgs e)
        {
            if (importTwoGraphs)
                onBeginSheet = true;
            ExportInitialGraph();
        }

        private void ExportInitialGraph()
        {
            ExportBtn.IsEnabled = false;
            graph.DefaultView.ChartArea.DataSeries.Clear();
            if (onBeginSheet)
                ExportChart(beginDataSet);
            else
                ExportChart(endDataSet);
        }

        private void ExportGraphs()
        {
            if (onBeginSheet)
            {
                LeftToExportBegin--;
                LeftToExportText.Text = LeftToExportBegin.ToString() + " Names Remaining";
                graph.DefaultView.ChartArea.DataSeries.Clear();
                if (onBeginSheet)
                    ExportChart(beginDataSet);
                else
                    ExportChart(endDataSet);
            }
            else
            {
                LeftToExportEnd--;
                LeftToExportText.Text = LeftToExportEnd.ToString() + " Names Remaining";
                graph.DefaultView.ChartArea.DataSeries.Clear();
                if (onBeginSheet)
                    ExportChart(beginDataSet);
                else
                    ExportChart(endDataSet);
            }
        }

        private void ExportChart(DataSet sheet)
        {
            graph.Width = 800;
            graph.Height = 600;
            graph.Measure(new System.Windows.Size(graph.Width, graph.Height));
            graph.Arrange(new System.Windows.Rect(graph.DesiredSize));
            graph.DefaultView.ChartArea.EnableAnimations = false;
            string semester = onBeginSheet ? " Beginning" : " End";
            int num;
            if (onBeginSheet)
            {
                string name = sheet.Tables[0].Columns[LeftToExportBegin].ColumnName.ToString();
                foreach (char c in name)
                {
                    if (int.TryParse(c.ToString(), out num))
                        name = name.Remove(name.IndexOf(c));
                }
                name = name.Trim();
                graph.DefaultView.ChartTitle.Content = name + semester + " of Semester Results";
            }
            else
            {
                string name = sheet.Tables[0].Columns[LeftToExportEnd].ColumnName.ToString();
                foreach (char c in name)
                {
                    if (int.TryParse(c.ToString(), out num))
                        name = name.Remove(name.IndexOf(c));
                }
                name = name.Trim();
                graph.DefaultView.ChartTitle.Content = name + semester + " of Semester Results";
            }
            graph.DefaultView.ChartTitle.HorizontalAlignment = HorizontalAlignment.Center;
            graph.DefaultView.ChartLegend.UseAutoGeneratedItems = true;
            DataSeries barSeries = new DataSeries();
            barSeries.LegendLabel = "Results";
            barSeries.Definition = new BarSeriesDefinition();

            string complete, very, moderate, somewhat, notatall, norating;
            if (onBeginSheet)
            {
                complete = sheet.Tables[0].Rows[startGraphBegin][LeftToExportBegin].ToString();
                very = sheet.Tables[0].Rows[startGraphBegin + 1][LeftToExportBegin].ToString();
                moderate = sheet.Tables[0].Rows[startGraphBegin + 2][LeftToExportBegin].ToString();
                somewhat = sheet.Tables[0].Rows[startGraphBegin + 3][LeftToExportBegin].ToString();
                notatall = sheet.Tables[0].Rows[startGraphBegin + 4][LeftToExportBegin].ToString();
                norating = sheet.Tables[0].Rows[startGraphBegin + 5][LeftToExportBegin].ToString();
            }
            else
            {
                complete = sheet.Tables[0].Rows[startGraphEnd][LeftToExportEnd].ToString();
                very = sheet.Tables[0].Rows[startGraphEnd + 1][LeftToExportEnd].ToString();
                moderate = sheet.Tables[0].Rows[startGraphEnd + 2][LeftToExportEnd].ToString();
                somewhat = sheet.Tables[0].Rows[startGraphEnd + 3][LeftToExportEnd].ToString();
                notatall = sheet.Tables[0].Rows[startGraphEnd + 4][LeftToExportEnd].ToString();
                norating = sheet.Tables[0].Rows[startGraphEnd + 5][LeftToExportEnd].ToString();
            }

            barSeries.Add(new DataPoint() { YValue = Convert.ToDouble(complete), XCategory = "Completely" });
            barSeries.Add(new DataPoint() { YValue = Convert.ToDouble(very), XCategory = "Very" });
            barSeries.Add(new DataPoint() { YValue = Convert.ToDouble(moderate), XCategory = "Moderately" });
            barSeries.Add(new DataPoint() { YValue = Convert.ToDouble(somewhat), XCategory = "Somewhat" });
            barSeries.Add(new DataPoint() { YValue = Convert.ToDouble(notatall), XCategory = "Not At All" });
            barSeries.Add(new DataPoint() { YValue = Convert.ToDouble(norating), XCategory = "No Rating" });
            graph.DefaultView.ChartArea.DataSeries.Add(barSeries);

            Dispatcher.BeginInvoke((Action)(() =>
            {
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Bar Charts\";
                if (!Directory.Exists(filePath))
                    Directory.CreateDirectory(filePath);
                using (FileStream fs = File.Create(filePath + graph.DefaultView.ChartTitle.Content.ToString() + ".png"))
                {
                    if (onBeginSheet)
                    {
                        graph.ExportToImage(fs, new PngBitmapEncoder());
                        if (LeftToExportBegin == 1)
                        {
                            ExportBtn.IsEnabled = false;
                            if (importTwoGraphs)
                            {
                                onBeginSheet = false;
                                ExportInitialGraph();
                            }
                            LeftToExportText.Text = "Graphs successfully exported!";
                        }
                        else
                            ExportGraphs();
                    }
                    else
                    {
                        graph.ExportToImage(fs, new PngBitmapEncoder());
                        if (LeftToExportEnd == 1)
                        {
                            ExportBtn.IsEnabled = false;
                            LeftToExportText.Text = "Graphs successfully exported!";
                        }
                        else
                            ExportGraphs();
                    }
                }
            }), DispatcherPriority.ApplicationIdle, null);
        }
#endregion Export Bar Graphs

#region Import Sheets
        private void ImportBeginSheet(object sender, RoutedEventArgs e)
        {
            ImportBeginSheet();
        }

        private void ImportBeginSheet()
        {
            onBeginSheet = true;
            if (!int.TryParse(CompletelyStartBegin.Text.ToString(), out startGraphBegin))
            {
                MessageBox.Show("Input valid 'Complete' row number.");
                return;
            }
            if (!int.TryParse(NumColumnsBegin.Text.ToString(), out ColNumBegin))
            {
                MessageBox.Show("Please enter a valid column number.");
                return;
            }
            if (!int.TryParse(NumNamesRowBeginBox.Text.ToString(), out NumNamesRowBegin))
            {
                MessageBox.Show("Please enter a valid row number.");
                return;
            }
            startGraphBegin -= 2;
            OpenFileDialog BeginSheet = new OpenFileDialog();
            BeginSheet.Multiselect = false;
            BeginSheet.DefaultExt = ".xlsx";
            BeginSheet.Filter = "Excel Spreadsheet (*.xlsx),(*.xls) |*.xlsx;*.xls";
            if (BeginSheet.ShowDialog() == true)
            {
                beginDataSet = GetDataTableFromExcel(BeginSheet.FileName);
                originalBeginDataSet = beginDataSet.Copy();
                beginDSBubble = beginDataSet.Copy();
                LeftToExportBegin = ColNumBegin;
                if (!importTwoGraphs)
                {
                    LeftToExportText.Text = LeftToExportBegin.ToString() + " Names Remaining";
                    ExportBtn.IsEnabled = CombineBtn.IsEnabled = ExportBubbleBtn.IsEnabled = true;
                }
                BeginImportText.Text = "Excel File Imported Successfully!";
            }
        }

        private void ImportEndSheet(object sender, RoutedEventArgs e)
        {
            ImportEndSheet();
        }

        private void ImportEndSheet()
        {
            onBeginSheet = false;
            if (!int.TryParse(CompletelyStartEnd.Text.ToString(), out startGraphEnd))
            {
                MessageBox.Show("Input valid 'Complete' row number.");
                return;
            }
            if (!int.TryParse(NumColumnsEnd.Text.ToString(), out ColNumEnd))
            {
                MessageBox.Show("Please enter a valid column number.");
                return;
            }
            if (!int.TryParse(NumNamesRowEndBox.Text.ToString(), out NumNamesRowEnd))
            {
                MessageBox.Show("Please enter a valid row number.");
                return;
            }
            startGraphEnd -= 2;
            OpenFileDialog EndSheet = new OpenFileDialog();
            EndSheet.Multiselect = false;
            EndSheet.DefaultExt = ".xlsx";
            EndSheet.Filter = "Excel Spreadsheet (*.xlsx),(*.xls) |*.xlsx;*.xls";
            if (EndSheet.ShowDialog() == true)
            {
                endDataSet = GetDataTableFromExcel(EndSheet.FileName);
                originalEndDataSet = endDataSet.Copy();
                endDSBubble = endDataSet.Copy();
                LeftToExportEnd = ColNumEnd;
                if (!importTwoGraphs)
                    LeftToExportText.Text = LeftToExportEnd.ToString() + " Names Remaining";
                ExportBtn.IsEnabled = CombineBtn.IsEnabled = ExportBubbleBtn.IsEnabled = true;
            }
            EndImportText.Text = "Excel File Imported Successfully!";
        }

        private void ImportTwoChecked(object sender, RoutedEventArgs e)
        {
            importTwoGraphs = true;
        }

        private void ImportTwoUnchecked(object sender, RoutedEventArgs e)
        {
            importTwoGraphs = false;
        }
        private void DisableNameChecked(object sender, RoutedEventArgs e)
        {
            BoolDisableNameBox = true;
        }

        private void DisableNameUnchecked(object sender, RoutedEventArgs e)
        {
            BoolDisableNameBox = false;
        }

#endregion Import Sheets

#region Combine Comments and Ratings
        private void CombineCommAndRatings_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveExcel = new SaveFileDialog();
            saveExcel.DefaultExt = ".xlsx";
            saveExcel.Filter = "Excel Worksheet (*.xlsx) |*.xlsx";
            if (importTwoGraphs)
            {
                CombineCommentsAndRatings2Charts();
                if (saveExcel.ShowDialog() == true)
                {
                    ExportDataTableToExcel(combinedDataSet.Tables[0], saveExcel.FileName);
                    CombineText.Text = "Excel File Exported Successfully!";
                }
            }
            else
            {
                if (onBeginSheet)
                {
                    CombineCommentsAndRatingsBegin();
                    if (saveExcel.ShowDialog() == true)
                    {
                        ExportDataTableToExcel(originalBeginDataSet.Tables[0], saveExcel.FileName);
                        CombineText.Text = "Excel File Exported Successfully!";
                    }
                }
                else
                {
                    CombineCommentsAndRatingsEnd();
                    if (saveExcel.ShowDialog() == true)
                    {
                        ExportDataTableToExcel(originalEndDataSet.Tables[0], saveExcel.FileName);
                        CombineText.Text = "Excel File Exported Successfully!";
                    }
                }
            }
        }

        public void CombineCommentsAndRatings2Charts()
        {
            InsertZerosAndComputeAverages(originalBeginDataSet, true);
            CompareWithTextFile(originalBeginDataSet);
            CombineCommentsAndRatingsBegin();
            combinedDataSet = originalBeginDataSet.Copy();
            InsertZerosAndComputeAverages(originalEndDataSet, false);
            CompareWithTextFile(originalEndDataSet);
            CombineCommentsAndRatingsEnd();
            foreach (DataColumn col in originalEndDataSet.Tables[0].Columns)
            {
                if (col.ColumnName == "Name")
                    continue;
                combinedDataSet.Tables[0].Columns.Add(col.ColumnName);
                int rowCount = 0;
                foreach (DataRow row in col.Table.Rows)
                {
                    string obj = row[col.ColumnName].ToString();
                    combinedDataSet.Tables[0].Rows[rowCount][col.ColumnName] = obj;
                    rowCount++;
                }
            }
            int num, j = 3;
            for (int i = ColNumBegin * 2 + 1; i < ColNumBegin * 2 + ColNumEnd * 2; i++) //begin names*2+1; begin*2 + end*2; ++
            {
                string name = combinedDataSet.Tables[0].Columns[i].ColumnName.ToString();
                foreach (char c in name)
                {
                    if (int.TryParse(c.ToString(), out num))
                        name = name.Remove(name.IndexOf(c));
                }
                name = name.Trim();
                if (TestName(name))
                {
                    combinedDataSet.Tables[0].Columns[i++].SetOrdinal(j++);
                    combinedDataSet.Tables[0].Columns[i].SetOrdinal(j++);
                    j += 2;
                }
                else
                    i++;
            }
            ChangeNameColumn();
            int CombinedRowCount = combinedDataSet.Tables[0].Rows.Count;
            combinedDataSet.Tables[0].Rows[CombinedRowCount - 1].Delete();
            combinedDataSet.Tables[0].Rows[CombinedRowCount - 2].Delete();
        }

        public void ChangeNameColumn()
        {
            for (int i = 0; i < AllNames.Count; i++)
                combinedDataSet.Tables[0].Rows[i][0] = AllNames.ElementAt(i);
        }

        public bool TestName(string name)
        {
            int num;
            for (int i = 1; i < originalBeginDataSet.Tables[0].Columns.Count; i += 2)
            {
                string col = originalBeginDataSet.Tables[0].Columns[i].ColumnName;
                foreach (char c in col)
                {
                    if (int.TryParse(c.ToString(), out num))
                        col = col.Remove(col.IndexOf(c));
                }
                col = col.Trim();
                if (name.Equals(col, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        public void CombineCommentsAndRatingsBegin()
        {
            InsertZerosAndComputeAverages(originalBeginDataSet, true);
            int j = 2;
            for (int i = 0; i < ColNumBegin; i++) //20 is number of names in columns
            {
                if (j % 2 == 0)
                    originalBeginDataSet.Tables[0].Columns[ColNumBegin + 1 + i].SetOrdinal(j); //21 is column where comments begin (number of names + 1)
                else
                {
                    originalBeginDataSet.Tables[0].Columns[ColNumBegin + 1 + i].SetOrdinal(j + 1);
                    j++;
                }
                j++;
            }
        }

        public void CombineCommentsAndRatingsEnd()
        {
            InsertZerosAndComputeAverages(originalEndDataSet, false);
            int j = 2;
            for (int i = 0; i < ColNumEnd; i++) //20 is number of names in columns
            {
                if (j % 2 == 0)
                    originalEndDataSet.Tables[0].Columns[ColNumEnd + 1 + i].SetOrdinal(j); //21 is column where comments begin (number of names + 1)
                else
                {
                    originalEndDataSet.Tables[0].Columns[ColNumEnd + 1 + i].SetOrdinal(j + 1);
                    j++;
                }
                j++;
            }
        }

        public void CompareWithTextFile(DataSet sheet)
        {
            if(string.IsNullOrEmpty(textFilePath))
            {
                OpenFileDialog TextFile = new OpenFileDialog();
                TextFile.Multiselect = false;
                TextFile.DefaultExt = ".txt";
                TextFile.Filter = "Text File (*.txt) |*.txt";
                if (TextFile.ShowDialog() == true)
                    textFilePath = TextFile.FileName;
            }

            AllNames = new List<string>();
            using (TextReader txt = new StreamReader(textFilePath))
            {
                string s = string.Empty;
                while ((s = txt.ReadLine()) != null)
                    AllNames.Add(s);
            }
            for (int i = 0; i < AllNames.Count; i++)
            {
                string s = sheet.Tables[0].Rows[i]["Name"].ToString();
                if (!AllNames.ElementAt(i).Equals(s, StringComparison.OrdinalIgnoreCase))
                    sheet.Tables[0].Rows.InsertAt(sheet.Tables[0].NewRow(), i);
            }
        }

        public void InsertZerosAndComputeAverages(DataSet sheet, bool beginSheet)
        {
            var table = sheet.Tables[0];
            var columns = table.Columns;
            var nameColumn = columns["Name"];
            int num;

            sheet.Tables[0].Rows.InsertAt(sheet.Tables[0].NewRow(), table.Rows.Count);
            nameColumn.Table.Rows[table.Rows.Count - 1][nameColumn] = "Average across rows";

            int row = beginSheet ? NumNamesRowBegin : NumNamesRowEnd;
            int col = beginSheet ? ColNumBegin : ColNumEnd;

            for (int i = 0; i < row; i++)
            {
                string name = nameColumn.Table.Rows[i][nameColumn].ToString();
                for (int j = 1; j <= col; j++)
                {
                    string header = sheet.Tables[0].Columns[j].ColumnName;
                    foreach (char c in header)
                    {
                        if (int.TryParse(c.ToString(), out num))
                            header = header.Remove(header.IndexOf(c));
                    }
                    header = header.Trim();
                    if (name.Equals(header, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.Tables[0].Rows[i][j] = "0";
                        table.Rows[table.Rows.Count - 1][j] = ComputeAveragesAcrossRow(sheet, i, col);
                        break;
                    }
                }
            }
            ComputeAveragesDownColumn(sheet, col);
        }

        public void ComputeAveragesDownColumn(DataSet sheet, int col)
        {
            sheet.Tables[0].Rows.InsertAt(sheet.Tables[0].NewRow(), sheet.Tables[0].Rows.Count);
            sheet.Tables[0].Columns[0].Table.Rows[sheet.Tables[0].Rows.Count - 1][sheet.Tables[0].Columns[0]] = "Average down columns";

            for (int i = 1; i <= col; i++)
            {
                double currentNum, sum = 0, totalNums = 0;
                for (int j = 0; j < (sheet.Tables[0].Rows.Count - 12); j++)
                {
                    if (double.TryParse(sheet.Tables[0].Rows[j][i].ToString(), out currentNum))
                    {
                        if (currentNum != 0)
                        {
                            sum += currentNum;
                            totalNums++;
                        }
                    }
                }
                sheet.Tables[0].Rows[sheet.Tables[0].Rows.Count - 1][i] = (sum / totalNums).ToString("N5");
            }
        }

        public string ComputeAveragesAcrossRow(DataSet sheet, int row, int col)
        {
            double currentNum, sum = 0, totalNums = 0;
            for (int i = 1; i <= col; i++)
            {
                if (double.TryParse(sheet.Tables[0].Rows[row][i].ToString(), out currentNum))
                {
                    if (currentNum != 0)
                    {
                        sum += currentNum;
                        totalNums++;
                    }
                }
            }
            return (sum / totalNums).ToString("N5");
        }

        public static bool ExportDataTableToExcel(DataTable dt, string filepath)
        {
            Excel.Application oXL;
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;

            try
            {
                // Start Excel and get Application object. 
                oXL = new Excel.Application();

                // Set some properties 
                oXL.Visible = true;
                oXL.DisplayAlerts = false;

                // Get a new workbook. 
                oWB = oXL.Workbooks.Add(Missing.Value);

                // Get the Active sheet 
                oSheet = (Excel.Worksheet)oWB.ActiveSheet;
                oSheet.Name = "Sheet1";

                int rowCount = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    rowCount += 1;
                    for (int i = 1; i < dt.Columns.Count + 1; i++)
                    {
                        // Add the header the first time through 
                        if (rowCount == 2)
                        {
                            oSheet.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                        }
                        oSheet.Cells[rowCount, i] = dr[i - 1].ToString();
                    }
                }

                // Save the sheet and close 
                oSheet = null;
                oWB.SaveAs(filepath, Excel.XlFileFormat.xlWorkbookDefault,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Excel.XlSaveAsAccessMode.xlExclusive,
                    Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value);
                oWB.Close(Missing.Value, Missing.Value, Missing.Value);
                oWB = null;
                oXL.Quit();
            }
            catch{ throw; }
            finally
            {
                // Clean up 
                // NOTE: When in release mode, this does the trick 
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            return true;
        }

#endregion Combine Comments and Ratings
    }

    public class BubblePoints
    {
        public int X, Y, total;

        public BubblePoints()
        { this.total = 0; }

        public BubblePoints(int x, int y)
        {
            this.X = x;
            this.Y = y;
            this.total = 0;
        }
    }
}