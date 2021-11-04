using CsvFileManage;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BarcodeMatchCsv
{
    static class Global
    {

        const int ciBarcodeColIdx = 23;
        const int ciEndingColIdx = 47;
        const int ciHandleColIdx = 0;
        const int ciTitleColIdx = 1;
        const int ciBodyColIdx = 2;
        const int ciAppendedColIdxPackingCsv = 11;

        static Dictionary<string, CsvRow> ProductsDictionaryByBarcode = new Dictionary<string, CsvRow>();
        static Dictionary<string, CsvRow> ProductsDictionaryByBody = new Dictionary<string, CsvRow>();
        static Dictionary<string, CsvRow> ProductsDictionaryByTitle = new Dictionary<string, CsvRow>();
        static Dictionary<string, CsvRow> ProductsDictionaryByHandle = new Dictionary<string, CsvRow>();

        static List<CsvRow> PackingList = new List<CsvRow>();
        static CsvRow rowResultHeader = new CsvRow();
        static public void main()
        {
            var dlg = new OpenFileDialog();
            dlg.InitialDirectory = Application.StartupPath;
            dlg.Title = "Select First csv File (it have to contain bacode and html)";
            dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            dlg.Multiselect = false;
            if (dlg.ShowDialog() != DialogResult.OK)                return;

            string strCsvFileForProducts = dlg.FileName;
            if (strCsvFileForProducts.Substring(strCsvFileForProducts.Length - 4).ToLower() != ".csv")
            {
                if( MessageBox.Show(null, "this is not csv file, continue?", "check", MessageBoxButtons.OKCancel )!=DialogResult.OK )
                {
                    return;
                }
            }

            dlg = new OpenFileDialog();
            dlg.InitialDirectory = Application.StartupPath;
            dlg.Title = "Select Second csv File (it have to conttain the products)";
            dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            dlg.Multiselect = false;
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string strCsvFileForPackingList = dlg.FileName;
            if (strCsvFileForPackingList.Substring(strCsvFileForPackingList.Length - 4).ToLower() != ".csv")
            {
                if (MessageBox.Show(null, "this is not csv file, continue?", "check", MessageBoxButtons.OKCancel) != DialogResult.OK)
                {
                    return;
                }
            }

            ReadingCsvFileToDictionary(strCsvFileForProducts);
            MakingPackingListFromCsvFile(strCsvFileForPackingList);


            var dlg1 = new SaveFileDialog();
            dlg1.InitialDirectory = Application.StartupPath;
            dlg1.Title = "Select result csv File.";
            dlg1.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            if (dlg1.ShowDialog() != DialogResult.OK) return;

            string strCsvFileForResults = dlg1.FileName;
            Processing_PackingList_From_ProductsDictionary(strCsvFileForResults);

        }

        public static void Processing_PackingList_From_ProductsDictionary(string sCsvSavingPath)
        {
            for(int i=0; i<PackingList.Count; i++)
            {
                CsvRow row = PackingList[i];
                string body = StripHTMLandSpace( row[ciBodyColIdx] );
                string title = row[ciTitleColIdx].Trim();
                string handle = row[ciHandleColIdx].Trim();

                string searchKey = handle.ToLower().Trim().Replace(" ", "-") + "_" + title.ToLower().Trim() + "_" + body.ToLower().Trim();
                CsvRow products_row = null;

                //foreach (KeyValuePair<string, CsvRow> entry in ProductsDictionaryByBarcode)
                //{
                //    string barcode = entry.Key;
                //    CsvRow rowProducts = (CsvRow)entry.Value;
                //    if (
                //        handle.ToLower().Replace(" ", "-").Contains(rowProducts[ciHandleColIdx].ToLower()) &&
                //        body.ToLower().Contains(rowProducts[ciBodyColIdx].ToLower()) &&
                //        title.ToLower().Contains(rowProducts[ciTitleColIdx].ToLower()))
                //        PackingList[i].Add(barcode);
                //}


                if (ProductsDictionaryByHandle.ContainsKey(searchKey))
                {
                    products_row = ProductsDictionaryByHandle[searchKey];
                    string barcode = products_row[ciBarcodeColIdx];
                    //PackingList[i].Add(barcode);
                    try
                    {
                        PackingList[i][ciAppendedColIdxPackingCsv] = barcode.Replace("'","").Trim();
                    }
                    catch
                    {

                    }
                }
            }

            Stream stream_w = null;
            try
            {
                stream_w = new FileStream(sCsvSavingPath, FileMode.Create, FileAccess.Write, FileShare.Write, 4096, true);
            }
            catch
            {
            }
            if(stream_w==null)
            {
                MessageBox.Show("Result file open failed");
                Application.Exit();
            }
            CsvWriter csvWriteHandler = new CsvWriter(stream_w);

            csvWriteHandler.WriteRow(rowResultHeader);
            foreach(CsvRow row in PackingList)
            {
                csvWriteHandler.WriteRow(row);
            }
            csvWriteHandler.Close();
            MessageBox.Show("Processing has been completed");
        }

        static string StripHTMLandSpace(string input)
        {
            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex("[ ]{2,}", options);
            input = regex.Replace(input, " ");
            return Regex.Replace(input, "<.*?>", String.Empty);
        }
        public static void ReadingCsvFileToDictionary(string csv_path)
        {

            int len = csv_path.Length;
            
            {
                String cvsFilePath = csv_path;
                Stream streamCSV = null;
                try
                {
                    streamCSV = new FileStream(cvsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
                }
                catch
                {
                    MessageBox.Show("Could not open the Products csv file.");
                    Application.Exit();
                }

                bool ret = false;
                CsvRow row = new CsvRow();
                CsvReader csvStream = new CsvReader(streamCSV);
                int iCount = 0;
                while ((ret = csvStream.ReadRow(row,2)))
                {
                    //if(!(ret = csvStream.ReadRow(row)))
                    //{
                    //    break;
                    //}
                    //while(row.Count< ciEndingColIdx+1)  //contains 0d in cell
                    //{
                    //    CsvRow tmpRow = new CsvRow();
                    //    if (!(ret = csvStream.ReadRow(tmpRow))) break;

                    //    int i = 0;
                    //    foreach(string item in tmpRow)
                    //    {
                    //        if (i++ == 0) row[row.Count - 1] = row[row.Count - 1] + item;
                    //        else
                    //            row.Add(item);
                    //    }
                    //}
                    //if (!ret) break;
                    row.LineText = "";
                    if (iCount++ == 0)  continue;;                    // header row
                    if (row[0].Trim().Length == 0) continue;        // empty row
                    if (Regex.IsMatch(row[0], @"^\d+$")) continue;  //colA is numeric condition
                    //if(iCount==137)
                    //{
                    //    System.Diagnostics.Debug.WriteLine(iCount);
                    //}
                    string sBarcodeKey = row[ciBarcodeColIdx].Replace("'", "").Replace(" ", "");
                    string sBodyKey = StripHTMLandSpace(row[ciBodyColIdx]);
                    string sTitleKey = row[ciTitleColIdx].Trim();
                    string sHandleKey = row[ciHandleColIdx].Trim();

                    string searchKey = sHandleKey.ToLower().Trim().Replace(" ", "") + "_" + sTitleKey.ToLower().Trim() + "_" + sBodyKey.ToLower().Trim();
                    CsvRow rowItem = new CsvRow();
                    foreach (string newitem in row) rowItem.Add(newitem);

                    if (!ProductsDictionaryByBarcode.ContainsKey( sBarcodeKey))
                        ProductsDictionaryByBarcode.Add(sBarcodeKey, rowItem);

                    if(!ProductsDictionaryByHandle.ContainsKey(searchKey))
                        ProductsDictionaryByHandle.Add(searchKey, rowItem);



                    //ProductsDictionaryByBody.Add(sBodyKey+"_"+ sTitleKey, row);        
                    //ProductsDictionaryByTitle.Add(sTitleKey, row);
                    //in some rows noth of BodyKey and TitleKey was same simulateniously.
                }
            }
        }
        public static void MakingPackingListFromCsvFile(string csv_path)
        {
            PackingList.Clear();
            int len = csv_path.Length;
            
            {
                String cvsFilePath = csv_path;

                Stream streamCSV = null;
                try
                {
                    streamCSV = new FileStream(cvsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
                }
                catch
                {
                    MessageBox.Show("Could not open the Packing List csv file.");
                    Application.Exit();
                }
                bool ret = false;
                CsvRow row = new CsvRow();
                CsvReader csvStream = new CsvReader(streamCSV);
                int iCount = 0;
                while ((ret = csvStream.ReadRow(row, 1)))
                {
                    if (!Regex.IsMatch(row[0], @"^\d+$") && iCount == 0)
                    {
                        foreach (string item in row) rowResultHeader.Add(item);
                        // rowResultHeader.Add("barcode");
                        rowResultHeader[ciAppendedColIdxPackingCsv] = "barcode";
                    }
                    if (iCount++ == 0) continue;                    // header row
                    if (row[0].Trim().Length == 0) continue;        // empty row
                    if (Regex.IsMatch(row[0], @"^\d+$")) continue;  //colA is numeric condition

                    string sBodyKey = StripHTMLandSpace(row[ciBodyColIdx]);
                    string sTitleKey = row[ciTitleColIdx].Trim();

                    CsvRow tmp = new CsvRow();
                    foreach (string item in row) tmp.Add(item);
                    PackingList.Add(tmp);
                }
            }
        }
    }
}
