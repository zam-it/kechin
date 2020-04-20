using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.IO.Compression;
using System.Diagnostics;

namespace ContinentalParser
{
    class Program
    {

        /// <summary>
        /// 
        /// </summary>
        static string[] headers = { "НАИМЕНОВАН", "ЦЕНА1", "ЦЕНА5", "ЦЕНАОТ5", "ПОРОГ1", "ПОРОГ5", "ЕДИЗМ1", "ЕДИЗМ5", "ЕДИЗМОТ5", "ДЛИНА", "ВЕС", "МАРКА", "ПРОИЗВЕЛ", "СКЛАД", "ВСЕГО", "ЦЕНА", "ЕДИЗМ", "ДАТА", "НОМЕР", "РЕЗЕРВ", "ЕДЗМВС" };
        /// <summary>
        /// 
        /// </summary>
        static string pathFile = "";
        /// <summary>
        /// 
        /// </summary>
        static string addrPrice = "";
        /// <summary>
        /// 
        /// </summary>
        static int rowcount = 1;
        /// <summary>
        /// 
        /// </summary>
        static void Main(string[] args)
        {


            try
            {
                Logger.Message("-----" + DateTime.Now.ToString() + "----------------------------------------");
                Logger.Message("Запуск парсера");
                //
                Logger.Message("Читаем конфигурацию..");
                addrPrice = ConfigurationManager.AppSettings["addrprice"];
                pathFile = ConfigurationManager.AppSettings["pathfile"];
                //
                Logger.Message("Проверка доступности CSV-каталога: \n" + Path.GetDirectoryName(pathFile));
                if (!Directory.Exists(Path.GetDirectoryName(pathFile)))
                {
                    Logger.Message("CSV-директория не существует. Укажите другой путь в файле конфигурации");
                }
                //
                Logger.Message("Скачиваем xls-прайс " + addrPrice);
                string pricebody = WebRequestClient();
                Logger.Message("xls-прайс удачно закачан");
                //
                Logger.Message("Удаляем временные файлы");
                if (File.Exists(pathFile + ".xls")) File.Delete(pathFile + ".xls");
                if (File.Exists(pathFile + ".temp")) File.Delete(pathFile + ".temp");
                //распаковываем
                MemoryStream readme = new MemoryStream(System.Text.Encoding.Default.GetBytes(pricebody));
                ZipStorer zip = ZipStorer.Open(readme, FileAccess.Read);
                List<ZipStorer.ZipFileEntry> dir = zip.ReadCentralDir();
                foreach (ZipStorer.ZipFileEntry entry in dir)
                {
                    if (Path.GetFileName(entry.FilenameInZip) == "price.xls")
                    {
                        Logger.Message("Создаем новый временный файл");
                        zip.ExtractFile(entry, pathFile + ".xls");
                        break;
                    }
                }
                zip.Close();
                //
                /*
                FileStream fs = File.Create(pathFile + ".xls");
                StreamWriter writer = new StreamWriter(fs, System.Text.Encoding.GetEncoding("Windows-1251"));
                writer.Write(pricebody);
                writer.Close();
                writer.Dispose();
                fs.Close();
                fs.Dispose();
                */
                //
                Logger.Message("Подключаем xml-файл в OleDb ");
                string ConnectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";Data Source={0}", pathFile + ".xls");
                //
                OleDbConnection cn = new OleDbConnection(ConnectionString);
                cn.Open();
                // Получаем список листов в файле
                DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                // Берем название первого листа

                Logger.Message("Выгружаем таблицу в файл: \n" + pathFile);
                FileStream fs = File.Create(pathFile + ".temp");
                StreamWriter writer = new StreamWriter(fs, System.Text.Encoding.GetEncoding("Windows-1251"));
                //Записываем заголовок в CSV
                string str_header = "";
                for (int col = 0; col < headers.Length; col++)
                {
                    str_header += "\"" + headers[col] + "\"";
                    if (col < headers.Length - 1) str_header += '\t';
                }
                writer.WriteLine(str_header);

                
                for (int ilst = 0; ilst < schemaTable.Rows.Count; ilst++)
                {

                    string sheetname = (string)schemaTable.Rows[ilst].ItemArray[2];
                    if (sheetname.IndexOf("Print_Area") != -1) continue;
                    if (sheetname.IndexOf("Область_печати") != -1) continue;
                    // Выбираем все данные с листа
                    string select = String.Format("SELECT * FROM [{0}]", sheetname);
                    DataSet ds = new DataSet();
                    OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
                    ad.Fill(ds);
                    cn.Close();
                    DataTable tb = ds.Tables[0];

                    if (sheetname.IndexOf("Лист") != -1)                    Sheet_Лист(tb, writer);
                    else if (sheetname.IndexOf("Труба 9941-81 (тн)") != -1) Sheet_Труба9941_81тн(tb, writer);
                    else if (sheetname.IndexOf("Трубы 9940, 14162, ASTM") != -1) Sheet_Трубы9940_14162_ASTM(tb, writer);
                    else if (sheetname.IndexOf("Пров-ка, электрод") != -1) Sheet_Шестигр_Пров_ка_электрод(tb, writer);

                    else if (sheetname.IndexOf("Тройники переходы") != -1) Sheet_Тройники_переходы(tb, writer);

                    else if (sheetname.IndexOf("Круг") != -1) Sheet_Круг(tb, writer);

                    else if (sheetname.IndexOf("отводы AISI") != -1) Sheet_отводы_AISI(tb, writer);

                    else if (sheetname.IndexOf("Отводы фланцы заглушки") != -1) Sheet_Отводы_фланцы_заглушки(tb, writer);

                    
                    

                    #region A
                 
                    #endregion

                }
                //
                writer.Close();
                writer.Dispose();
                fs.Close();
                fs.Dispose();
                //
                //заменяем старый файл
                if (File.Exists(pathFile)) File.Delete(pathFile);
                File.Move(pathFile + ".temp", pathFile);
                //
                Logger.Message("Удаляем временные файлы");
                File.Delete(pathFile + ".temp");
                if (File.Exists(pathFile + ".xls")) File.Delete(pathFile + ".xls");
                //
                Logger.Message("Парсер удачно завершен");

                Console.WriteLine(Process.GetCurrentProcess().ProcessName + " - OK");
                return;
            }
            catch (Exception ex)
            {
                Logger.Message("Произошла ошибка");
                Logger.Message(ex.Message + ex.Source);
            }
            Console.WriteLine(Process.GetCurrentProcess().ProcessName + " - ERROR");
        }
        /// <summary>
        /// </summary>
        public static void Sheet_Лист(DataTable tb, StreamWriter writer)
        {
            string gost = "";
            string fpname = "";

            
            


            for (int icol = 0; icol < 7; icol += 6)
            {
                for (int irow = 6; irow < tb.Rows.Count; irow++)
                {
                    DataRow row = tb.Rows[irow];
                    //
                    string tmpstr = row.ItemArray[icol].ToString();
                    //пропускаем записи, где цена не указана
                    decimal Cost1 = 0;
                    decimal.TryParse(row.ItemArray[3 + icol].ToString(), out Cost1);
                    //
                    if ((tmpstr.IndexOf("Х17") >= 0) || (tmpstr.IndexOf("Х18") >= 0) || (tmpstr.IndexOf("Х23") >= 0))
                    {
                        fpname = "ЛИСТ " + row.ItemArray[icol].ToString();
                    }
                    else if (Cost1 > 0)
                    {
                        if (tmpstr.Trim() != "")
                            gost = tmpstr;
                        else gost = "";

                        string Name = fpname + " " + gost + " мм";
                        string Measure = "т";
                        
                        decimal Cost5 = 0;
                        decimal CostOt5 = 0;
                        decimal.TryParse(row.ItemArray[2 + icol].ToString(), out Cost5);
                        decimal.TryParse(row.ItemArray[1 + icol].ToString(), out CostOt5);

                        string Porog1 = "1";
                        string Porog5 = "5";

                        Name = Name.Replace("   ", " ");
                        Name = Name.Replace("  ", " ");
                        Name = Name.Replace("   ", " ");
                        Name = Name.Replace("  ", " ");
                        Name = Name.Trim();
                        //записываем строку в CSV-файл
                        RowCsv csvrow = new RowCsv();
                        // "НАИМЕНОВАН", 
                        csvrow.Add(RowCsv.Type.String, Name);
                        //"ЦЕНА1", 
                        csvrow.Add(RowCsv.Type.Numeric, Cost1.ToString());
                        //"ЦЕНА5", 
                        csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                        //"ЦЕНАОТ5", 
                        csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                        //"ПОРОГ1", 
                        csvrow.Add(RowCsv.Type.Numeric, Porog1);
                        //"ПОРОГ5", 
                        csvrow.Add(RowCsv.Type.Numeric, Porog5);
                        //"ЕДИЗМ1", 
                        csvrow.Add(RowCsv.Type.String, Measure);
                        //"ЕДИЗМ5", 
                        csvrow.Add(RowCsv.Type.String, Measure);
                        //"ЕДИЗМОТ5", 
                        csvrow.Add(RowCsv.Type.String, Measure);
                        //"ДЛИНА", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"ВЕС", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"МАРКА", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"ПРОИЗВЕЛ", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"СКЛАД", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"ВСЕГО", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"ЦЕНА", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"ЕДИЗМ", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"ДАТА", 
                        csvrow.Add(RowCsv.Type.Null);
                        //"НОМЕР", 
                        csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                        //"РЕЗЕРВ"
                        csvrow.Add(RowCsv.Type.Null);
                        //
                        writer.WriteLine(csvrow.GetRow());
                        rowcount++;



                    }
                }
            }

        }
        /// <summary>
        /// </summary>
        public static void Sheet_Труба9941_81тн(DataTable tb, StreamWriter writer)
        {
            string fpname = "";

            int sp = GetBeginStr(tb, 1, "D", 12, "S");
            if (sp > 0)
            {
                for (int icol = 1; icol < 10; icol += 9)
                {
                    for (int irow = sp; irow < tb.Rows.Count; irow++)
                    {
                        DataRow row = tb.Rows[irow];

                        string tmpstr = row.ItemArray[icol].ToString();
                        decimal Cost1 = 0;
                        decimal.TryParse(row.ItemArray[6 + icol].ToString(), out Cost1);

                        //пропускаем записи, где цена не указана

                        if (tmpstr.IndexOf("Х18") >= 0)
                        {
                            fpname = tmpstr;
                        }
                        else if (Cost1 > 0)
                        {
                            string Name = "Труба "+ fpname + " " + tmpstr + "х" + row.ItemArray[2 + icol].ToString() + " мм";
                            
                            decimal Cost5 = 0;
                            decimal CostOt5 = 0;
                            decimal.TryParse(row.ItemArray[5 + icol].ToString(), out Cost5);
                            decimal.TryParse(row.ItemArray[4 + icol].ToString(), out CostOt5);

                            string Porog1 = "1";
                            string Porog5 = "5";
                            string Measure = "т";

                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Trim();
                            //записываем строку в CSV-файл
                            RowCsv csvrow = new RowCsv();
                            // "НАИМЕНОВАН", 
                            csvrow.Add(RowCsv.Type.String, Name);
                            //"ЦЕНА1", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost1.ToString());
                            //"ЦЕНА5", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                            //"ЦЕНАОТ5", 
                            csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                            //"ПОРОГ1", 
                            csvrow.Add(RowCsv.Type.Numeric, Porog1);
                            //"ПОРОГ5", 
                            csvrow.Add(RowCsv.Type.Numeric, Porog5);
                            //"ЕДИЗМ1", 
                            csvrow.Add(RowCsv.Type.String, Measure);
                            //"ЕДИЗМ5", 
                            csvrow.Add(RowCsv.Type.String, Measure);
                            //"ЕДИЗМОТ5", 
                            csvrow.Add(RowCsv.Type.String, Measure);
                            //"ДЛИНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВЕС", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"МАРКА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПРОИЗВЕЛ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"СКЛАД", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВСЕГО", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЦЕНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДАТА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"НОМЕР", 
                            csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                            //"РЕЗЕРВ"
                            csvrow.Add(RowCsv.Type.Null);
                            //
                            writer.WriteLine(csvrow.GetRow());
                            rowcount++;



                        }
                    }
                }
            }
        }
        /// <summary>
        /// </summary>
        public static void Sheet_Трубы9940_14162_ASTM(DataTable tb, StreamWriter writer)
        {
            string fpname = "";
            string pr = "";

            
            

            int sp = GetBeginStr(tb, 0, "D", 13, "S");
            if (sp > 0)
            {
                for (int icol = 0; icol < 11; icol += 11)
                {
                    for (int irow = sp; irow < tb.Rows.Count; irow++)
                    {
                        DataRow row = tb.Rows[irow];

                        string tmpstr = row.ItemArray[icol].ToString();
                        pr= row.ItemArray[3 + icol].ToString().Trim(); //D
				        if(pr=="") pr=row.ItemArray[5 + icol].ToString().Trim(); //F
				        
                        decimal Cost1 = 0;
                        decimal.TryParse(row.ItemArray[7 + icol].ToString(), out Cost1);


                        //пропускаем записи, где цена не указана

                        if (tmpstr.IndexOf("Х18") >= 0 && tmpstr.IndexOf("Труб") >= 0)
                        {
                            fpname = tmpstr;
                        }
                        else if (Cost1 > 0)
                        {
                            string Name = fpname +" "+ row.ItemArray[icol].ToString() + "x" + row.ItemArray[2+icol].ToString()+ " мм";
                            string Porog1 = "";
                            string Porog5 = "";
                            string Measure1 = "";
                            string Measure5 = "";
                            string MeasureOt5 = "";

                            decimal Cost5 = 0;
                            decimal CostOt5 = 0;
                            if (pr.IndexOf("т") >= 0)
                            {
                                
                                decimal.TryParse(row.ItemArray[6 + icol].ToString(), out Cost5);
                                decimal.TryParse(row.ItemArray[5 + icol].ToString(), out CostOt5);

                                Measure1 = "т";
                                Measure5 = "т";
                                MeasureOt5 = "т";
                                Porog1 = "1";
                                Porog5 = "5";
                            }
                            else 
                            {
                                Cost5 = 0;
                                CostOt5 = 0;
                                Measure1 = "м";
                            }

                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Trim();

                            //записываем строку в CSV-файл
                            RowCsv csvrow = new RowCsv();
                            // "НАИМЕНОВАН", 
                            csvrow.Add(RowCsv.Type.String, Name);
                            //"ЦЕНА1", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost1.ToString());
                            //"ЦЕНА5", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                            //"ЦЕНАОТ5", 
                            csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                            //"ПОРОГ1", 
                            csvrow.Add(RowCsv.Type.Numeric, Porog1);
                            //"ПОРОГ5", 
                            csvrow.Add(RowCsv.Type.Numeric, Porog5);
                            //"ЕДИЗМ1", 
                            csvrow.Add(RowCsv.Type.String, Measure1);
                            //"ЕДИЗМ5", 
                            csvrow.Add(RowCsv.Type.String, Measure5);
                            //"ЕДИЗМОТ5", 
                            csvrow.Add(RowCsv.Type.String, MeasureOt5);
                            //"ДЛИНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВЕС", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"МАРКА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПРОИЗВЕЛ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"СКЛАД", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВСЕГО", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЦЕНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДАТА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"НОМЕР", 
                            csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                            //"РЕЗЕРВ"
                            csvrow.Add(RowCsv.Type.Null);
                            //
                            writer.WriteLine(csvrow.GetRow());
                            rowcount++;



                        }
                    }
                }
            }



        }
        /// <summary>
        /// </summary>
        public static void Sheet_Шестигр_Пров_ка_электрод(DataTable tb, StreamWriter writer)
        {
          	string electr ="ЭЛЕКТРОДЫ";
            string fpname = "";
            string str = "";
            string tmpstr = "";
            //
            string Name="";
            
            
            string Measure1 = "";
            string Measure5 = "";
            string MeasureOt5 = "";
            string Porog1 = "";
            string Porog5 = "";
            decimal tmp=0;
            //
            try
            {
                int sp = GetBeginStr(tb, 0, "РАЗМЕР", 7, "ГОСТ");
                if (sp > 0)
                {
                    for (int icol = 0; icol <= 8; icol += 7)
                    {
                        int g = 0;
                        for (int irow = sp; irow < tb.Rows.Count; irow++)
                        {
                            DataRow row = tb.Rows[irow];
                            decimal Cost1 =0;
                            decimal.TryParse(row.ItemArray[4 + icol].ToString(), out tmp);
                            

                            tmpstr = row.ItemArray[0 + icol].ToString();

                            if (tmpstr.IndexOf("ГОСТ ") >= 0)
                            {
                                if (str.IndexOf(electr) < 0)  //не элетроды
                                {
                                    if (icol == 0)
                                    {
                                        if (tmpstr.IndexOf("ПРОВОЛОКА") < 0)
                                            fpname = "ШЕСТИГРАННИК " + tmpstr;
                                        else fpname = tmpstr;
                                    }
                                    else
                                    {
                                        fpname = "ПРОВОЛОКА " + tmpstr;
                                    }
                                }
                            }
                            else if (tmp > 0)
                            {
                                decimal Cost5 = 0;
                                decimal CostOt5 = 0;
                                if (fpname.IndexOf("ШЕСТИГРАННИК ") >= 0)
                                {
                                    
                                    decimal.TryParse(row.ItemArray[3 + icol].ToString(), out Cost1);
                                    decimal.TryParse(row.ItemArray[2 + icol].ToString(), out Cost5);
                                    decimal.TryParse(row.ItemArray[1 + icol].ToString(), out CostOt5);

                                    Name = fpname + " " + row.ItemArray[0 + icol].ToString() + " мм";
                                    // cost записаны
                                    Porog1 = "1";
                                    Porog5 = "5";
                                    Measure1 = "т";
                                    Measure5 = "т";
                                    MeasureOt5 = "т";
                                }
                                else if (tmpstr.Length > 0)     //'проволока
                                {
                                    Cost1 =0;
                                    Cost5 = 0;
                                    CostOt5 = 0;

                                    if (icol == 0) // 'в первой колонке
                                    {
                                        decimal.TryParse(row.ItemArray[4 + icol].ToString(), out Cost1);
                                        decimal.TryParse(row.ItemArray[3 + icol].ToString(), out Cost5);
                                        Name = fpname + " " + row.ItemArray[2 + icol].ToString() + " мм";
                                    }
                                    else
                                    {
                                        decimal.TryParse(row.ItemArray[3 + icol].ToString(), out Cost1);
                                        decimal.TryParse(row.ItemArray[2 + icol].ToString(), out Cost5);
                                        Name = fpname + " " + row.ItemArray[1 + icol].ToString() + " мм";
                                    }
                                    Cost1 = Cost1 * 1000;
                                    Cost5 = Cost5 * 1000;
                                    CostOt5 = 0;
                                    Porog1 = "1";
                                    Porog5 = "5";
                                    Measure1 = "т";
                                    Measure5 = "т";
                                    MeasureOt5 = "т";
                                }
                                else if (str.IndexOf(electr) >= 0) //'электроды - строка данных 
                                {
                                    Name = fpname + "ГОСТ " + row.ItemArray[1 + icol].ToString() + " " + row.ItemArray[2 + icol].ToString() + " мм";
                                    decimal.TryParse(row.ItemArray[3 + icol].ToString(), out Cost1);
                                    Cost1 = Cost1 * 1000;
                                    Measure1 = "т";
                                }

                                Name = Name.Replace("   ", " ");
                                Name = Name.Replace("  ", " ");
                                Name = Name.Replace("   ", " ");
                                Name = Name.Replace("  ", " ");
                                Name = Name.Trim();

                                //записываем строку в CSV-файл
                                RowCsv csvrow = new RowCsv();
                                // "НАИМЕНОВАН", 
                                csvrow.Add(RowCsv.Type.String, Name);
                                //"ЦЕНА1", 
                                csvrow.Add(RowCsv.Type.Numeric, Cost1!=0?Cost1.ToString():"");
                                //"ЦЕНА5", 
                                csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                                //"ЦЕНАОТ5", 
                                csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                                //"ПОРОГ1", 
                                csvrow.Add(RowCsv.Type.Numeric, Porog1);
                                //"ПОРОГ5", 
                                csvrow.Add(RowCsv.Type.Numeric, Porog5);
                                //"ЕДИЗМ1", 
                                csvrow.Add(RowCsv.Type.String, Measure1);
                                //"ЕДИЗМ5", 
                                csvrow.Add(RowCsv.Type.String, Measure5);
                                //"ЕДИЗМОТ5", 
                                csvrow.Add(RowCsv.Type.String, MeasureOt5);
                                //"ДЛИНА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ВЕС", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"МАРКА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ПРОИЗВЕЛ", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"СКЛАД", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ВСЕГО", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ЦЕНА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ЕДИЗМ", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ДАТА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"НОМЕР", 
                                csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                                //"РЕЗЕРВ"
                                csvrow.Add(RowCsv.Type.Null);
                                //
                                writer.WriteLine(csvrow.GetRow());
                                rowcount++;

                            }
                            else if (str.IndexOf(electr) >= 0) //подраздел таблицы элетроды
                            {
                                fpname = str + row.ItemArray[1 + icol].ToString() + " ";
                            }
                            else if (row.ItemArray[1 + icol].ToString().IndexOf(electr) >= 0)
                            {
                                str = electr + " ";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
        /// <summary>
        /// </summary>
        public static void Sheet_Тройники_переходы(DataTable tb, StreamWriter writer)
        {

            string fpname = "";
            string tmpstr = "";
            string Name = "";
            decimal Cost1 = 0;
            string Measure1 = "";
            string str="";

            int sp = GetBeginStr(tb, 0, "РАЗМЕР", 8, "РАЗМЕР");
            if (sp > 0)
            {
                for (int icol = 0; icol < 12; icol += 4)
                {
                    fpname = tb.Rows[sp + 2].ItemArray[0 + icol].ToString();
                   
                    str= " 08(12)Х18Н10Т "+ tb.Rows[sp + 3].ItemArray[0 + icol].ToString().Replace("с геометрией по","");


                    for (int irow = sp + 2; irow < tb.Rows.Count; irow++)
                    {
                        DataRow row = tb.Rows[irow];

                        tmpstr = row.ItemArray[0 + icol].ToString();

                        decimal.TryParse(row.ItemArray[1 + icol].ToString(), out Cost1);

                        if (tmpstr.Length > 0 && Cost1 > 0)
                        {
                            Name = fpname + " " + row.ItemArray[0 + icol].ToString() + " мм" + str;
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Trim();

                            Measure1 = "м";


                            //записываем строку в CSV-файл
                            RowCsv csvrow = new RowCsv();
                            // "НАИМЕНОВАН", 
                            csvrow.Add(RowCsv.Type.String, Name);
                            //"ЦЕНА1", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost1.ToString());
                            //"ЦЕНА5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЦЕНАОТ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПОРОГ1", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПОРОГ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ1", 
                            csvrow.Add(RowCsv.Type.String, Measure1);
                            //"ЕДИЗМ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМОТ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДЛИНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВЕС", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"МАРКА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПРОИЗВЕЛ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"СКЛАД", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВСЕГО", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЦЕНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДАТА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"НОМЕР", 
                            csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                            //"РЕЗЕРВ"
                            csvrow.Add(RowCsv.Type.Null);
                            //
                            writer.WriteLine(csvrow.GetRow());
                            rowcount++;


                        }

                    }
                }

            }
        }
        /// <summary>
        /// </summary>
        public static void Sheet_Круг(DataTable tb, StreamWriter writer)
        {
            string fpname = "";
            string str = "Круг нержавеющий ";
            string tmpstr = "";

            string Name = "";
            decimal Cost1 = 0;
            decimal Cost5 = 0;
            decimal CostOt5 = 0;
            string Measure1 = "т";
            string Measure5 = "т";
            string MeasureOt5 = "т";
            string Porog1 = "1";
            string Porog5 = "5";


            int sp = GetBeginStr(tb, 0, "РАЗМЕР", 12, "РАЗМЕР");
            if (sp > 0)
            {
                for (int icol = 0; icol < 18; icol += 6)
                {
                    for (int irow = sp ; irow < tb.Rows.Count; irow++)
                    {
                        DataRow row = tb.Rows[irow];
                        tmpstr = row.ItemArray[icol].ToString();
                        decimal.TryParse(row.ItemArray[1 + icol].ToString(), out Cost1);

                        if (tmpstr.Length > 0 && (Cost1 > 0))
                        {
                            Name = str + tmpstr + " мм " + fpname;
                            decimal.TryParse(row.ItemArray[3 + icol].ToString(), out Cost1);
                            decimal.TryParse(row.ItemArray[2 + icol].ToString(), out Cost5);
                            decimal.TryParse(row.ItemArray[1 + icol].ToString(), out CostOt5);
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Trim();

                            //записываем строку в CSV-файл
                            RowCsv csvrow = new RowCsv();
                            // "НАИМЕНОВАН", 
                            csvrow.Add(RowCsv.Type.String, Name);
                            //"ЦЕНА1", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost1 != 0 ? Cost1.ToString() : "");
                            //"ЦЕНА5", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                            //"ЦЕНАОТ5", 
                            csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                            //"ПОРОГ1", 
                            csvrow.Add(RowCsv.Type.Numeric, Porog1);
                            //"ПОРОГ5", 
                            csvrow.Add(RowCsv.Type.Numeric, Porog5);
                            //"ЕДИЗМ1", 
                            csvrow.Add(RowCsv.Type.String, Measure1);
                            //"ЕДИЗМ5", 
                            csvrow.Add(RowCsv.Type.String, Measure5);
                            //"ЕДИЗМОТ5", 
                            csvrow.Add(RowCsv.Type.String, MeasureOt5);
                            //"ДЛИНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВЕС", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"МАРКА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПРОИЗВЕЛ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"СКЛАД", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВСЕГО", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЦЕНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДАТА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"НОМЕР", 
                            csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                            //"РЕЗЕРВ"
                            csvrow.Add(RowCsv.Type.Null);
                            //
                            writer.WriteLine(csvrow.GetRow());
                            rowcount++;

                        }
                        else if (tmpstr.IndexOf("ГОСТ") >= 0 && (row.ItemArray[1 + icol].ToString().Length == 0) && (row.ItemArray[2 + icol].ToString().Length == 0))
                        {
                            fpname = tmpstr;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// </summary>
        public static void Sheet_отводы_AISI(DataTable tb, StreamWriter writer) 
        {
            string fpname = "";
            string tmpstr = "";

            string Name = "";
            decimal Cost1 = 0;
            decimal Cost5 = 0;
            decimal CostOt5 = 0;
            string Measure1 = "т";


            int sp = GetBeginStr(tb, 0, "РАЗМЕР", 8, "РАЗМЕР");
            if (sp > 0)
            {
                for (int icol = 0; icol <= 11; icol += 4)
                {
                    for (int irow = sp; irow < tb.Rows.Count; irow++)
                    {
                        DataRow row = tb.Rows[irow];

                        tmpstr = row.ItemArray[icol].ToString();
                        decimal.TryParse(row.ItemArray[2 + icol].ToString(), out Cost1);

                        if (tmpstr.Length>0 && Cost1 > 0)   
                        {
					         Name= fpname + row.ItemArray[0 + icol].ToString() + " мм";
                             Measure1 = "шт";
                             Name = Name.Replace("   ", " ");
                             Name = Name.Replace("  ", " ");
                             Name = Name.Replace("   ", " ");
                             Name = Name.Replace("  ", " ");
                             Name = Name.Trim();

                             //записываем строку в CSV-файл
                             RowCsv csvrow = new RowCsv();
                             // "НАИМЕНОВАН", 
                             csvrow.Add(RowCsv.Type.String, Name);
                             //"ЦЕНА1", 
                             csvrow.Add(RowCsv.Type.Numeric, Cost1 != 0 ? Cost1.ToString() : "");
                             //"ЦЕНА5", 
                             csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                             //"ЦЕНАОТ5", 
                             csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                             //"ПОРОГ1", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ПОРОГ5", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ЕДИЗМ1", 
                             csvrow.Add(RowCsv.Type.String, Measure1);
                             //"ЕДИЗМ5", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ЕДИЗМОТ5", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ДЛИНА", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ВЕС", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"МАРКА", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ПРОИЗВЕЛ", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"СКЛАД", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ВСЕГО", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ЦЕНА", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ЕДИЗМ", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"ДАТА", 
                             csvrow.Add(RowCsv.Type.Null);
                             //"НОМЕР", 
                             csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                             //"РЕЗЕРВ"
                             csvrow.Add(RowCsv.Type.Null);
                             //
                             writer.WriteLine(csvrow.GetRow());
                             rowcount++;
	      
                        }
	                    else if (tmpstr.IndexOf("AISI")>=0 && (row.ItemArray[ 1+icol].ToString().Length==0) && (tb.Rows[irow-1].ItemArray[icol].ToString().Length>0) )
                        {
	             	        fpname=tb.Rows[irow-1].ItemArray[icol].ToString() + " " + tmpstr + " ";
				        }





                    }
                }
            }

        
        
        
        }
        /// <summary>
        /// </summary>
        public static void Sheet_Отводы_фланцы_заглушки(DataTable tb, StreamWriter writer)
        {
            string fpname = "";
            string tmpstr = "";

            string Name = "";
            decimal Cost1 = 0;
            decimal Cost5 = 0;
            decimal CostOt5 = 0;
            string Measure1 = "";
            int u = 0;
            int irow = 0;
            int icol = 0;


            int sp = GetBeginStr(tb, 0, "РАЗМЕР", 5, "РАЗМЕР");
            if (sp > 0)
            {
                for (icol = 0; icol <= 5; icol += 5)
                {
                    for (irow = sp; irow < tb.Rows.Count; irow++)
                    {
                        DataRow row = tb.Rows[irow];
                        tmpstr = row.ItemArray[icol].ToString();
                        decimal.TryParse(row.ItemArray[2 + icol].ToString(), out Cost1);

                        if (tmpstr.Length > 0 && Cost1 > 0)
                        {
                            Name = fpname + row.ItemArray[0 + icol].ToString() + " мм";
                            Measure1 = "шт";

                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Replace("   ", " ");
                            Name = Name.Replace("  ", " ");
                            Name = Name.Trim();

                            //записываем строку в CSV-файл
                            RowCsv csvrow = new RowCsv();
                            // "НАИМЕНОВАН", 
                            csvrow.Add(RowCsv.Type.String, Name);
                            //"ЦЕНА1", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost1 != 0 ? Cost1.ToString() : "");
                            //"ЦЕНА5", 
                            csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                            //"ЦЕНАОТ5", 
                            csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                            //"ПОРОГ1", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПОРОГ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ1", 
                            csvrow.Add(RowCsv.Type.String, Measure1);
                            //"ЕДИЗМ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМОТ5", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДЛИНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВЕС", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"МАРКА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ПРОИЗВЕЛ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"СКЛАД", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ВСЕГО", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЦЕНА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ЕДИЗМ", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"ДАТА", 
                            csvrow.Add(RowCsv.Type.Null);
                            //"НОМЕР", 
                            csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                            //"РЕЗЕРВ"
                            csvrow.Add(RowCsv.Type.Null);
                            //
                            writer.WriteLine(csvrow.GetRow());
                            rowcount++;
                        }
                        else if (tmpstr.IndexOf("ГОСТ") >= 0 && tb.Rows[irow].ItemArray[1 + icol].ToString().Length == 0 && tb.Rows[irow - 1].ItemArray[icol].ToString().Length > 0)
                        {
                            fpname = tb.Rows[irow - 1].ItemArray[icol].ToString() + " " + tmpstr + " ";
                        }
                        else if (icol > 0)
                        {
                            if (tb.Rows[irow].ItemArray[icol - 1].ToString().IndexOf("ГОСТ") >= 0)
                            {
                                fpname = tb.Rows[irow].ItemArray[icol - 1].ToString();
                                u = icol - 1;
                                break;
                            }
                        }

                    }
                }



                string[] tstr = new string[15];


                //таблица по фланцам
                if (tb.Rows[irow + 1].ItemArray[u].ToString().IndexOf("Ду") >= 0 && fpname.Length > 0)
                {
                    irow = irow + 1;
                    icol = u;
                    for (int i = icol + 1; i < icol + 10; i++)
                    {
                        tstr[i - icol - 1] = tb.Rows[irow].ItemArray[i].ToString();
                        if (tstr[i - icol - 1].Length == 0)
                        {
                            u = i - icol - 2;
                            break;
                        }
                    }

                    for (irow = irow + 1; irow < tb.Rows.Count; irow++)
                    {
                        tmpstr = tb.Rows[irow].ItemArray[icol].ToString();

                        if (tmpstr.Length > 0 && tmpstr.Length < 4)
                        {
                            for (int i = u ; i >= 0; i--)
                            {
                                Name = fpname + " Ду" + tmpstr + " " + tstr[i];
                                Cost5 = 0;
                                CostOt5 = 0;
                                decimal.TryParse(tb.Rows[irow].ItemArray[i + icol+1].ToString(), out Cost1);

                                if (Cost1 > 0)
                                {
                                    Measure1 = "шт";
                                }
                                else continue;

                                Name = Name.Replace("   ", " ");
                                Name = Name.Replace("  ", " ");
                                Name = Name.Replace("   ", " ");
                                Name = Name.Replace("  ", " ");
                                Name = Name.Trim();

                                //записываем строку в CSV-файл
                                RowCsv csvrow = new RowCsv();
                                // "НАИМЕНОВАН", 
                                csvrow.Add(RowCsv.Type.String, Name);
                                //"ЦЕНА1", 
                                csvrow.Add(RowCsv.Type.Numeric, Cost1.ToString());
                                //"ЦЕНА5", 
                                csvrow.Add(RowCsv.Type.Numeric, Cost5 != 0 ? Cost5.ToString() : "");
                                //"ЦЕНАОТ5", 
                                csvrow.Add(RowCsv.Type.Numeric, CostOt5 != 0 ? CostOt5.ToString() : "");
                                //"ПОРОГ1", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ПОРОГ5", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ЕДИЗМ1", 
                                csvrow.Add(RowCsv.Type.String, Measure1);
                                //"ЕДИЗМ5", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ЕДИЗМОТ5", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ДЛИНА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ВЕС", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"МАРКА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ПРОИЗВЕЛ", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"СКЛАД", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ВСЕГО", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ЦЕНА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ЕДИЗМ", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"ДАТА", 
                                csvrow.Add(RowCsv.Type.Null);
                                //"НОМЕР", 
                                csvrow.Add(RowCsv.Type.Numeric, rowcount.ToString());
                                //"РЕЗЕРВ"
                                csvrow.Add(RowCsv.Type.Null);
                                //
                                writer.WriteLine(csvrow.GetRow());
                                rowcount++;


                                //пишем в таблу
                            }
                        }
                    }
                }
            }
        }



        public static int GetBeginStr(DataTable tb, int NumColumn1, string strfind1 ,int NumColumn2, string strfind2)
    {
         for (int irow = 0; irow < tb.Rows.Count; irow++)
         {
               DataRow row = tb.Rows[irow];
            if( row.ItemArray[NumColumn1].ToString().ToUpper().IndexOf(strfind1.ToUpper()) >=0 &&  
                row.ItemArray[NumColumn2].ToString().ToUpper().IndexOf(strfind2.ToUpper()) >=0)
            {
                return irow;
            }
         }
        return -1;
    }




        //класс для формирования CSV-файлов
        public class TableScv : DataTable
        {
            public static bool SaveScvFile(string File)
            {

                return false;
            }

        }

        //временный класс
        public class RowCsv
        {
            public enum Type
            {
                Null,
                String,
                Numeric
            }

            public void SetItem(string colname, object g) 
            {
            }

            private string row = "";

            public void Add(Type type)
            {
                if (type == Type.Null) row += "\t";
            }
            public void Add(Type type, string data)
            {
                if (type == Type.Null) row += "\t";
                if (type == Type.Numeric) row += data + "\t";
                if (type == Type.String) row += "\"" + data + "\"" + "\t";
            }

            /// <summary>
            /// 
            /// </summary>
            public string GetRow()
            {
                
                return row;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        public static class Logger
        {
            public static void Message(string text)
            {
                StreamWriter sw = File.AppendText("ContinentalParser.log");
                sw.WriteLine(text);
                sw.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        static string WebRequestClient()
        {
            WebRequest request = WebRequest.Create(addrPrice);
            request.Credentials = CredentialCache.DefaultCredentials;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            Encoding encode = System.Text.Encoding.GetEncoding("Windows-1251");
            StreamReader reader = new StreamReader(dataStream, encode);
            string responseFromServer = reader.ReadToEnd();
            reader.Close();
            dataStream.Close();
            response.Close();

            return responseFromServer;
        }



    }
}
