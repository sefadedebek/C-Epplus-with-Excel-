using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace project
{

    class Program
    {

        static void Main(string[] args)
        {


            FileInfo file2 = new FileInfo(@"FORMUL-GIRDI.xlsx");
            ExcelPackage package2 = new ExcelPackage(file2);
            ExcelWorksheet FORMUL = package2.Workbook.Worksheets[1];


                double sifir_bes_kisa = Convert.ToDouble(FORMUL.Cells[2, 2].Value);
                double sifir_bes_uzun = Convert.ToDouble(FORMUL.Cells[2, 3].Value);
                double alti_on_kisa = Convert.ToDouble(FORMUL.Cells[3, 2].Value);
                double alti_on_uzun = Convert.ToDouble(FORMUL.Cells[3, 3].Value);
                double onbir_onbes_kisa = Convert.ToDouble(FORMUL.Cells[4, 2].Value);
                double onbir_onbes_uzun = Convert.ToDouble(FORMUL.Cells[4, 3].Value);
                double onalti_yirmi_kisa = Convert.ToDouble(FORMUL.Cells[5, 2].Value);
                double onalti_yirmi_uzun = Convert.ToDouble(FORMUL.Cells[5, 3].Value);
                double yirmibir_otuz_kisa = Convert.ToDouble(FORMUL.Cells[6, 2].Value);
                double yirmibir_otuz_uzun = Convert.ToDouble(FORMUL.Cells[6, 3].Value);
                double artan_kisa = Convert.ToDouble(FORMUL.Cells[7, 2].Value);
                double artan_uzun = Convert.ToDouble(FORMUL.Cells[7, 3].Value);




                double hesapla(double adet, double desi, String mesafe)
                {

                    double ucret = 0;


                    if (desi >= 0 && desi < 6)
                    {
                        if (mesafe == "ŞEHİRİÇİ" || mesafe == "KISA" || mesafe == "YAKIN")
                            ucret = adet * sifir_bes_kisa;
                        else if (mesafe == "UZAK" || mesafe == "ORTA")
                            ucret = adet * sifir_bes_uzun;
                        else
                            ucret = 0;

                    }

                    else if (desi >= 6 && desi < 11)
                    {
                        if (mesafe == "ŞEHİRİÇİ" || mesafe == "KISA" || mesafe == "YAKIN")
                            ucret = adet * alti_on_kisa;
                        else if (mesafe == "UZAK" || mesafe == "ORTA")
                            ucret = adet * alti_on_uzun;
                        else
                            ucret = 0;
                    }

                    else if (desi >= 11 && desi < 16)
                    {
                        if (mesafe == "ŞEHİRİÇİ" || mesafe == "KISA" || mesafe == "YAKIN")
                            ucret = adet * onbir_onbes_kisa;
                        else if (mesafe == "UZAK" || mesafe == "ORTA")
                            ucret = adet * onbir_onbes_uzun;
                        else
                            ucret = 0;
                    }

                    else if (desi >= 16 && desi < 21)
                    {
                        if (mesafe == "ŞEHİRİÇİ" || mesafe == "KISA" || mesafe == "YAKIN")
                            ucret = adet * onalti_yirmi_kisa;
                        else if (mesafe == "UZAK" || mesafe == "ORTA")
                            ucret = adet * onalti_yirmi_uzun;
                        else
                            ucret = 0;
                    }

                    else if (desi >= 21 && desi <= 30)
                    {
                        if (mesafe == "ŞEHİRİÇİ" || mesafe == "KISA" || mesafe == "YAKIN")
                            ucret = adet * yirmibir_otuz_kisa;
                        else if (mesafe == "UZAK" || mesafe == "ORTA")
                            ucret = adet * yirmibir_otuz_uzun;
                        else
                            ucret = 0;
                    }

                    else if (desi > 30)
                    {
                        if (mesafe == "ŞEHİRİÇİ" || mesafe == "KISA" || mesafe == "YAKIN")
                            ucret = adet * (yirmibir_otuz_kisa + ((desi - 30) * artan_kisa));
                        else if (mesafe == "UZAK" || mesafe == "ORTA")
                            ucret = adet * (yirmibir_otuz_uzun + ((desi - 30) * artan_uzun));
                        else
                            ucret = 0;
                    }
                    else
                    {
                        ucret = 0;

                    }
                    return ucret;

                }



            FileInfo file = new FileInfo(@"EKSTRE-GIRDI.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {

                excelPackage.Workbook.Worksheets.Add("Pivot Rapor");


                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];

                worksheet.Name = "İşlem Sonucu";


                int rowCount = worksheet.Dimension.End.Row;     //satır sayısını al


                worksheet.Cells[1, 5].Value = "ÜCRET";


                for (int row = 2; row <= rowCount; row++)
                {

                    double adet;
                    double desi;
                    String mesafe;
                    //double ucret;

                    adet = Convert.ToDouble(worksheet.Cells[row, 2].Value.ToString());
                    desi = Convert.ToDouble(worksheet.Cells[row, 3].Value.ToString());
                    mesafe = worksheet.Cells[row, 4].Value.ToString();
                    //worksheet.Cells[row, 5].Value = ucretdizisi[row];
                    worksheet.Cells[row, 5].Value = hesapla(adet, desi, mesafe);
                }


                ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets[2];



                //kaynak sayfasındaki veri aralığını tanımlama
                var dataRange = worksheet.Cells[worksheet.Dimension.Address];

                //pivot tablo oluşturma
                var pivotTable = worksheet2.PivotTables.Add(worksheet2.Cells["B2"], dataRange, "PivotTable");


                //label alanı
                pivotTable.RowFields.Add(pivotTable.Fields["Mesafe"]);
                pivotTable.DataOnRows = false;


                //veri alanları
                var field = pivotTable.DataFields.Add(pivotTable.Fields["ADET"]);
                field.Name = "Kargo Adeti (Her Satır Bİr Adet Olacak Şekilde) ";
                field.Function = DataFieldFunctions.Count;

                field = pivotTable.DataFields.Add(pivotTable.Fields["ADET"]);
                field.Name = "Toplam Kargo Adeti";
                field.Function = DataFieldFunctions.Sum;

                field = pivotTable.DataFields.Add(pivotTable.Fields["KG_DESI"]);
                field.Name = "Kargo Desisi (KG) ";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "0.00";

                field = pivotTable.DataFields.Add(pivotTable.Fields["ÜCRET"]);
                field.Name = "Toplam Ücret";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "₺#,##0.00";


                FileInfo excelFile1 = new FileInfo(@"C:\Users\sefad\Desktop\RAPOR.xlsx");
                FileInfo excelFile2 = new FileInfo(@"D:\RAPOR.xlsx");
                //FileInfo excelFile3 = new FileInfo(@"RAPOR.xlsx");


                excelPackage.SaveAs(excelFile1);
                excelPackage.SaveAs(excelFile2);
                //excelPackage.SaveAs(excelFile3);

            }

        }


    }


}

