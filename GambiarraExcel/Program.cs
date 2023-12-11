using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace GambiarraExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            ReadExcelAndWriteInserts();
        }

        private static void ReadExcelAndWriteInserts()
        {
            var filePath = "C:\\Users\\thiag\\Downloads\\ctfapp.xlsx";
            var txtFilePath = "C:\\Users\\thiag\\Desktop\\inserts.txt";

            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(filePath);
            Worksheet ws = wb.Worksheets[1];
            File.WriteAllText(txtFilePath, string.Empty);

            StreamWriter writer = new StreamWriter(txtFilePath);
            try
            {
                Range usedRange = ws.UsedRange;

                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                writer.WriteLine("TRUNCATE TABLE cte.CORRESPONDENCIACTFAPP");

                //dados da primeira linha
                string codCnae = "0151-2/01";
                string codAtividade = "21-74";
                for (int row = 1; row <= rowCount; row++)
                {
                    //os descritores tão aparecendo sem CNAE
                    var desDescritor = "NULL";
                    string nextCodCnae = usedRange.Cells[row, 1].Value?.ToString().Replace(" ", "");

                    if (string.IsNullOrEmpty(nextCodCnae))
                    {
                        desDescritor = $"(SELECT TOP(1) idDescritor FROM CNAEDESCRITORES WHERE desDescritor = '{usedRange.Cells[row, 2].Value?.ToString().Replace("'", "''")}')";
                        nextCodCnae = codCnae;
                    }
                    else
                    {
                        codCnae = nextCodCnae;
                    }

                    string precisaCTFAPP = usedRange.Cells[row, 12].Value?.ToString().Trim();
                    string nextCodAtividade = usedRange.Cells[row, 13].Value?.ToString();
                    if (!string.IsNullOrEmpty(precisaCTFAPP) && !precisaCTFAPP.ToLower().Contains("não"))
                        codAtividade = nextCodAtividade;
                    else if (!string.IsNullOrEmpty(precisaCTFAPP) && precisaCTFAPP.ToLower().Contains("não"))
                        codAtividade = "NULL";

                    nextCodAtividade = codAtividade;

                    string descricao = string.IsNullOrEmpty(usedRange.Cells[row, 14].Value?.ToString()) ? "NULL" : $"'{usedRange.Cells[row, 14].Value?.ToString().Replace("'", "''")}'";

                    string insert = $"INSERT INTO cte.CORRESPONDENCIACTFAPP (flaAtivo, ordem, descricao, idCNAE, idDescritor, codCategoriaCTF, idAtividade) VALUES(1, 999, {descricao}, (SELECT TOP(1) idCNAE FROM CNAE WHERE codCNAE = '{nextCodCnae}'), {desDescritor}, {(nextCodAtividade == "NULL" ? "NULL" : "'" + nextCodAtividade.Split('-')[0] + "'")}, {(nextCodAtividade == "NULL" ? "NULL" : $"(SELECT TOP(1) idAtividade FROM CAEATIVIDADE WHERE codAtividade = '{nextCodAtividade}')")})";

                    writer.WriteLine(insert);
                    Console.WriteLine(row + " - " + insert);
                }

                Console.WriteLine($"SQL INSERT statements have been written to file");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                wb.Close(false);
                excel.Quit();
                writer.Dispose();

                // Release COM objects to avoid memory leaks
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }
    }
}
