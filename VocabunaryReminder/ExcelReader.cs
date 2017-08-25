using System;
using System.Windows;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace VocabunaryReminder
{
    class ExcelReader
    {
        public void GenerateEXcelData(string SlnoAbbreviation)
        {
            OleDbConnection oledbConn = new OleDbConnection();
            try
            {
                string path = System.IO.Path.GetFullPath("JLPT-N4-EXCEL_.xlsx");


                if (Path.GetExtension(path) == ".xls")
                {
                    oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"");
                }
                else if (Path.GetExtension(path) == ".xlsx")
                {
                    oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");
                }
                oledbConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                cmd.Connection = oledbConn;
                if (!String.IsNullOrEmpty(SlnoAbbreviation) && SlnoAbbreviation != "Select")
                {
                    cmd.CommandText = "SELECT TOP 1 [Kanji], [Hiragana], [Example], [Translation],[Dictionary Entry],[JLPT Kanji],[JLPT Furigana],[Romaji],[JLPT English] " +
                        "  FROM [JLPT-N4$] where [Kanji]= @Slno_Abbreviation";
                    cmd.Parameters.AddWithValue("@Slno_Abbreviation", SlnoAbbreviation);
                }
                else
                {
                    cmd.CommandText = "SELECT [Kanji], [Hiragana], [Example], [Translation],[Dictionary Entry],[JLPT Kanji],[JLPT Furigana],[Romaji],[JLPT English] FROM [JLPT-N4$]";
                }
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                oledbConn.Close();
            }
        }
    }
}
