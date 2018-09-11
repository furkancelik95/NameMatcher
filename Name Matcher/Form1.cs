using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Threading;

namespace Name_Matcher
{
    public partial class Form1 : Form
    {
        //List<string> alfabe = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        List<string> kelimeler = new List<string>();
        List<string> karşılaştırma = new List<string>();
        string textadı;
        string excelad;

        public Form1()
        {
            InitializeComponent();
            checkBox2.Checked = true;
            MessageBox.Show("Lütfen işleme başlamadan önce işlem yapacağınız dosyanın bir kopyasını kaydedin.");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Excel Dosyası Yükle";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = openFileDialog1.FileName;
            }
            excelad = textBox4.Text;
            FileInfo newfile = new FileInfo(textBox4.Text);
            ExcelPackage package = new ExcelPackage(newfile);

            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

            //var rowCnt = worksheet.Dimension.End.Row;
            var colCnt = worksheet.Dimension.End.Column;

            for (int j = 1; j <= colCnt; j++)
            {
                string deger = worksheet.Cells[1, j].Text;
                karşılaştırma.Add(deger);
            }

            for (int i = 0; i < colCnt; i++)
            {
                comboBox1.Items.Add(karşılaştırma[i]);
                comboBox2.Items.Add(karşılaştırma[i]);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Text Dosyası Yükle";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "txt";
            openFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textadı = openFileDialog1.FileName;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {

            kelimeler = File.ReadAllLines(textadı, Encoding.Default).ToList();

            for (int i = 0; i < kelimeler.Count; i++)
            {
                listBox1.Items.Add(kelimeler.ElementAt(i));
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkBox2.Checked == true && checkBox1.Checked == false)
                {
                    var firstValue = comboBox1.SelectedIndex;
                    var secondvalue = comboBox2.SelectedIndex;

                    FileInfo newfile = new FileInfo(textBox4.Text);
                    ExcelPackage package = new ExcelPackage(newfile);

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    var rowCnt = worksheet.Dimension.End.Row;
                    var colCnt = worksheet.Dimension.End.Column;

                    for (int i = 1; i <= rowCnt; i++)
                    {
                        for (int j = 1; j <= colCnt; j++)
                        {
                            string ilkdeger = worksheet.Cells[i, firstValue + 1].Text.ToString();
                            string ikincideger = worksheet.Cells[i, secondvalue + 1].Text.ToString();
                            string ilkdeger_yedek = ilkdeger.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ");
                            string ikincideger_yedek = ikincideger.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ");
                            string sonuc2 = "";
                            for (int k = 0; k < listBox1.Items.Count; k++)
                            {
                                ilkdeger_yedek = ilkdeger_yedek.Replace(listBox1.Items[k].ToString(), "");
                                ikincideger_yedek = ikincideger_yedek.Replace(listBox1.Items[k].ToString(), "");
                            }
                            double levensdegeri = CalculateSimilarity(ilkdeger_yedek, ikincideger_yedek);
                            List<string> sonuc = CheckSubWordsString(ilkdeger_yedek, ikincideger_yedek, kelimeler);
                            sonuc = sonuc.Distinct().ToList();
                            string matchtype = MatchType(ilkdeger_yedek, ikincideger_yedek, sonuc, levensdegeri);
                            int count = CheckSubWordsCount(ilkdeger_yedek, ikincideger_yedek, kelimeler);

                            for (int k = 0; k < sonuc.Count(); k++)
                            {
                                sonuc2 = sonuc2 + sonuc[k] + " ";
                            }
                            if (i == 1)
                            {
                                worksheet.Cells[i, colCnt + 1].Value = "Levenshtein Değeri";
                                worksheet.Cells[i, colCnt + 2].Value = "MatchType";
                                worksheet.Cells[i, colCnt + 3].Value = "MatchingWords";
                                worksheet.Cells[i, colCnt + 4].Value = "MatchingWordsCount";
                            }
                            else
                            {
                                sonuc2 = sonuc2.Replace("-", "");
                                worksheet.Cells[i, colCnt + 1].Value = levensdegeri;
                                worksheet.Cells[i, colCnt + 2].Value = matchtype;
                                worksheet.Cells[i, colCnt + 3].Value = sonuc2.Trim();
                                worksheet.Cells[i, colCnt + 4].Value = count;
                            }
                        }
                        if (i >= rowCnt / 1.5)
                        {
                            label1.Text = "75%";
                            label1.Invalidate();
                            label1.Update();
                        }
                        else if (i >= rowCnt / 2)
                        {
                            label1.Text = "50%";
                            label1.Invalidate();
                            label1.Update();
                        }
                        else if (i >= rowCnt / 4)
                        {
                            label1.Text = "25%";
                            label1.Invalidate();
                            label1.Update();
                        }
                    }
                    package.SaveAs(newfile);
                    label1.Text = "100%";
                    MessageBox.Show("İşlem tamamlandı.");
                }
                else if (checkBox1.Checked == true && checkBox2.Checked == false)
                {
                    var firstValue = comboBox1.SelectedIndex;
                    var secondvalue = comboBox2.SelectedIndex;

                    FileInfo newfile = new FileInfo(textBox4.Text);
                    ExcelPackage package = new ExcelPackage(newfile);

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    var rowCnt = worksheet.Dimension.End.Row;
                    var colCnt = worksheet.Dimension.End.Column;

                    for (int i = 1; i <= rowCnt; i++)
                    {
                        for (int j = 1; j <= colCnt; j++)
                        {
                            string ilkdeger = worksheet.Cells[i, firstValue + 1].Text.ToString();
                            string ikincideger = worksheet.Cells[i, secondvalue + 1].Text.ToString();
                            string ilkdeger_yedek = ilkdeger.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ");
                            string ikincideger_yedek = ikincideger.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ");
                            string sonuc2 = "";
                            for (int k = 0; k < listBox1.Items.Count; k++)
                            {
                                ilkdeger_yedek = ilkdeger_yedek.Replace(listBox1.Items[k].ToString(), "");
                                ikincideger_yedek = ikincideger_yedek.Replace(listBox1.Items[k].ToString(), "");
                            }
                            int soundexdegeri = Difference(ilkdeger_yedek, ikincideger_yedek);
                            double levensdegeri = CalculateSimilarity(ilkdeger_yedek, ikincideger_yedek);
                            List<string> sonuc = CheckSubWordsString(ilkdeger_yedek, ikincideger_yedek, kelimeler);
                            sonuc = sonuc.Distinct().ToList();
                            string matchtype = MatchType(ilkdeger_yedek, ikincideger_yedek, sonuc, levensdegeri);
                            int count = CheckSubWordsCount(ilkdeger_yedek, ikincideger_yedek, kelimeler);

                            for (int k = 0; k < sonuc.Count(); k++)
                            {
                                sonuc2 = sonuc2 + sonuc[k] + " ";
                            }
                            if (i == 1)
                            {
                                worksheet.Cells[i, colCnt + 1].Value = "Soundex Değeri";
                                worksheet.Cells[i, colCnt + 2].Value = "MatchType";
                                worksheet.Cells[i, colCnt + 3].Value = "MatchingWords";
                                worksheet.Cells[i, colCnt + 4].Value = "MatchingWordsCount";
                            }
                            else
                            {
                                sonuc2 = sonuc2.Replace("-", "");
                                worksheet.Cells[i, colCnt + 1].Value = soundexdegeri;
                                worksheet.Cells[i, colCnt + 2].Value = matchtype;
                                worksheet.Cells[i, colCnt + 3].Value = sonuc2.Trim();
                                worksheet.Cells[i, colCnt + 4].Value = count;
                            }
                        }
                        if (i >= rowCnt / 1.5)
                        {
                            label1.Text = "75%";
                            label1.Invalidate();
                            label1.Update();
                        }
                        else if (i >= rowCnt / 2)
                        {
                            label1.Text = "50%";
                            label1.Invalidate();
                            label1.Update();
                        }
                        else if (i >= rowCnt / 4)
                        {
                            label1.Text = "25%";
                            label1.Invalidate();
                            label1.Update();
                        }
                    }
                    package.SaveAs(newfile);
                    label1.Text = "100%";
                    MessageBox.Show("İşlem tamamlandı.");
                }
                else if (checkBox1.Checked == true && checkBox2.Checked == true)
                {
                    var firstValue = comboBox1.SelectedIndex;
                    var secondvalue = comboBox2.SelectedIndex;

                    FileInfo newfile = new FileInfo(textBox4.Text);
                    ExcelPackage package = new ExcelPackage(newfile);

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    var rowCnt = worksheet.Dimension.End.Row;
                    var colCnt = worksheet.Dimension.End.Column;

                    for (int i = 1; i <= rowCnt; i++)
                    {
                        for (int j = 1; j <= colCnt; j++)
                        {
                            string ilkdeger = worksheet.Cells[i, firstValue + 1].Text.ToString();
                            string ikincideger = worksheet.Cells[i, secondvalue + 1].Text.ToString();
                            string ilkdeger_yedek = ilkdeger.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ");
                            string ikincideger_yedek = ikincideger.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ");
                            string sonuc2 = "";
                            for (int k = 0; k < listBox1.Items.Count; k++)
                            {
                                ilkdeger_yedek = ilkdeger_yedek.Replace(listBox1.Items[k].ToString(), "");
                                ikincideger_yedek = ikincideger_yedek.Replace(listBox1.Items[k].ToString(), "");
                            }

                            double levensdegeri = CalculateSimilarity(ilkdeger_yedek, ikincideger_yedek);
                            int soundexdegeri = Difference(ilkdeger_yedek, ikincideger_yedek);
                            List<string> sonuc = CheckSubWordsString(ilkdeger_yedek, ikincideger_yedek, kelimeler);
                            sonuc = sonuc.Distinct().ToList();
                            string matchtype = MatchType(ilkdeger_yedek, ikincideger_yedek, sonuc, levensdegeri);
                            int count = CheckSubWordsCount(ilkdeger_yedek, ikincideger_yedek, kelimeler);

                            for (int k = 0; k < sonuc.Count(); k++)
                            {
                                sonuc2 = sonuc2 + sonuc[k] + " ";
                            }
                            if (i == 1)
                            {
                                worksheet.Cells[i, colCnt + 1].Value = "Levenshtein Değeri";
                                worksheet.Cells[i, colCnt + 2].Value = "Soundex Değeri";
                                worksheet.Cells[i, colCnt + 3].Value = "MatchType";
                                worksheet.Cells[i, colCnt + 4].Value = "MatchingWords";
                                worksheet.Cells[i, colCnt + 5].Value = "MatchingWordsCount";
                            }
                            else
                            {
                                sonuc2 = sonuc2.Replace("-", "");
                                worksheet.Cells[i, colCnt + 1].Value = levensdegeri;
                                worksheet.Cells[i, colCnt + 2].Value = soundexdegeri;
                                worksheet.Cells[i, colCnt + 3].Value = matchtype;
                                worksheet.Cells[i, colCnt + 4].Value = sonuc2.Trim();
                                worksheet.Cells[i, colCnt + 5].Value = count;
                            }
                        }
                        if (i >= rowCnt / 1.5)
                        {
                            label1.Text = "75%";
                            label1.Invalidate();
                            label1.Update();
                        }
                        else if (i >= rowCnt / 2)
                        {
                            label1.Text = "50%";
                            label1.Invalidate();
                            label1.Update();
                        }
                        else if (i >= rowCnt / 4)
                        {
                            label1.Text = "25%";
                            label1.Invalidate();
                            label1.Update();
                        }
                    }
                    package.SaveAs(newfile);
                    label1.Text = "100%";
                    MessageBox.Show("İşlem tamamlandı.");
                }
                else
                {
                    MessageBox.Show("Lütfen bir karşılaştırma metodu(Levenshtein veya Soundex) seçiniz.");
                }
            }
            catch
            {
                MessageBox.Show("Lütfen karşılaştırılacak sütunları seçtiğinizden emin olup tekrar deneyiniz.");
            }
        }
        public static string Soundex(string data)
        {
            StringBuilder result = new StringBuilder();

            if (data != null && data.Length > 0)
            {
                string previousCode = "", currentCode = "", currentLetter = "";

                result.Append(data.Substring(0, 1));

                for (int i = 1; i < data.Length; i++)
                {
                    currentLetter = data.Substring(i, 1).ToLower();
                    currentCode = "";

                    if ("bfpv".IndexOf(currentLetter) > -1)
                        currentCode = "1";
                    else if ("cgğjkqsxz".IndexOf(currentLetter) > -1)
                        currentCode = "2";
                    else if ("dt".IndexOf(currentLetter) > -1)
                        currentCode = "3";
                    else if (currentLetter == "l")
                        currentCode = "4";
                    else if ("mn".IndexOf(currentLetter) > -1)
                        currentCode = "5";
                    else if (currentLetter == "r")
                        currentCode = "6";

                    if (currentCode != previousCode)
                        result.Append(currentCode);

                    if (result.Length == 4) break;

                    if (currentCode != "")
                        previousCode = currentCode;
                }
            }
            if (result.Length < 4)
                result.Append(new String('0', 4 - result.Length));

            return result.ToString().ToUpper();
        }

        public static int Difference(string data1, string data2)
        {
            int result = 0;
            string soundex1 = Soundex(data1);
            string soundex2 = Soundex(data2);

            if (soundex1 == soundex2)
                result = 4;
            else
            {
                string sub1 = soundex1.Substring(1, 3);
                string sub2 = soundex1.Substring(2, 2);
                string sub3 = soundex1.Substring(1, 2);
                string sub4 = soundex1.Substring(1, 1);
                string sub5 = soundex1.Substring(2, 1);
                string sub6 = soundex1.Substring(3, 1);

                if (soundex2.IndexOf(sub1) > -1)
                    result = 3;
                else if (soundex2.IndexOf(sub2) > -1)
                    result = 2;
                else if (soundex2.IndexOf(sub3) > -1)
                    result = 2;
                else
                {
                    if (soundex2.IndexOf(sub4) > -1)
                        result++;

                    if (soundex2.IndexOf(sub5) > -1)
                        result++;

                    if (soundex2.IndexOf(sub6) > -1)
                        result++;
                }
                if (soundex1.Substring(0, 1) == soundex2.Substring(0, 1))
                    result++;
            }
            return (result == 0) ? 1 : result;
        }

        public static int CheckSubWordsCount(string data1, string data2, List<string> res)
        {
            data1 = data1.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ").ToUpper().Trim();
            data2 = data2.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ").ToUpper().Trim();
            int count = 0;
            string[] words = data1.Split(' ');
            words = words.Distinct().ToArray();
            for (int i = 0; i < words.Length; i++)
            {
                if (data2.Contains(words[i]))
                    if (!res.Contains(words[i]))
                        if (words[i] != "0" && words[i] != "1" && words[i] != "2" && words[i] != "3" && words[i] != "4" &&
                            words[i] != "5" && words[i] != "6" && words[i] != "7" && words[i] != "8" && words[i] != "9")
                            if(words[i].Count()>1)
                            count++;
            }
            return count;
        }

        public static string MatchType(string data1, string data2, List<string> sonuc, double levens)
        {
            data1 = data1.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ").ToUpper().Trim();
            data2 = data2.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").Replace("  ", " ").ToUpper().Trim();
            string[] words = data1.Split(' ');
            if (data1.Equals(data2))
                return "Exact";
            else if (sonuc.Count() >= 2 && levens >= 0.8)
                return "Exact";
            else if (sonuc.Count() >= 3 && levens <= 0.4)
                return "CloseToExact";
            else if (sonuc.Count() >= 2 && levens >= 0.51)
                return "CloseToExact";
            else if ((data2.StartsWith(words[0]) && data2.EndsWith(words.Last())) || (data2.StartsWith(words.Last()) && data2.EndsWith(words[0])))
                return "StartsEndsWith";
            else if (data2.StartsWith(words[0]))
                return "StartsWith";
            else if (data2.EndsWith(words.Last()))
                return "EndsWith";
            else if ((sonuc.Count() >= 2 && levens < 0.4) || (sonuc.Count() < 2 && levens >= 0.51))
                return "DifferentLetters";
            else if (sonuc.Count() == 1)
                return "DifferentWords";
            else
                return "No Match";
        }

        public static List<string> CheckSubWordsString(string data1, string data2, List<string> kontrol)
        {
            data1 = data1.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("  ", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").ToUpper().Trim();
            data2 = data2.Replace("ı", "i").Replace("İ", "I")
                                        .Replace("ö", "o").Replace("Ö", "O")
                                        .Replace("ç", "c").Replace("Ç", "C")
                                        .Replace("ş", "s").Replace("Ş", "S")
                                        .Replace("ğ", "g").Replace("Ğ", "G")
                                        .Replace("ü", "u").Replace("Ü", "U").Replace("-", " ").Replace("  ", " ").Replace("/", " ")
                                        .Replace("\"", " ").Replace(".", " ").Replace("(", " ").Replace(")", " ").ToUpper().Trim();
            List<string> result = new List<string>();
            result.Clear();
            string[] words = data1.Split(' ');
            for (int i = 0; i < words.Length; i++)
            {
                if (data2.Contains(words[i]))
                    if (!kontrol.Contains(words[i]))
                        if (words[i] != "0" && words[i] != "1" && words[i] != "2" && words[i] != "3" && words[i] != "4" &&
                            words[i] != "5" && words[i] != "6" && words[i] != "7" && words[i] != "8" && words[i] != "9" )
                            if(words[i].Count()>1)
                            result.Add(words[i]);
            }
            return result;
        }

        double CalculateSimilarity(string source, string target)
        {
            if ((source == null) || (target == null)) return 0.0;
            if ((source.Length == 0) || (target.Length == 0)) return 0.0;
            if (source == target) return 1.0;

            int stepsToSame = ComputeLevenshteinDistance(source, target);
            return (1.0 - ((double)stepsToSame / (double)Math.Max(source.Length, target.Length)));
        }

        int ComputeLevenshteinDistance(string source, string target)
        {
            if ((source == null) || (target == null)) return 0;
            if ((source.Length == 0) || (target.Length == 0)) return 0;
            if (source == target) return source.Length;

            int sourceWordCount = source.Length;
            int targetWordCount = target.Length;

            // Step 1
            if (sourceWordCount == 0)
                return targetWordCount;

            if (targetWordCount == 0)
                return sourceWordCount;

            int[,] distance = new int[sourceWordCount + 1, targetWordCount + 1];

            // Step 2
            for (int i = 0; i <= sourceWordCount; distance[i, 0] = i++) ;
            for (int j = 0; j <= targetWordCount; distance[0, j] = j++) ;

            for (int i = 1; i <= sourceWordCount; i++)
            {
                for (int j = 1; j <= targetWordCount; j++)
                {
                    // Step 3
                    int cost = (target[j - 1] == source[i - 1]) ? 0 : 1;

                    // Step 4
                    distance[i, j] = Math.Min(Math.Min(distance[i - 1, j] + 1, distance[i, j - 1] + 1), distance[i - 1, j - 1] + cost);
                }
            }
            return distance[sourceWordCount, targetWordCount];
        }
    }

}
