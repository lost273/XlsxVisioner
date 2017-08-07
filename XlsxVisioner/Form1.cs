using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;

namespace XlsxVisioner
{
    public partial class Form1 : Form
    {
        int rowCounter = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            rowCounter = 0;
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";
            //После выбора файла создается новый объект «Application» или приложение «Excel», 
            //которое может содержать одну или более книг, ссылки на которые содержит свойство «Workbooks».
            ExcelObj.Application app = new ExcelObj.Application();
            //Книги - объекты «Workbook», могут содержать одну или более страниц, ссылки на которые содержит свойство «Worksheets».
            ExcelObj.Workbook workbook;
            //Страницы – «Worksheet», могут содержать объекты ячейки или группы ячеек, ссылки на которые становятся доступными через объект «Range».
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            //Полученные данные из файла будут заноситься в таблицу «dt», созданную с использованием класса «DataTable».
            DataTable dt = new DataTable();

            //Массив общей выручки
            double[] totalInfo = {0,0,0,0};

            int sheetsCount = 1;
            //В коде присутствует проверка, что пользователь действительно выбрал файл, если данное условие выполнено, 
            //в текстовое поле с помощью свойства «FileName», класса «OpenFileDialog» помещается путь, 
            //имя и расширение выбранного файла в элемент управления «textBox1».
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.SafeFileName;
                //Для открытия существующего документа используется метод «Open» из набора «Excel.Workbooks», 
                //в качестве основного параметра указывается путь к файлу, остальные параметры остаются пустыми.
                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

                progressBar1.Maximum = workbook.Sheets.Count;

                do
                {
                    //Устанавливаем номер листа из котрого будут извлекаться данные
                    //Листы нумеруются от 1
                    //workbook.Sheets.Count = значение общего кол-ва листов
                    NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(sheetsCount);
                    //Чтобы получить объект Microsoft.Office.Interop.Excel.Range, 
                    //который представляет все ячейки, содержащие значение на данный момент, 
                    //используется свойство страницы «Worksheet.UsedRange».
                    ShtRange = NwSheet.UsedRange;
                    //После получения объекта «Range», с помощью цикла «For» загружается первая строка из таблицы 
                    //и каждое значение устанавливается в качестве имени колонки таблицы.

                    if (sheetsCount == 1) //первая инициализация
                    {
                        for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                        {
                            dt.Columns.Add(new DataColumn(Convert.ToString((ShtRange.Cells[2, Cnum] as ExcelObj.Range).Value2)));
                        }
                        dt.AcceptChanges();

                      /*  string[] columnNames = new String[dt.Columns.Count];
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            columnNames[0] = dt.Columns[i].ColumnName;
                        }*/

                        //Далее таким же способом загружаются все оставшиеся строки с добавлением в таблицу.
                        for (int Rnum = 3; Rnum <= ShtRange.Rows.Count; Rnum++)
                        {
                            DataRow dr = dt.NewRow();
                            bool ifmod = true;
                            for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                            {
                                if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                                {
                                    dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                                }
                            }
                            for (int tRnum = 0; tRnum < dt.Rows.Count; tRnum++)
                            {
                                if (((dr[0].ToString() == dt.Rows[tRnum][0].ToString()) && (dt.Rows[tRnum][0].ToString() != "") && (dr[2].ToString() == dt.Rows[tRnum][2].ToString())))
                                {
                                    dt.Rows[tRnum][1] = Convert.ToDouble(dr[1]) + Convert.ToDouble(dt.Rows[tRnum][1]);
                                    dt.Rows[tRnum][3] = Convert.ToDouble(dr[3]) + Convert.ToDouble(dt.Rows[tRnum][3]);
                                    ifmod = false;
                                }
                            }
                            if ((dr[0].ToString() != "")&&(ifmod))
                            {
                                dt.Rows.Add(dr);
                                dt.AcceptChanges();
                            }
                        }
                        //собираем информацию в определенных ячейках
                        totalInfo[0] = Convert.ToDouble(dt.Rows[0][4]); //4 столбец - Сумма за месяц
                        totalInfo[1] = Convert.ToDouble(dt.Rows[2][5]); //5 столбец - Сумма по Z
                        totalInfo[2] = Convert.ToDouble(dt.Rows[1][5]); //5 столбец - Сумма по CARD
                    }
                    else //последующее считывание листов
                    {
                        for (int Rnum = 3; Rnum <= ShtRange.Rows.Count; Rnum++)
                        {
                            DataRow dr = dt.NewRow();
                            for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                            {
                                if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                                {
                                    dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                                }
                            }

                            //собираем информацию в определенных ячейках
                            if ((Rnum == 3)&&(dr[4].ToString() != ""))
                            {
                                totalInfo[0] = totalInfo[0] + Convert.ToDouble(dr[4]); //4 столбец - Сумма за месяц
                            }
                            if ((Rnum == 5)&&(dr[5].ToString() != ""))
                            {
                                totalInfo[1] = totalInfo[1] + Convert.ToDouble(dr[5]); //5 столбец - Сумма по Z
                            }
                            if ((Rnum == 4)&&(dr[5].ToString()!= ""))
                            {
                                totalInfo[2] = totalInfo[2] + Convert.ToDouble(dr[5]); //5 столбец - Сумма по CARD
                            }

                            //сравниваем и модифицируем итоговую таблицу
                            bool ifmod=true;
                            for (int tRnum = 0; tRnum < dt.Rows.Count; tRnum++)
                            {
                                if (((dr[0].ToString() == dt.Rows[tRnum][0].ToString())&&(dt.Rows[tRnum][0].ToString()!= "")&&(dr[2].ToString() == dt.Rows[tRnum][2].ToString())))
                                {
                                    dt.Rows[tRnum][1] = Convert.ToDouble(dr[1])+ Convert.ToDouble(dt.Rows[tRnum][1]);
                                    dt.Rows[tRnum][3] = Convert.ToDouble(dr[3]) + Convert.ToDouble(dt.Rows[tRnum][3]);
                                    ifmod = false;
                                }
                            }
                            if ((ifmod)&&(dr[0].ToString()!=""))
                            {
                                dt.Rows.Add(dr);
                                dt.AcceptChanges();
                            }
                        }
                    }
                    sheetsCount++;
                    progressBar1.Value += 1;
                } while (workbook.Sheets.Count >= sheetsCount);
                //удаляем последние 2 столбца
                dt.Columns.RemoveAt(4);
                dt.Columns.RemoveAt(4);
                //По завершении загрузки данных с указанного листа, сформированная таблица «dt» подключается к элементу управления «dataGridView1». 
                //Так же открытый объект «Application» или приложение «Excel» закрывается.
                app.Quit();
                // Находим значения и формируем результатирующую таблицу
                
                dataGridView1.DataSource = dt;
                
              

                progressBar1.Value = 0;

                Chart1.Visible = true;
                Chart1.Series.Clear();
                Chart1.Titles.Clear();
                // Форматировать диаграмму
                Chart1.BackColor = Color.Gray;
                Chart1.BackSecondaryColor = Color.WhiteSmoke;
                Chart1.BackGradientStyle = GradientStyle.DiagonalRight;

                Chart1.BorderlineDashStyle = ChartDashStyle.Solid;
                Chart1.BorderlineColor = Color.Gray;
                Chart1.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;

                // Форматировать область диаграммы
                Chart1.ChartAreas[0].BackColor = Color.Wheat;

                // Добавить и форматировать заголовок
                Chart1.Titles.Add("Структура продаж");
                Chart1.Titles[0].Font = new Font("Courier New", 10);

                Chart1.Series.Add(new Series("ColumnSeries") {ChartType = SeriesChartType.Pie});
                //формируем массив значений для графика
                string[] xValues = { "Заправка лазерных - ", "Заправка струйных - ", "Ремонт картриджей - ", "Ремонт принтера - ", "Чернила - ", "Печать - ", "Товар - " };
                double[] yValues= {0,0,0,0,0,0,0};

                for (int Rnum = 0; Rnum < dt.Rows.Count; Rnum++)
                {
                    if (dt.Rows[Rnum][0].ToString() == "заправка лазерного") { yValues[0] += Convert.ToDouble(dt.Rows[Rnum][3]);}
                    else if (dt.Rows[Rnum][0].ToString() == "заправка струйного") { yValues[1] += Convert.ToDouble(dt.Rows[Rnum][3]); }
                     else if (dt.Rows[Rnum][0].ToString() == "ремонт картриджа") { yValues[2] += Convert.ToDouble(dt.Rows[Rnum][3]); }
                      else if (dt.Rows[Rnum][0].ToString() == "ремонт принтера") { yValues[3] += Convert.ToDouble(dt.Rows[Rnum][3]); }
                       else if (dt.Rows[Rnum][0].ToString() == "чернила") { yValues[4] += Convert.ToDouble(dt.Rows[Rnum][3]); }
                        else if (dt.Rows[Rnum][0].ToString() == "печать") { yValues[5] += Convert.ToDouble(dt.Rows[Rnum][3]); }
                         else  { yValues[6] += Convert.ToDouble(dt.Rows[Rnum][3]); }
                }
                
                xValues[0] = xValues[0] + Math.Round((yValues[0] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues[1] = xValues[1] + Math.Round((yValues[1] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues[2] = xValues[2] + Math.Round((yValues[2] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues[3] = xValues[3] + Math.Round((yValues[3] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues[4] = xValues[4] + Math.Round((yValues[4] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues[5] = xValues[5] + Math.Round((yValues[5] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues[6] = xValues[6] + Math.Round((yValues[6] / totalInfo[0] * 100), 2).ToString() + " %";
                
                Chart1.Series["ColumnSeries"].Points.DataBindXY(xValues, yValues);
                Chart1.Series["ColumnSeries"].IsValueShownAsLabel = true;
                
                Chart1.ChartAreas[0].Area3DStyle.Enable3D = true;

                // второй график

                Chart2.Visible = true;
                Chart2.Series.Clear();
                Chart2.Titles.Clear();
                // Форматировать диаграмму
                Chart2.BackColor = Color.Gray;
                Chart2.BackSecondaryColor = Color.WhiteSmoke;
                Chart2.BackGradientStyle = GradientStyle.DiagonalRight;

                Chart2.BorderlineDashStyle = ChartDashStyle.Solid;
                Chart2.BorderlineColor = Color.Gray;
                Chart2.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;

                // Форматировать область диаграммы
                Chart2.ChartAreas[0].BackColor = Color.Wheat;

                // Добавить и форматировать заголовок
                Chart2.Titles.Add("Общая выручка - " + totalInfo[0].ToString() + " руб");
                Chart2.Titles[0].Font = new Font("Courier New", 10);

                Chart2.Series.Add(new Series("ColumnSeries") { ChartType = SeriesChartType.Pie });
                //формируем массив значений для графика
                string[] xValues2 = { "Z отчет - ", "CARD отчет - ", "Остаток - "};
                //вычисляем остаток
                totalInfo[3]= totalInfo[0]- totalInfo[1];
                double[] yValues2 = { totalInfo[1] - totalInfo[2], totalInfo[2], totalInfo[3]};
                
                xValues2[0] = xValues2[0] + Math.Round((yValues2[0] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues2[1] = xValues2[1] + Math.Round((yValues2[1] / totalInfo[0] * 100), 2).ToString() + " %";
                xValues2[2] = xValues2[2] + Math.Round((yValues2[2] / totalInfo[0] * 100), 2).ToString() + " %";

                Chart2.Series["ColumnSeries"].Points.DataBindXY(xValues2, yValues2);
                Chart2.Series["ColumnSeries"].IsValueShownAsLabel = true;

                Chart2.ChartAreas[0].Area3DStyle.Enable3D = true;

                button2.Visible = true;
                

            }
            else
                Application.Exit();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            
            int z = 0;
            StringFormat str = new StringFormat();
            str.Alignment = StringAlignment.Near;
            str.LineAlignment = StringAlignment.Center;
            str.Trimming = StringTrimming.EllipsisCharacter;
            
            int width = 500 / (dataGridView1.Columns.Count - 1); // ширина ячейки
            int realwidth = 100; // общая ширина
            int height = 15; // высота строки

            int realheight = 60; // общая высота


            // Рисуем название файла
            if (rowCounter == 0)
            {
                e.Graphics.FillRectangle(Brushes.ForestGreen, realwidth, realheight, width, height);
                e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height);
                e.Graphics.DrawString(textBox1.Text, dataGridView1.Font, Brushes.Black, realwidth, realheight);

                realheight = realheight + height;
            }

            // Рисуем названия колонок
            for (z = 0; z < dataGridView1.Columns.Count; z++)
            {
                e.Graphics.FillRectangle(Brushes.RoyalBlue, realwidth, realheight, width, height);
                e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height);

                e.Graphics.DrawString(dataGridView1.Columns[z].HeaderText, dataGridView1.Font, Brushes.Black, realwidth, realheight);

                realwidth = realwidth + width;
            }
            realheight = realheight + height;

            // Рисуем остальную таблицу
            while (rowCounter < dataGridView1.Rows.Count)
            {
                realwidth = 100;

                if (dataGridView1.Rows[rowCounter].Cells[0].Value == null)
                {
                    dataGridView1.Rows[rowCounter].Cells[0].Value = "";
                }
                e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height);
                e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height);
                e.Graphics.DrawString(dataGridView1.Rows[rowCounter].Cells[0].Value.ToString(), dataGridView1.Font, Brushes.Black, realwidth, realheight);
                realwidth = realwidth + width;
                
                if (dataGridView1.Rows[rowCounter].Cells[1].Value == null)
                {
                    dataGridView1.Rows[rowCounter].Cells[1].Value = "";
                }
                e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height);
                e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height);
                e.Graphics.DrawString(dataGridView1.Rows[rowCounter].Cells[1].Value.ToString(), dataGridView1.Font, Brushes.Black, realwidth, realheight);
                realwidth = realwidth + width;
                
                if (dataGridView1.Rows[rowCounter].Cells[2].Value == null)
                {
                    dataGridView1.Rows[rowCounter].Cells[2].Value = "";
                }
                e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height);
                e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height);
                e.Graphics.DrawString(dataGridView1.Rows[rowCounter].Cells[2].Value.ToString(), dataGridView1.Font, Brushes.Black, realwidth, realheight);
                realwidth = realwidth + width;
                
                if (dataGridView1.Rows[rowCounter].Cells[3].Value == null)
                {
                    dataGridView1.Rows[rowCounter].Cells[3].Value = "";
                }
                e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height);
                e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height);
                e.Graphics.DrawString(dataGridView1.Rows[rowCounter].Cells[3].Value.ToString(), dataGridView1.Font, Brushes.Black, realwidth, realheight);
                realwidth = realwidth + width;

                ++rowCounter;
                realheight = realheight + height;

                //если 1000 пикселей уже напечатано - переводим на новый лист
                if (realheight >= 1000) { e.HasMorePages = true; break; }
                // если 1000 пикселей не напечатано, а таблица закончилась
                if ((realheight < 1000) && (rowCounter >= dataGridView1.Rows.Count)) { e.HasMorePages = false; rowCounter = 0; break; }

            }

            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            printDialog1.ShowDialog();
            printPreviewDialog1.ShowDialog();
            //this.printDocument1.Print();
        }

        private void Chart2_Click(object sender, EventArgs e)
        {
            Chart2.Printing.PrintPreview();
        }

        private void Chart1_Click(object sender, EventArgs e)
        {
            Chart1.Printing.PrintPreview();
        }
    }
}
