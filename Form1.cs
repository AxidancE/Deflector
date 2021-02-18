using AutoUpdaterDotNET;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProjectZ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.SelectTab(tabPage1);
            KeyPreview = true; //Переменная перехватывает нажатие клавиш, что позволяет использовать "горячие" клавиши.
            AutoUpdater.Synchronous = true;
            AutoUpdater.Start("https://raw.githubusercontent.com/AxidancE/Deflector/main/Version.xml");
        }

        public static string path = Application.StartupPath.ToString() + @"\Resources\textes";

        private Excel.Application xlApp;
        private Excel.Workbook xlAppBook;
        private Excel.Workbooks xlAppBooks;
        private Excel.Sheets xlSheets;
        private int flagexcelapp = 0;
        private readonly String strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) //Функция перехвата клавиш
        {
            if (keyData == (Keys.Control | Keys.Enter))
            {
                button1.PerformClick();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        public void Main_prog()
        {
            /*
            TODO
            Оптимизация - порезать куски кода.
            Структуризация - тут всё понятно.
            */
            int defect;
            double width, height, const_width;
            width = Convert.ToDouble(textBox3.Text); //Длина трубы на чертеже
            height = Convert.ToDouble(textBox5.Text); //Высота трубы на чертеже
            const_width = 2;

            double width_scheme, out_width, pre_width, out_height, pre_height; ;
            //Переменная для длины трубы по формам (11м прим.)
            //Длина дефекта
            //Длина до дефекта
            //Высота дефекта
            //Высота до дефекта

            double r_radius, r_radius_second;
            double x_deg, y_deg, one_deg;
            double const_deg = 90;

            string original_width, original_height;

            if (textBox1.Text == "")
            {
                width_scheme = 1;
                textBox1.Text = "1";
            }
            else
            {
                width_scheme = Convert.ToInt32(textBox1.Text);
            }

            defect = Convert.ToInt32(numericUpDown1.Value);

            Excel.Worksheet xlSheet_01 = (Excel.Worksheet)xlApp.Worksheets.get_Item(5);
            ((Excel.Worksheet)this.xlApp.ActiveWorkbook.Sheets[5]).Select();
            Excel.Range xlRange_01 = xlApp.get_Range("D3", $"D{xlSheet_01.UsedRange.Rows.Count}");

            try
            {
                Excel.Range currentFind_inTry = xlRange_01.Find(defect);
                Excel.Range range_def_inTry = xlSheet_01.Cells[currentFind_inTry.Row, 5];
            }
            catch
            {
                MessageBox.Show("Ошибка данных. Дефект не найден." +
                    "\nПроверьте настройки сортировки.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                numericUpDown1.Value++;
                return;
            }

            //as Excel.Range
            Excel.Range currentFind = xlRange_01.Find(defect); //Поиск по клетке
            Excel.Range range_def = xlSheet_01.Cells[currentFind.Row, 5];
            Excel.Range width_range = xlSheet_01.Cells[currentFind.Row, 8];
            Excel.Range height_range = xlSheet_01.Cells[currentFind.Row, 9];
            Excel.Range F_deg = xlSheet_01.Cells[currentFind.Row, 19];
            Excel.Range S_deg = xlSheet_01.Cells[currentFind.Row, 20];
            Excel.Range defect_orig = xlSheet_01.Cells[currentFind.Row, 22];

            //MessageBox.Show(xlSheet_01.Cells[currentFind.Row, 22].toString());

            try
            {
                //Console.WriteLine("1");
                double range_def_do_inTry = Convert.ToDouble(range_def.Value2);
                double width_range_do_inTry = Convert.ToDouble(width_range.Value2);
                double height_range_do_inTry = Convert.ToDouble(height_range.Value2);
                double F_deg_do_inTry = Convert.ToDouble(F_deg.Value2);
                double S_deg_do_inTry = Convert.ToDouble(S_deg.Value2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка данных. Проверьте значения в Excel.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                dataGridView1[0, 0].Value = range_def.Value2.ToString(); //Расстояние от начала
                dataGridView1[1, 0].Value = width_range.Value2.ToString(); //Длина продольная
                dataGridView1[2, 0].Value = height_range.Value2.ToString(); //Длина окружная
                dataGridView1[3, 0].Value = F_deg.Value2.ToString(); //Начальный градус
                dataGridView1[4, 0].Value = S_deg.Value2.ToString(); //Конечный градус
                numericUpDown1.Value++;
                return;
            }

            double range_def_do = Convert.ToDouble(range_def.Value2);
            double width_range_do = Convert.ToDouble(width_range.Value2);
            double height_range_do = Convert.ToDouble(height_range.Value2);
            double F_deg_do = Convert.ToDouble(F_deg.Value2);
            double S_deg_do = Convert.ToDouble(S_deg.Value2);

            range_def_do *= 1000;

            dataGridView1[0, 0].Value = range_def_do.ToString(); //Расстояние от начала
            dataGridView1[1, 0].Value = width_range_do.ToString(); //Длина продольная
            dataGridView1[2, 0].Value = height_range_do.ToString(); //Длина окружная
            dataGridView1[3, 0].Value = F_deg_do.ToString(); //Начальный градус
            dataGridView1[4, 0].Value = S_deg_do.ToString(); //Конечный градус

            //MessageBox.Show($"{range_def_do}\n{width_range}\n{height_range}\n{F_deg}\n{S_deg}");

            defect = Convert.ToInt32(defect_orig.Value2);

            double x, y, a, b, c;
            double Sh_x1, Sh_y1, Sh_x2, Sh_y2;

            foreach (DataGridViewRow rw in this.dataGridView1.Rows)
            {
                for (int i = 0; i < rw.Cells.Count; i++)
                {
                    if (rw.Cells[i].Value == null || rw.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[i].Value.ToString()))
                    {
                        rw.Cells[i].Value = 0;
                        //MessageBox.Show(i.ToString());
                        //Если в клетке пусто, то она становится равна нулю
                    }
                }
            }

            x = Convert.ToDouble(dataGridView1[0, 0].Value.ToString()); // Длина до дефекта
            y = Convert.ToDouble(dataGridView1[1, 0].Value.ToString()); // Длина продольная
            a = Convert.ToDouble(dataGridView1[2, 0].Value.ToString()); // Длина окруж
            b = Convert.ToDouble(dataGridView1[3, 0].Value.ToString()); // Угол начала
            c = Convert.ToDouble(dataGridView1[4, 0].Value.ToString()); // Угол конца

            //Console.WriteLine("Отладка: Х - " + x);

            original_width = dataGridView1[1, 0].Value.ToString();
            original_height = dataGridView1[2, 0].Value.ToString();

            // Переменные для перерасчета
            double b_orig, c_orig;
            b_orig = b;
            c_orig = c;
            //Console.WriteLine
            double b_deg, c_deg;
            //b - 1 точка || c - 2 точка

            if (radioButton5.Checked == false)
            {
                Counted(b, c, out b_deg, out c_deg);
            }
            else
            {
                b += 90;
                c += 90;
                if (b > 359) b -= 360;
                if (c > 359) c -= 360;
                if (F_deg_do == 0 && S_deg_do > 357)
                {
                    b = 0; c = 360;
                }
                Counted(b, c, out b_deg, out c_deg);
            }

            if (b == c)
            {
                int i = 0, j = 3;
                while (b != 0 && i < j)
                {
                    b--;
                    i++;
                }
                i = 2;
                while (c != 359 && i < j)
                {
                    c++;
                    i++;
                }
            }

            // x в начале дефекта
            pre_width = x / width_scheme * width;
            pre_width = Math.Round(pre_width, 2);

            // x тела дефекта
            out_width = y / width_scheme * width;
            out_width = Math.Round(out_width, 2);

            if (out_width < const_width)
            {
                while (out_width < const_width)
                {
                    out_width += 0.3;
                    y++;
                }
            }
            // y в начале дефекта
            pre_height = -(b_deg / 180 * height);
            pre_height = Math.Round(pre_height, 2);

            // y тела дефекта
            out_height = -(Math.Abs(b_deg - c_deg) / 180 * height);
            out_height = Math.Round(out_height, 2);

            double const_rad = 37;
            if (radioButton4.Checked != true)
            {
                const_rad++;
            }

            //Переменные для полочки продольной
            Sh_x1 = pre_width;
            Sh_y1 = pre_height;

            Sh_x2 = pre_width + out_width;
            Sh_y2 = pre_height + out_height;

            /*
            Создавать нужно 2 дуги. Их центр в 0-0.
            Самые высокие значения 140-140
            Используется полярная Система Координат

            r - Длина вектора от центра до точки.
            f - Полярный угол (угол наклона вектора от центра до точки).*
            a - Поворот точки на угол (а) относительно центра координат.**

            *  По стандарту f = 90, из-за количества градусов колена трубы.
            ** Если угол отрицательный, поворот реверсируется.

            2 основные формулы:

            x = r*sin(f-a);
            y = r*cos(f-a);

            //Нужно брать длину трубы и делить ее на 90. Получать количество длины равной одному градусу.
            //Расстояние до дефекта (pre_width) - умножать на длину в 1 гр (11 при 1000) - получать количество градусов
            \которое вычитается из 90 (разность не должна быть отрицательной)

            // (70 / 180) * 1deg
            // (70 / 180) (0.38~) * 2deg
            // 140 - 1deg
            // 140 - 2 deg
            // высоту трубы (70) делить на 180 (половину развертки), и умножать на нужный угол
             */

            double cfd, cfd_orig;//c_for_defect
            if (b > c)
            {
                cfd = ((b + c) / 2) - 180;
                cfd_orig = cfd;
            }
            else
            {
                cfd = (b + c) / 2;
                cfd_orig = cfd;
            }

            cfd = DegreeToRadian(cfd);
            double x1_sin_def, y1_cos_def;

            x1_sin_def = const_rad * (Math.Sin(cfd));
            y1_cos_def = const_rad * (Math.Cos(cfd));
            x1_sin_def = Math.Round(x1_sin_def, 2);
            y1_cos_def = Math.Round(y1_cos_def, 2);

            if (radioButton4.Checked == true)
            {
                y1_cos_def -= 130;
                x1_sin_def += pre_width;
            }

            //--- Математика для отвода ---//
            one_deg = width_scheme / 90.0; //207 = 2.3
            x_deg = x / one_deg; //100 = 43.47
            y_deg = (x + y) / one_deg;
            //x - "Длина до дефекта"
            //y - "длина дефекта"

            // Высоту делить на полную развертку и умножать на полученную верхнюю/нижнюю точку дефекта
            r_radius = 140.0 - 70.0 / 180.0 * b_deg;
            r_radius_second = 140.0 - 70.0 / 180.0 * c_deg;

            //Возможно это отрезки для отвода
            Overload_Sin_Cos(out double x1_cos_elbow, out double y1_sin_elbow, r_radius, x_deg, const_deg);
            Overload_Sin_Cos(out double x2_cos_elbow, out double y2_sin_elbow, r_radius, y_deg, const_deg);

            Overload_Sin_Cos(out double x1_cos_elbow_second, out double y1_sin_elbow_second, r_radius_second, x_deg, const_deg);
            Overload_Sin_Cos(out double x2_cos_elbow_second, out double y2_sin_elbow_second, r_radius_second, y_deg, const_deg);

            //--- Работа с текстовыми файлами касательно дефекта ---//

            string text = File.ReadAllText(path + @"\Deflex.cdm", System.Text.Encoding.GetEncoding(1251));

            // sw_case 5 - прямая
            // sw_case 6 - отвод
            // Сделать выноски отдельными для каждого типа

            string ChangeText(string TextToChange, int sw_case, string TextToString)
            {
                string OriginalName = TextToChange;
                TextToChange = File.ReadAllText(path + @"\" + OriginalName + ".txt", System.Text.Encoding.GetEncoding(1251));

                //Создание дуги дефекта на окружности в разрезе
                Overload_Sin_Cos(out double x1_Sliced_cos, out double y1_Sliced_sin, const_rad, c, 90);
                Overload_Sin_Cos(out double x2_Sliced_cos, out double y2_Sliced_sin, const_rad, b, 90);
                //Console.WriteLine($"1.\n{Math.Round(x1_Sliced_cos,2)} - x1\n{Math.Round(y1_Sliced_sin, 2)} - y1; C - {c}\n");
                //Console.WriteLine($"2.\n{Math.Round(x2_Sliced_cos,2)} - x2\n{Math.Round(y2_Sliced_sin, 2)} - y2; B - {b}");

                double Pipe_type = 0;

                if (radioButton4.Checked == true)
                {
                    Pipe_type = -130;
                    y1_Sliced_sin -= 130;
                    y2_Sliced_sin -= 130;
                }

                //Создание дефекта на прямой трубе
                if (sw_case == 5)
                {
                    TextToChange = TextToChange.Replace("iObjParam.x = 0.0", "iObjParam.x = " + pre_width);
                    TextToChange = TextToChange.Replace("iObjParam.y = 0.0", "iObjParam.y = " + pre_height);
                    TextToChange = TextToChange.Replace("iObjParam.height = 0.0", "iObjParam.height = " + out_height);
                    TextToChange = TextToChange.Replace("iObjParam.width = 0.0", "iObjParam.width = " + out_width);

                    TextToChange = TextToChange.Replace("ksArcByPoint(0.0, -130.0, 37.0, 0, 0, 0, 0, 1, 7)", //x, y, r, x1, y1, x2, y2 [?, ?]
                "ksArcByPoint(" +
                $"{pre_width}, " +
                Pipe_type + ", " +
                const_rad + ", " +
                (x1_Sliced_cos + pre_width) + ", " +
                y1_Sliced_sin + ", " +
                (x2_Sliced_cos + pre_width) + ", " +
                y2_Sliced_sin + ", 1, 7)");
                }

                //Создание дефекта на отводе
                else if (sw_case == 6)
                {
                    //Создание дуги дефекта на разрезе
                    TextToChange = TextToChange.Replace("ksArcByPoint(0.0, 0.0, 38.0, 0, 0, 0, 0, 1, 7)", //x, y, r, x1, y1, x2, y2 [?, ?]
                "ksArcByPoint(" +
                "0.0, " +
                Pipe_type + ", " +
                const_rad + ", " +
                x1_Sliced_cos + ", " +
                y1_Sliced_sin + ", " +
                x2_Sliced_cos + ", " +
                y2_Sliced_sin + ", 1, 7)");

                    TextToChange = TextToChange.Replace("x1", //x, y, r, x1, y1, x2, y2 [1 - против часовой | -1 по часовой, 7 - вид линии]
                "iDocument2D.ksArcByPoint(" + 0.0 + ", 0.0, " +
                r_radius + ", " +
                x1_cos_elbow + ", " +
                y1_sin_elbow + ", " +
                x2_cos_elbow + ", " +
                y2_sin_elbow + ", -1, 7)");

                    TextToChange = TextToChange.Replace
                        (
                        "y1", //x, y, r, x1, y1, x2, y2 [?, ?]
                            "iDocument2D.ksArcByPoint(" + 0.0 + ", 0.0, " +
                            r_radius_second + ", " +
                            x1_cos_elbow_second + ", " +
                            y1_sin_elbow_second + ", " +
                            x2_cos_elbow_second + ", " +
                            y2_sin_elbow_second + ", -1, 7)"
                        );

                    TextToChange = TextToChange.Replace(
                        "x2", $@"iDocument2D.ksLineSeg(
                            {x1_cos_elbow},
                            {y1_sin_elbow},
                            {x1_cos_elbow_second},
                            {y1_sin_elbow_second}, 7)"//x1, y1, x2, y2
                        );
                    TextToChange = TextToChange.Replace(
                        "y2", $@"iDocument2D.ksLineSeg(
                            {x2_cos_elbow_second},
                            {y2_sin_elbow_second},
                            {x2_cos_elbow},
                            {y2_sin_elbow}, 7)"//x1, y1, x2, y2
                        );
                }

                if (sw_case == 7)
                {
                    TextToChange = TextToChange.Replace("x1 = 70.0", "x1   = " + Sh_x1);
                    TextToChange = TextToChange.Replace("y1 = 0.0", "y1   = " + Sh_y1);
                    TextToChange = TextToChange.Replace("ang1 = 0.0", "ang1  = " + (90 - x_deg));
                    TextToChange = TextToChange.Replace("rad = 105.0", "rad  = " + (r_radius + 8));
                    TextToChange = TextToChange.Replace("str = \"400 \"", "str  = " + TextToString); //Текст
                }
                if (sw_case == 8)
                {
                    TextToChange = TextToChange.Replace("x1 = 70.0", "x1   = " + Sh_x1); //x1 - правая
                    TextToChange = TextToChange.Replace("y1 = 0.0", "y1   = " + Sh_y1);
                    TextToChange = TextToChange.Replace("x2 = 0.0", "x2   = " + Sh_x2); //x2 - левая точка
                    TextToChange = TextToChange.Replace("y2 = 70.0", "y2   = " + Sh_y2);
                    TextToChange = TextToChange.Replace("ang1 = 0.0", "ang1 = " + (90 - x_deg));
                    TextToChange = TextToChange.Replace("ang2 = 90.0", "ang2 = " + (90 - y_deg));
                    TextToChange = TextToChange.Replace("rad = 105.0", "rad  = " + (r_radius + 8));
                    TextToChange = TextToChange.Replace("str = \"400 \"", "str  = " + TextToString); //Текст

                    //Console.WriteLine("1)" + x_deg + "; " + y_deg);
                }

                //Нет у do_def
                if (sw_case != 3)
                    TextToChange = TextToChange.Replace("iLDimSourceParam.x1 = 0.0", "iLDimSourceParam.x1 = " + Sh_x1);

                //Не изменяется у всех
                TextToChange = TextToChange.Replace("iLDimSourceParam.y1 = 0.0", "iLDimSourceParam.y1 = " + Sh_y1);

                //Меняется только у prodol
                TextToChange = TextToChange.Replace("iLDimSourceParam.x2 = 0.0", "iLDimSourceParam.x2 = " + Sh_x1);
                if (sw_case == 1)
                    TextToChange = TextToChange.Replace("iLDimSourceParam.x2 = " + Sh_x1, "iLDimSourceParam.x2 = " + Sh_x2);
                //Не меняется только у okruj
                TextToChange = TextToChange.Replace("iLDimSourceParam.y2 = 0.0", "iLDimSourceParam.y2 = " + Sh_y1);
                if (sw_case == 2)
                    TextToChange = TextToChange.Replace("iLDimSourceParam.y2 = " + Sh_y1, "iLDimSourceParam.y2 = " + Sh_y2);
                if (sw_case == 9)
                {
                    TextToChange = TextToChange.Replace("iLDimSourceParam.y2 = " + Sh_y1, "iLDimSourceParam.y2 = " + y1_sin_elbow_second);
                    TextToChange = TextToChange.Replace("iLDimSourceParam.x2 = " + Sh_x1, "iLDimSourceParam.x2 = " + x1_cos_elbow_second);
                }
                if (sw_case == 11)
                {
                    TextToChange = TextToChange.Replace("Change_me", Convert.ToString(pre_width));
                    TextToChange = TextToChange.Replace("Y_ch_1", Convert.ToString(pre_height));
                    if (pre_width < 102.5)
                    {
                        TextToChange = TextToChange.Replace("Directioned", "100");
                    }
                    else
                    {
                        TextToChange = TextToChange.Replace("Directioned", "(-180)");
                    }
                    //Console.WriteLine("X = " + pre_width.ToString());
                }
                TextToChange = TextToChange.Replace("iChar255.str = \"1.0\"", "iChar255.str = \"" + TextToString + "\" ");
                //MessageBox.Show(TextToChange);

                if (out_width < 13 && sw_case == 1) TextToChange = TextToChange.Replace("iDimDrawingParam.textPos = 0",
                    "iDimDrawingParam.textPos = 2");
                if (out_height > (-13) && sw_case == 2) TextToChange = TextToChange.Replace("iDimDrawingParam.textPos = 0",
                    "iDimDrawingParam.textPos = 2");

                text = text.Replace("#" + OriginalName, TextToChange);

                return text;
            }

            //Создание дефекта для отвода или прямой трубы
            if (radioButton4.Checked == true) //Прямая
            {
                ChangeText("Orig_defect", 5, "Error.");
                if (Convert.ToInt32(original_width) != 0 && Convert.ToDouble(original_width) != width_scheme)
                {
                    ChangeText("prodol", 1, original_width);
                }

                // Параметры для окружной выноски
                if (a != 0)
                {
                    ChangeText("okruj", 2, original_height);
                }
                Console.WriteLine(x);
                if (x != 0)
                {
                    ChangeText("do_def", 3, x.ToString());
                }
                ChangeText("Sliced_pipe", 11, x.ToString());
            }
            else if (radioButton4.Checked != true)
            {
                ChangeText("Elbow", 6, "Error :)"); //Отвод
                Sh_x1 = x1_cos_elbow;
                Sh_y1 = y1_sin_elbow;

                Sh_x2 = x2_cos_elbow;
                Sh_y2 = y2_sin_elbow;

                if (a != 0)
                {
                    ChangeText("okruj", 9, original_height);
                }

                if (x != 0)
                {
                    ChangeText("elbow_do_def", 7, x.ToString());
                }

                if (Convert.ToInt32(original_width) != 0)
                {
                    ChangeText("dl_def_elb", 8, original_width);
                }
            }

            // Параметры для продольной выноски

            //x1 - полочка | х2 - линия-выноска

            string defecto = File.ReadAllText(path + @"\defecto.txt", System.Text.Encoding.GetEncoding(1251));
            string defecto_second = File.ReadAllText(path + @"\defecto_second.txt", System.Text.Encoding.GetEncoding(1251));

            string do_def_func(string deflex, string TextToChange, int pos)
            {
                deflex = deflex.Replace("\"Дефект №5\"", "\"Дефект №" + defect + "\"");

                deflex = deflex.Replace("0.00", b_orig.ToString());
                deflex = deflex.Replace("0.01", c_orig.ToString());

                if (pos == 1)
                {
                    Console.WriteLine(x1_sin_def);
                    Console.WriteLine(y1_cos_def);
                    //Привязки линии-выноски
                    deflex = deflex.Replace("x = -111", "x = " + x1_sin_def);
                    deflex = deflex.Replace("y = -130", "y = " + y1_cos_def);//iMathPointParam
                                                                             //С какой стороны будет линия-выноска
                    if (cfd_orig <= 90)
                    {
                        //Console.WriteLine(1);
                        deflex = deflex.Replace("x = -222", "x = " + (x1_sin_def + 15));
                        deflex = deflex.Replace("y = -135.0", "y = " + (y1_cos_def + 10));//iLeaderParam
                    }
                    else if (cfd_orig > 90 && cfd_orig <= 180)
                    {
                        //Console.WriteLine(2);
                        deflex = deflex.Replace("x = -222", "x = " + (x1_sin_def + 15));
                        deflex = deflex.Replace("y = -135.0", "y = " + (y1_cos_def - 15));
                    }
                    else if (cfd_orig > 180 && cfd_orig <= 270)
                    {
                        //Console.WriteLine(3);
                        deflex = deflex.Replace("x = -222", "x = " + (x1_sin_def - 15));
                        deflex = deflex.Replace("y = -135.0", "y = " + (y1_cos_def - 15));

                        deflex = deflex.Replace("iLeaderParam.dirX = 1", "iLeaderParam.dirX = -1");
                    }
                    else
                    {
                        //Console.WriteLine(4);
                        deflex = deflex.Replace("x = -222", "x = " + (x1_sin_def - 15));
                        deflex = deflex.Replace("y = -135.0", "y = " + (y1_cos_def + 10));
                        deflex = deflex.Replace("iLeaderParam.dirX = 1", "iLeaderParam.dirX = -1");
                    }
                }
                if (pos == 2)
                {
                    if (radioButton4.Checked == true)
                    {
                        deflex = deflex.Replace("x = -111", "x = " + (out_width + pre_width));
                        deflex = deflex.Replace("y = -130", "y = " + (out_height + pre_height));//iMathPointParam

                        deflex = deflex.Replace("x = -222", "x = " + (out_width + pre_width + 15));
                        deflex = deflex.Replace("y = -135.0", "y = " + (out_height + pre_height - 15));//iLeaderParam
                    }
                    else
                    {
                        deflex = deflex.Replace("x = -111", "x = " + (x2_cos_elbow_second));
                        deflex = deflex.Replace("y = -130", "y = " + (y2_sin_elbow_second));//iMathPointParam

                        deflex = deflex.Replace("x = -222", "x = " + (x2_cos_elbow_second + 15));
                        deflex = deflex.Replace("y = -135.0", "y = " + (y2_sin_elbow_second - 15));//iLeaderParam
                    }
                }
                text = text.Replace("#" + TextToChange, deflex);
                return text;
            }

            do_def_func(defecto, "defecto", 1);
            do_def_func(defecto_second, "secdefecto", 2);

            File.WriteAllText(path + @"\Deflex_ch.cdm", text, Encoding.GetEncoding(1251));
            numericUpDown1.Value++;
        }

        //Математика расчета расположения дефекта по градусным мерам
        private void Counted(double b, double c, out double b_deg, out double c_deg)
        {
            b_deg = 0;
            c_deg = 180;
            int i = 2, j = 4;

            //Console.WriteLine("0.0) " + b + "; " + c);

            if (b == c && b != 0)
                while (i < j && b != 0)
                {
                    b--;
                    i++;
                }
            i = 2;
            while (i < j && c < 359)
            {
                c++;
                i++;
            }

            if (b < 180 && c <= 180 && c > b)
            {
                //Console.WriteLine(1);
                while (c - b < 6)
                {
                    c++;
                    if (c - b < 6 && b > 0) b--;
                }

                b_deg = b;
                c_deg = c;
            }
            else if (b < 180 && c <= 180 && c < b)
            {
                //Console.WriteLine(1/2);
                b_deg = 0;
                c_deg = 180;
            }
            else if (b <= 180 && c > 180)
            {
                while (Math.Abs(c - b) < 6 || Math.Abs(b - c) < 6)
                {
                    c++;
                    if (Math.Abs(c - b) < 6 || Math.Abs(b - c) < 6) b--;
                }
                //Console.WriteLine(2);
                c_deg = 180; // Нижняя точка
                             //Если одинаковый градус
                if (360 - c == b) b_deg = b;
                //Если b выше
                else if (b < 360 - c) b_deg = b;
                else if (b > 360 - c) b_deg = 360 - c; // 360 - c
            }
            else if (b >= 180 && c > 180 && c > b)
            {
                while (Math.Abs(c - b) < 6 || Math.Abs(b - c) < 6)
                {
                    c++;
                    if (Math.Abs(c - b) < 6 || Math.Abs(b - c) < 6 && b > 180) b--;
                }
                c_deg = 360 - b;
                b_deg = 360 - c;
            }
            else if (b >= 180 && c >= 180 && c < b)
            {
                b_deg = 0;
                c_deg = 180;
            }
            else if (b >= 180 && c <= 180)
            {
                while (Math.Abs(c - b) < 6 || Math.Abs(b - c) < 6)
                {
                    c++;
                    if (Math.Abs(c - b) < 6 || Math.Abs(b - c) < 6) b--;
                }
                b_deg = 0;
                if (360 - b == c) c_deg = c;
                else if (c < 360 - b) c_deg = 360 - b;
                else if (c > 360 - b) c_deg = c;
            }
            else if (b == c)
            {
                //Console.WriteLine(5);
                while (Math.Abs(b - c) < 6)
                {
                    c++;
                    b--;
                }
                if (b < 180 && c < 180)
                {
                    b_deg = b;
                    c_deg = c;
                }
                else if (b < 180 && c > 180)
                {
                    c_deg = 180;
                    if (b > (360 - c))
                    {
                        b_deg = b;
                    }
                    else b_deg = 360 - c;
                }
                else if (b > 180 && c > 180)
                {
                    b_deg = 360 - c;
                    c_deg = 360 - b;
                }
                else if (b > 180 && c < 180)
                {
                    b_deg = 0;
                    if (c > (360 - b))
                    {
                        c_deg = c;
                    }
                    else c_deg = 360 - b;
                }
            }
            else if (b >= 180 && c > 180 && c <= 359)
            {
                if (c <= 270)
                {
                    b_deg = c;
                    c_deg = b;
                }
            }
            else
            {
                MessageBox.Show("ERROR");
                b_deg = 0;
                c_deg = 180;
            }
            //Console.WriteLine("0.1) " + b + "; " + c);
        }

        public void Overload_Sin_Cos(out double x_cos, out double y_sin, double rad, double n_deg, double const_deg)
        {
            x_cos = rad * Math.Cos(DegreeToRadian(const_deg - n_deg));
            y_sin = rad * Math.Sin(DegreeToRadian(const_deg - n_deg));
        }

        public void Second_Prog()
        {
            //Функция создания продольных сварных швов
            int sect;
            sect = Convert.ToInt32(numericUpDown2.Value);

            Excel.Worksheet xlSheet_02 = (Excel.Worksheet)xlApp.Worksheets.get_Item(4);
            ((Excel.Worksheet)this.xlApp.ActiveWorkbook.Sheets[4]).Select();
            Excel.Range xlRange_02 = xlApp.get_Range("C2", $"C{xlSheet_02.UsedRange.Rows.Count}");

            Excel.Range Sect_find = xlRange_02.Find(sect); //Поиск по клетке
            Excel.Range Pipe_length = xlSheet_02.Cells[Sect_find.Row, 5] as Excel.Range;
            Excel.Range Diameter = xlSheet_02.Cells[Sect_find.Row, 7] as Excel.Range;
            Excel.Range Thickness = xlSheet_02.Cells[Sect_find.Row, 6] as Excel.Range;

            try
            {
                double Pipe_length_do_inTry = Convert.ToDouble(Pipe_length.Value2);
                double Diameter_do_inTry = Convert.ToDouble(Diameter.Value2);
                double Thickness_do_inTry = Convert.ToDouble(Thickness.Value2);
            }
            catch
            {
                MessageBox.Show("Ошибка данных. Проверьте значения в Excel.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                textBox7.Text = Pipe_length.Value2.ToString();
                textBox2.Text = Diameter.Value2.ToString();
                textBox6.Text = Thickness.Value2.ToString();
                numericUpDown2.Value += 10;
                return;
            }

            double Pipe_length_do = Convert.ToDouble(Pipe_length.Value2);
            double Diameter_do = Convert.ToDouble(Diameter.Value2);
            double Thickness_do = Convert.ToDouble(Thickness.Value2);

            Pipe_length_do *= 1000;

            textBox7.Text = Pipe_length_do.ToString();
            textBox2.Text = Diameter_do.ToString();
            textBox6.Text = Thickness_do.ToString();

            double A_1, A_1_Orig, A_2_Orig, height_pipe, pipe_width, y, y_2; //A1 - Число для конвертации y1,2            
            height_pipe = Convert.ToDouble(textBox5.Text);
            pipe_width = Convert.ToDouble(textBox3.Text);
            //dataGridView1[0, 0].Value = range_def_do.ToString(); //Расстояние от начала

            if (textBox4.Text == "")
            {
                caseSwitch = 3;
                radioButton3.Checked = true;
            }

            if (checkBox4.Checked == true)
            {
                numericUpDown3.Value = numericUpDown2.Value;
            }

            if (textBox2.Text == "") textBox2.Text = "Empty";
            if (textBox7.Text == "") textBox7.Text = "9999";
            if (textBox6.Text == "") textBox6.Text = "9999";

            string pipe = File.ReadAllText(path + @"\Pipe.cdm", System.Text.Encoding.GetEncoding(1251)); //Общий файл
            string Pp_sh, Pp_sh_r;

            string T_cm = File.ReadAllText(path + @"\T_cm.txt", System.Text.Encoding.GetEncoding(1251)); //Толщина трубы
            string W_pp;

            if (radioButton4.Checked == true)
            {
                Pp_sh = File.ReadAllText(path + @"\Welds\Pp_sh.txt", System.Text.Encoding.GetEncoding(1251)); //Кольцевой шов (левый)
                Pp_sh_r = File.ReadAllText(path + @"\Welds\Pp_sh_r.txt", System.Text.Encoding.GetEncoding(1251)); //Кольцевой шов (правый);
                W_pp = File.ReadAllText(path + @"\W_pp.txt", System.Text.Encoding.GetEncoding(1251)); //Высота трубы
            }
            else
            {
                Pp_sh = File.ReadAllText(path + @"\Welds\Pp_sh_elbow.txt", System.Text.Encoding.GetEncoding(1251)); //Кольцевой шов (левый)
                Pp_sh_r = File.ReadAllText(path + @"\Welds\Pp_sh_elbow_r.txt", System.Text.Encoding.GetEncoding(1251)); //Кольцевой шов (правый);
                W_pp = File.ReadAllText(path + @"\W_elbow_pp.txt", System.Text.Encoding.GetEncoding(1251)); //Высота трубы
            }
            if(caseSwitch != 3)
            {
            if (radioButton4.Checked == true)
            {

                string F_sh = File.ReadAllText(path + @"\F_sh.txt", System.Text.Encoding.GetEncoding(1251));
                string S_sh = File.ReadAllText(path + @"\S_sh.txt", System.Text.Encoding.GetEncoding(1251));
                
                int visible_def = 1;
                A_1 = Convert.ToDouble(textBox4.Text); // Градусы
                double A_2 = A_1 + 180;
                A_1_Orig = A_1;

                y = -(A_1 / 180 * height_pipe);
                y = Math.Round(y, 2);


                if (A_1 > 180)
                {
                    A_1 = 360 - A_1;
                    visible_def = 9;
                }
                if (caseSwitch == 1)
                {
                    //Прога рисует 1 продольный св шов в диапазоне 0-360
                    F_sh = F_sh.Replace("(0.0, -35.0, 205.0, -35.0, 9)",
                        "( 0.0, " + y + ", " + pipe_width + ", " + y + ", " + visible_def + ")");
                    F_sh = F_sh.Replace("iTextItemParam.s = \"10/242\"",
                        "iTextItemParam.s = \"" + numericUpDown2.Value + "/" + A_1_Orig + "\"");

                    //iParagraphParam.x = 175.0

                    F_sh = F_sh.Replace(".y = 0.0",
                        ".y = " + (y + 4));
                    F_sh = F_sh.Replace("iParagraphParam.x = 175.0",
                        "iParagraphParam.x = " + (pipe_width - 27));

                    pipe = pipe.Replace("#F_sh", F_sh);
                }
                else if (caseSwitch == 2)
                {
                    //Прога рисует 2 продольных св швв в диапазоне 0-180 и 180-360 соответственно
                    

                    A_2 = 180 - A_1;
                    A_2_Orig = 360 - A_2;

                    y_2 = -(A_2 / 180 * height_pipe);
                    y_2 = Math.Round(y_2, 2);

                    F_sh = F_sh.Replace("(0.0, -35.0, 205.0, -35.0, 9)",
                        "( 0.0, " + y + ", " + pipe_width + ", " + y + ", 1)");
                    F_sh = F_sh.Replace("iTextItemParam.s = \"10/242\"",
                        "iTextItemParam.s = \"" + numericUpDown2.Value + "/" + A_1_Orig + "\"");
                    F_sh = F_sh.Replace(".y = 0.0",
                        ".y = " + (y - 8));
                    F_sh = F_sh.Replace("iParagraphParam.x = 175.0",
                        "iParagraphParam.x = " + (pipe_width - 27));

                    S_sh = S_sh.Replace("(0.0, -35.0, 205.0, -35.0, 9)",
                        "( 0.0, " + y_2 + ", " + pipe_width + ", " + y_2 + ", 9)");
                    S_sh = S_sh.Replace("iTextItemParam.s = \"10/242\"",
                        "iTextItemParam.s = \"" + numericUpDown2.Value + "/" + A_2_Orig + "\"");
                    S_sh = S_sh.Replace(".y = 0.0",
                        ".y = " + (y_2 + 4));
                    S_sh = S_sh.Replace("iParagraphParam.x = 175.0",
                        "iParagraphParam.x = " + (pipe_width - 27));

                    pipe = pipe.Replace("#F_sh", F_sh);
                    pipe = pipe.Replace("#S_sh", S_sh);
                }
            }            
            else if (radioButton6.Checked == true || radioButton5.Checked == true)
            {

                string F_sh = File.ReadAllText(path + @"\B_sh_elb.txt", System.Text.Encoding.GetEncoding(1251));
                string S_sh = File.ReadAllText(path + @"\B_sh_elb.txt", System.Text.Encoding.GetEncoding(1251));
                
                int visible_def = 1;
                int visible_def_second;                
                A_1 = Convert.ToDouble(textBox4.Text); // Градусы
                double A_2 = A_1 + 180;
                A_1_Orig = A_1;
                A_2_Orig = A_1 + 180;

                if (radioButton6.Checked == true)
                {


                    if (A_1 > 180)
                    {
                        A_1 = 360 - A_1;
                        visible_def = 9;
                    }
                    y = 140 - (A_1 / 180 * height_pipe) / 2;

                    y = Math.Round(y, 2);
                    Math.Abs(y);

                    y_2 = 140 - ((180 - A_1) / 180 * height_pipe) / 2;

                    if (visible_def == 1) visible_def_second = 9;
                    else visible_def_second = 1;
                }
                else
                {
                    //1 гр = 0,38(8)
                    if (A_1 > 90 && A_1 <= 270)
                    {
                        A_1 = 180 - A_1;
                        visible_def = 9;
                    }
                    Console.WriteLine(A_1);
                    y = 140 * 0.75 - (A_1 / 180 * height_pipe) / 2;
                    Console.WriteLine(y);
                    if (A_1 > 90 && A_1 < 180) y += 35;
                    if (A_1 >= 180 && A_1 <= 270) y += 70;
                    if (A_1 > 270) y += 140;
                    
                    
                    Console.WriteLine(y+"\n");


                    if (caseSwitch == 2)
                    {

                        if (A_2 > 90 && A_2 <= 270)
                        {
                            A_2 = 180 - A_2;
                        }
                        y_2 = 140 * 0.75 - (A_2 / 180 * height_pipe) / 2;
                        if (A_2 >= 180 && A_2 <= 270) y_2 += 70;
                        if (A_2 > 270) y_2 += 140;
                    }
                    else
                    {
                        y_2 = 0;
                    }

                    y = Math.Round(y, 2);
                    Math.Abs(y);
                    
                    y_2 = Math.Round(y_2, 2);
                    Math.Abs(y_2);

                    if (visible_def == 1) visible_def_second = 9;
                    else visible_def_second = 1;
                }

                
                if (caseSwitch == 1)
                {
                    F_sh = F_sh.Replace("#F_sh", $"iDocument2D.ksArcByPoint(0.0, 0.0, {y}, {y}, {0.0}, {0.0}, {y}, 1, {visible_def})");//R, x1, y1, x2, y2
                    F_sh = F_sh.Replace("#a ", "");
                    F_sh = F_sh.Replace("76.32", $"{y-11}");
                    
                    F_sh = F_sh.Replace("40/225", $"{numericUpDown2.Value}/{A_1_Orig}");
                    pipe = pipe.Replace("#F_sh", F_sh);
                }
                else if(caseSwitch == 2)
                {
                    F_sh = F_sh.Replace("#F_sh", $"iDocument2D.ksArcByPoint(0.0, 0.0, {y}, {y}, {0.0}, {0.0}, {y}, 1, {visible_def})");//R, x1, y1, x2, y2
                    F_sh = F_sh.Replace("#S_sh", $"iDocument2D.ksArcByPoint(0.0, 0.0, {y_2}, {y_2}, {0.0}, {0.0}, {y_2}, 1, {visible_def_second})");//R, x1, y1, x2, y2

                    F_sh = F_sh.Replace("#a ", "");
                    S_sh = S_sh.Replace("#b ", "");

                    if (radioButton6.Checked == true)
                    {
                        if (A_1 <= 90)
                        {
                            F_sh = F_sh.Replace("40/225", $"{numericUpDown2.Value}/{A_2_Orig}");
                            S_sh = S_sh.Replace("60/335", $"{numericUpDown2.Value}/{A_1_Orig}");
                        }
                        else
                        {
                            F_sh = F_sh.Replace("40/225", $"{numericUpDown2.Value}/{A_1_Orig}");
                            S_sh = S_sh.Replace("60/335", $"{numericUpDown2.Value}/{A_2_Orig}");
                        }
                    }
                    else
                    {
                        if (A_1 <= 180)
                        {
                            F_sh = F_sh.Replace("40/225", $"{numericUpDown2.Value}/{A_1_Orig}");
                            S_sh = S_sh.Replace("60/335", $"{numericUpDown2.Value}/{A_2_Orig}");
                        }
                        else
                        {
                            F_sh = F_sh.Replace("40/225", $"{numericUpDown2.Value}/{A_2_Orig}");
                            S_sh = S_sh.Replace("60/335", $"{numericUpDown2.Value}/{A_1_Orig}");
                        }
                    }
                    
                    pipe = pipe.Replace("#F_sh", F_sh);
                    pipe = pipe.Replace("#S_sh", S_sh);
                }
            }
            }
            int curr_incr_num = 10;

            numericUpDown3.Increment = curr_incr_num;

            //Отрисовка кольцевых швов
            if (checkBox1.Checked == true)
            {
                checkBox2.Enabled = true;
            }

            if (checkBox1.Checked == true && checkBox2.Checked == false)
            {
                Pp_sh = Pp_sh.Replace("К23",
                "K" + numericUpDown3.Value);
                pipe = pipe.Replace("#Pp_sh", Pp_sh);
            }
            else if (checkBox2.Checked == true)
            {
                Pp_sh_r = Pp_sh_r.Replace("К24",
                "K" + (Convert.ToInt32(numericUpDown3.Value)));
                Pp_sh_r = Pp_sh_r.Replace("iParagraphParam.x = 210.0",
                    "iParagraphParam.x = " + (pipe_width + 8))
                    ;
                pipe = pipe.Replace("#Pp_2sh_r", Pp_sh_r);
            }
            else
            {
                Pp_sh = Pp_sh.Replace("К23",
                "K" + numericUpDown3.Value);
                Pp_sh_r = Pp_sh_r.Replace("К24",
                "K" + (Convert.ToInt32(numericUpDown3.Value) + curr_incr_num));
                Pp_sh_r = Pp_sh_r.Replace("iParagraphParam.x = 210.0",
                    "iParagraphParam.x = " + (pipe_width + 8));

                pipe = pipe.Replace("#Pp_sh", Pp_sh);
                pipe = pipe.Replace("#Pp_2sh_r", Pp_sh_r);
            }

            //Выноска с номером и толщиной трубы
            T_cm = T_cm.Replace("15,7", textBox2.Text);
            T_cm = T_cm.Replace("Секция №10", "Секция №" + numericUpDown2.Value);

            if (radioButton4.Checked == true)
            {
                // Leader - Выноска
                // MathPoint - Линия-указатель
                T_cm = T_cm.Replace("iLeaderParam.x = 157",
                        "iLeaderParam.x = " + (pipe_width + 7));
                T_cm = T_cm.Replace("iLeaderParam.y = 30",
                        "iLeaderParam.y = " + (-(height_pipe + 15)));
                T_cm = T_cm.Replace("iMathPointParam.x = 139",
                        "iMathPointParam.x = " + (pipe_width - 2));
                T_cm = T_cm.Replace("iMathPointParam.y = 16",
                        "iMathPointParam.y = " + (-(height_pipe)));
            }

            pipe = pipe.Replace("#T_cm", T_cm);

            W_pp = W_pp.Replace("11505", textBox7.Text);

            W_pp = W_pp.Replace("iLDimSourceParam.x2 = 205.0",
                    "iLDimSourceParam.x2 = " + (pipe_width));

            string diam = File.ReadAllText(path + @"\Diameter.txt", System.Text.Encoding.GetEncoding(1251));
            if (radioButton4.Checked == false)
            {
                diam = diam.Replace("-11.0", "-20"); //dx
                diam = diam.Replace("0.1", "97.5"); //dy
                diam = diam.Replace("-47.0", "-28"); //x1
                diam = diam.Replace("0.2", "140"); //y1
                diam = diam.Replace("-45.0", "-27.5"); //x2
                diam = diam.Replace("-70.0", "70"); //y2
            }
            diam = diam.Replace("1220", textBox6.Text);
            pipe = pipe.Replace("#diam", diam);
            pipe = pipe.Replace("#W_pp", W_pp);
            textBox1.Text = textBox7.Text;

            //Указывать только 1 поперечный сварной шов
            File.WriteAllText(path + @"\Pipe_ch.cdm", pipe, Encoding.GetEncoding(1251));
            numericUpDown2.Value += 10;
            numericUpDown3.Value += curr_incr_num;
        }

        private double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Main_prog();
        }

        private int caseSwitch = 1;

        private void Button4_Click(object sender, EventArgs e)
        {
            Second_Prog();
        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            textBox4.Enabled = true;
            caseSwitch = 1;
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            textBox4.Enabled = true;
            caseSwitch = 2;
        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            textBox4.Enabled = false;
            textBox4.Text = "";
            caseSwitch = 3;
        }

        private void RadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            textBox3.Text = "205"; //Длина трубы на чертеже
            textBox5.Text = "70"; //Высота трубы на чертеже
        }

        private void RadioButton6_CheckedChanged(object sender, EventArgs e)
        {
            textBox3.Text = "140"; //Длина трубы на чертеже
            textBox5.Text = "140"; //Высота трубы на чертеже
        }

        private void RadioButton5_CheckedChanged(object sender, EventArgs e)
        {
            textBox3.Text = "140"; //Длина трубы на чертеже
            textBox5.Text = "140"; //Высота трубы на чертеже
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            toolStripMenuItem1_Click(null, null);
            Close();
        }

        private void OpenExcelToolStripMenuItem_Click(object sender, EventArgs e)//Открыть Excel
        {
            openFileDialog1.Filter = "Excel Files(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm|All files(*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            button6.Enabled = true;
            string pathToXlsx = openFileDialog1.FileName;
            //string pathToXlsx = filename;

            try
            {// Присоединение к открытому приложению Excel (если оно открыто)
                xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                flagexcelapp = 1; // устанавливаем флаг в 1, будем знать что присоединились
            }
            catch
            {
                xlApp = new Excel.Application();// Если нет, то создаём новое приложение
            }
            finally
            {
                xlApp.Workbooks.Open(Path.GetFullPath(pathToXlsx));
                xlAppBooks = xlApp.Workbooks; // Получаем список открытых книг
                xlAppBook = xlAppBooks[xlAppBooks.Count];
                xlSheets = xlAppBook.Worksheets;
            }

            tabControl1.SelectTab(tabPage1);
            OpenExcelToolStripMenuItem.Enabled = false;
            toolStripMenuItem1.Enabled = true;
            tabControl1.Enabled = true;
            label13.Text = "Excel подключен";
            label13.ForeColor = System.Drawing.Color.Green;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            button6.Enabled = false;

            if (flagexcelapp == 0)
            {
                xlAppBook.Close(false, false, false);
                xlApp.Quit();
                Process[] List;
                List = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in List)
                {
                    proc.Kill();
                }
            }
            else
            {
                xlAppBook.Close(false, false, false);
            }
            OpenExcelToolStripMenuItem.Enabled = true;
            toolStripMenuItem1.Enabled = false;
            tabControl1.Enabled = false;
            label13.Text = "Excel отключен";
            label13.ForeColor = System.Drawing.Color.OrangeRed;
        }

        private void AboutProgramToolStripMenuItem_Clicked(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Автор: Мазиков А.С." +
                "\nВерсия: " + strVersion,
                "О программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void HelperToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("file:///" + Application.StartupPath.ToString() + "/Resources/Help/Help.chm");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка, файл справки не найден",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            //txt.Start();
        }
    }
}

