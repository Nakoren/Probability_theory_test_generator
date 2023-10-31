using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Status;
using static Test.Form1;

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Status_TB.Hide();
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Add(Visible: true);
            try
            {
                var json = File.ReadAllText("practic.json"); // считываем файл json

                var data = JsonConvert.DeserializeObject<Data>(json); // десериализуем данные

                var tasksById = data.Types // группируем задачи по номеру задачи
                    .SelectMany(t => t.Tasks)
                    .GroupBy(task => task.Id);

                List<List<Task>> variants = new List<List<Task>>();
                for (int i = 0; i < 4; i++)
                {
                    variants.Add(new List<Task>(new Task[tasksById.Count()]));
                }

                Random rand = new Random();
                foreach (var tasks in tasksById) // выводим задачи в консоль
                {
                    foreach (var task in tasks)
                    {
                        int var = rand.Next(0, 4);
                        while (variants[var][task.Id - 1] != null)
                        {
                            var = rand.Next(0, 4);
                        }
                        variants[var][task.Id - 1] = task;
                    }
                }
                
                int varKol = Convert.ToInt32(numericUpDown1.Value);
                if (varKol <= 0)
                {
                    throw new IndexOutOfRangeException();
                }

                //БЛОК ВЫВОДА В WORD


                string check = null;
                var appDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                Range doc_range = doc.Range();
                for (int i = 0; i < varKol; i++)
                {

                   
                    Paragraph var_id = doc.Paragraphs.Add();
                    Range var_id_range = var_id.Range;
                    var_id_range.Text = $"Тест 2. Вариант {i + 1}:\n";
                    var_id_range.InsertParagraphAfter();
                    var_id_range.Bold = 1;
                    var_id_range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    Paragraph title = doc.Paragraphs.Add();
                    Range title_r = title.Range;
                    title_r.Text = "Фамилия____________________________________ Группа_______________";
                    title_r.InsertParagraphAfter();
                    title_r.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;


                    foreach (Task task in variants[i % 4])
                    {
                        Paragraph task_ID = doc.Paragraphs.Add();
                        Range task_id_range = task_ID.Range;
                        task_id_range.Bold = 1;
                        task_id_range.Text = $"Задание {task.Id}:";
                        task_id_range.InsertParagraphAfter();

                        Paragraph task_1_text = doc.Paragraphs.Add();
                        Range task_1_text_r = task_1_text.Range;
                        task_1_text_r.Text = task.Text_1;
                        task_1_text_r.InsertParagraphAfter();
                        

                        if (task.ImagesSource != "")
                        {
                            Paragraph task_image = doc.Paragraphs.Add();
                            Range task_image_r = task_image.Range;
                            float width = 1;
                            float height = 1;

                            if (task.Id == 3)
                            {
                                width = (float)55;
                                height = (float)55;
                            }

                            if (task.Id == 4)
                            {
                                width = (float)65;
                                height = (float)65;
                            }
                            if (task.Id == 5)
                            {
                                width = (float)75;
                                height = (float)75;
                            }
                            if (task.Id == 6)
                            {
                                width = (float)75;
                                height = (float)75;
                            }
                            if (task.Id == 7)
                            {
                                width = (float)60;
                                height = (float)60;
                            }
                            if (task.Id == 8)
                            {
                                width = (float)75;
                                height = (float)75;
                            }

                            var relativePath = "content\\" + task.ImagesSource;
                            var fullPath = Path.Combine(appDir, relativePath);
                            var image = task_image_r.InlineShapes.AddPicture(fullPath);
                            image.ScaleWidth = width;
                            image.ScaleHeight = height;

                            task_image_r.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            //image.Width = height;
                            //image.Height = width;

                        }
                        


                        if (task.Text_2 != "")
                        {
                            Paragraph task_2_text = doc.Paragraphs.Add();
                            Range task_2_text_r = task_2_text.Range;
                            check = task.Text_2;
                            task_2_text_r.Text = task.Text_2;
                            check = task_2_text_r.Text;
                            task_2_text_r.InsertParagraphAfter();
                            check = task_2_text_r.Text;
                        }
                        
                        for (int k = task.Answers.Count - 1; k >= 1; k--)
                        {
                            int j = rand.Next(k + 1);
                            string temp = task.Answers[j];
                            task.Answers[j] = task.Answers[k];
                            task.Answers[k] = temp;
                            
                        }

                        if (!(((task.Id == 6) || (task.Id == 5)) && ((task.Var == 3) || (task.Var == 4))))
                        {
                            int abc = 1;
                            Paragraph par_answers = doc.Paragraphs.Add();
                            Range par_answers_r = par_answers.Range;
                            par_answers_r.Text = "";
                            int count = 1;
                            foreach (var answer in task.Answers)
                            {

                                par_answers_r.Text += $"{(char)('а' + abc - 1)}){answer} ";
                                if ((count == 2) && ((task.Id == 15) || (task.Id == 20) || (task.Id == 17) || (task.Id == 12) || (task.Id == 11) || (task.Id == 10) || (task.Id == 13) || (task.Id == 16)))
                                {
                                    par_answers_r.Text += '\n';
                                }


                                if (answer == task.Correct_Answer && i < 4)
                                    task.Correct_Answer = (char)('а' + abc - 1) + ") " + task.Correct_Answer;
                                else if (answer == task.Correct_Answer.Substring(task.Correct_Answer.IndexOf(" ") + 1) && i >= 4)
                                    task.Correct_Answer = (char)('а' + abc - 1) + ") " + task.Correct_Answer.Substring(task.Correct_Answer.IndexOf(" ") + 1);
                                abc++;
                                count++;
                            }
                            par_answers_r.Text += '\r';
                            par_answers_r.InsertParagraphAfter();
                            par_answers_r.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        }
                        else
                        {

                            int abc = 1;
                            int count = 1;
                            foreach (var answer in task.Answers)
                            {
                                
                                Paragraph par_answer_t = doc.Paragraphs.Add();
                                Range par_answer_t_r = par_answer_t.Range;
                                par_answer_t_r.Text = $"{(char)('а' + abc - 1)})";
                                par_answer_t_r.InsertParagraphAfter();

                                Paragraph par_answer_pic = doc.Paragraphs.Add();
                                Range par_answer_pic_r = par_answer_pic.Range;
                                var relativePath = "content\\" + answer;
                                var fullPath = Path.Combine(appDir, relativePath);
                                var image = par_answer_pic_r.InlineShapes.AddPicture(fullPath);
                                image.ScaleWidth = 50;
                                image.ScaleHeight = 50;



                                if (answer == task.Correct_Answer && i < 4)
                                    task.Correct_Answer = task.Correct_Answer;
                                else if (answer == task.Correct_Answer.Substring(task.Correct_Answer.IndexOf(" ") + 1) && i >= 4)
                                    task.Correct_Answer = task.Correct_Answer.Substring(task.Correct_Answer.IndexOf(" ") + 1);
                                abc++;
                                count++;
                            }
                        }


                       


                    }
                    Paragraph PageBr_par_1 = doc.Paragraphs.Add();
                    PageBr_par_1.Range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                    Paragraph par_cor_ans = doc.Paragraphs.Add();
                    Range par_cor_ans_r = par_cor_ans.Range;
                    par_cor_ans_r.Text = "Ответы для варианта " + (i+1);
                    par_cor_ans_r.InsertParagraphAfter();
                    par_cor_ans_r.Bold = 1;

                    int counter = 1;
                    foreach (Task task in variants[i % 4])
                    {
                        if (((task.Id == 6) || (task.Id == 5)) && ((task.Var == 3) || (task.Var == 4)))
                        {
                            Paragraph par_temp = doc.Paragraphs.Add();
                            Range par_temp_r = par_temp.Range;
                            par_temp_r.Text = counter + ")";
                            par_temp_r.InsertParagraphAfter();

                            Paragraph par_answer_pic = doc.Paragraphs.Add();
                            Range par_answer_pic_r = par_answer_pic.Range;
                            var relativePath = "content\\" + task.Correct_Answer;
                            var fullPath = Path.Combine(appDir, relativePath);
                            var image = par_answer_pic_r.InlineShapes.AddPicture(fullPath);
                            image.ScaleWidth = 60;
                            image.ScaleHeight = 60;
                        }
                        else
                        {
                            Paragraph par_temp = doc.Paragraphs.Add();
                            Range par_temp_r = par_temp.Range;
                            par_temp_r.Text = counter + ") " + task.Correct_Answer;
                            par_temp_r.InsertParagraphAfter();
                        }
                        counter++;
                        
                    }

                    Paragraph PageBr_par_2 = doc.Paragraphs.Add();
                    PageBr_par_2.Range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                }
                
                doc_range.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle;
                doc_range.ParagraphFormat.SpaceAfter = 0;
                doc_range.ParagraphFormat.SpaceBefore = 0;
                doc_range.Font.Size = 12;

                Status_TB.Text = "Генерация тестов прошла успешно.";
                Status_TB.ForeColor = Color.Blue;
                Status_TB.Show();

                doc.Save();
                doc.Close();
                app.Quit();
            }
            catch (IndexOutOfRangeException)
            {
                MessageBox.Show("Необходимо указать хоть 1 вариант");
            }
            catch
            {
                Status_TB.Text = "Произошёл сбой программы.\n Убедитесь в том, что у вас отсутствуют другие активные процессы word";
                Status_TB.ForeColor = Color.Red;
                Status_TB.Show();
                app.Quit();
            }
        }

        public class Task
        {
            public int Id { get; set; }
            public int Var { get; set; }
            public string Text_1 { get; set; }
            public string ImagesSource { get; set; }
            public string Text_2 { get; set; }
            public string Correct_Answer { get; set; }
            public List<string> Answers { get; set; }
        }
        public class Type
        {
            public List<Task> Tasks { get; set; }
        }
        public class Data
        {
            public List<Type> Types { get; set; }
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Information_form Inf_f = new Information_form();
            Inf_f.Show();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}