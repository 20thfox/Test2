using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop;

namespace Test2
{
    public partial class Form1 : Form
    {
        private Word.Word.Application wordapp;
        private Word.Word.Document worddocument;
        private int NumOfProt = 0;
        private string SaveName;
        private bool GeneralFault = false;

        private Object trueObj = true;
        private Object falseObj = false;
        Object missingObj = System.Reflection.Missing.Value;
        public Form1()
        {
            InitializeComponent();

        }

        /*private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            wordapp = new Word.Word.Application();
            wordapp.Visible = true;
            Object template = @"C:\111.docx"; //Type.Missing;
            Object newTemplate = false;
            Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
            Object visible = true;
            //Object begin = 538; предполагаемое место Присоединения
            //Object end = 540;
            worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
            //Word.Word.Range range = worddocument.Range(ref begin, ref end);
            //range.Select();

        } */ //Открытие документа по верхнему меню.

        private void button1_Click(object sender, EventArgs e)
        {
            wordapp = new Word.Word.Application
            { Visible = false };

            if (checkedListBox1.GetItemChecked(0) == true)
            {
                NumOfProt++;
                SaveName = "Вторичная коммутация";
                GenVtorCom();
                progressBar1.Value++;
            }
            if(checkedListBox1.GetItemChecked(1) == true)
            {
                NumOfProt++;
                SaveName = "Металлосвязь";
                GenGround();
                progressBar1.Value++;
            }
            if(checkedListBox1.GetItemChecked(2) == true)
            {
                NumOfProt++;
                SaveName = "Электродвигатели";
                GenEngine();
                progressBar1.Value++;
            }
            if(checkedListBox1.GetItemChecked(3) == true)
            {
                NumOfProt++;
                SaveName = "Параметрирование ПЛК";
                GenPLC();
                progressBar1.Value++;
            }
            if(checkedListBox1.GetItemChecked(4) == true)
            {
                NumOfProt++;
                SaveName = "Кабельные линии";
                CabLine();
                progressBar1.Value++;
            }
            if(checkedListBox1.GetItemChecked(5) == true)
            {
                NumOfProt++;
                SaveName = "Испытание контрольных кабельных линий";
                CheckCabLine();
                progressBar1.Value++;
            }

            if (GeneralFault == false)
            {
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                NumOfProt = 0;
                MessageBox.Show("Завершено успешно");
            }

        }

        private void GenVtorCom()
        {
            try
            {
                Object template = @"C:\Users\Twent\Desktop\Templates\VtorCom.docx";
                Object newTemplate = false;
                Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;


                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                GenFormat();

                //Чтото происходит

                Save();

            }
            catch (Exception)
            {
                wordapp.Quit(ref falseObj, ref missingObj,ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
            }

        }
        private void CabLine()
        {
            try
            {
                Object template = @"C:\Users\Twent\Desktop\Templates\CabLine1.docx";
                Object newTemplate = false;
                Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;



                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible); //создание объекта документа по шаблону
                GenFormat();

                Save();

                //
                //вывод таблицы из проги и вставка в документ
                //резерв table2
            }
            catch (Exception)
            {
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
            }

        }
        private void GenPLC() {
            try
            {
                Object template = @"C:\Users\Twent\Desktop\Templates\PLC.docx";
                Object newTemplate = false;
                Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;




                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                GenFormat();

                //Чтото происходит

                Save();

            }
            catch (Exception)
            {
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
            }

        }
        private void GenGround() {
            try
            {
                Object template = @"C:\Users\Twent\Desktop\Templates\Ground.docx";
                Object newTemplate = false;
                Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;




                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                GenFormat();

                //Чтото происходит

                Save();

            }
            catch (Exception)
            {
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
            }
        }
        private void GenEngine() {
            try
            {
                Object template = @"C:\Users\Twent\Desktop\Templates\Engine.docx";
                Object newTemplate = false;
                Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;




                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                GenFormat();

                //Чтото происходит

                Save();

            }
            catch (Exception)
            {
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
            }
        }
        private void CheckCabLine() {
            try
            {
                Object template = @"C:\Users\Twent\Desktop\Templates\CheckCabLine.docx";
                Object newTemplate = false;
                Object documentType = Word.Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;




                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                GenFormat();

                //Чтото происходит

                Save();

            }
            catch (Exception)
            {
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
            }
        }

        private void GenFormat()
        {
            //Объявление всякой хрени
            Object findText;
            Object replaceText;
            //
            //Замена объекта и присоединения
            //
            Word.Word.Table table1 = worddocument.Tables[1]; //Обращение к таблице по индексу 1
            table1.Cell(2, 4).Range.InsertAfter(textBox1.Text); //вставка значения поля в ячейку таблицы
            table1.Cell(4, 4).Range.InsertAfter(textBox2.Text);
            //
            //Замента номера протокола, температуры, давления и влаги
            //тут будет чтото с FIND или нет
            findText = "п00-0-0-0000";
            replaceText = textBox3.Text + "-" + NumOfProt;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@Temp";
            replaceText = textBox4.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@Pres";
            replaceText = textBox5.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@Vlag";
            replaceText = textBox6.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);

            //
            //замена нижнего колонтитула
            //резерв table 3 
            foreach (Word.Word.Section sec in worddocument.Sections)
            {
                var range = sec.Footers[Word.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                Word.Word.Table table3 = range.Tables[1];
                table3.Cell(1, 1).Range.InsertAfter(textBox3.Text + "-" + NumOfProt);
            }

            //
            //Испытания произвели, фамилии даты и прочее
            //

            var countTabl = worddocument.Tables.Count;
            Word.Word.Table lastTable = worddocument.Tables[countTabl];
            lastTable.Cell(1, 2).Range.InsertAfter(comboBox1.Text);
            lastTable.Cell(3, 2).Range.InsertAfter(comboBox2.Text);
            lastTable.Cell(1, 4).Range.InsertAfter("/ " + textBox7.Text + " /");
            lastTable.Cell(3, 4).Range.InsertAfter("/ " + textBox8.Text + " /");
            lastTable.Cell(6, 4).Range.InsertAfter("/ " + textBox9.Text + " /");
            lastTable.Cell(5, 2).Range.InsertAfter(textBox10.Text);
            lastTable.Cell(7, 2).Range.InsertAfter(textBox11.Text);

        }
        private void Save()
        {
            
            //здесь будет сохранение
            Object fileName = @"C:\Users\Twent\Desktop\TEST2\" + textBox3.Text + "-" + NumOfProt + " " + SaveName + ".docx";
            Object fileFormat = Word.Word.WdSaveFormat.wdFormatDocumentDefault;
            worddocument.SaveAs2(ref fileName, ref fileFormat);
            worddocument.Close(ref falseObj, ref missingObj, ref missingObj);

            worddocument = null;
            //wordapp = null;
            //Вывод сообщения?
            //label27.Text = "Завершено"; //это ваще потом сделать 


        }
        private void genFaultActive()
        {
            GeneralFault = true;
            MessageBox.Show("Чтото пошло не так");
        }
    }
}
