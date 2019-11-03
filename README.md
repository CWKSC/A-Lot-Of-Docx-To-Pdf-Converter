# A-Lot-Of-Docx-To-Pdf-Converter

a lot of docx to pdf converter

[C# 笔记 - 批量 docx 到 pdf 转换器 - 知乎](#https://zhuanlan.zhihu.com/p/89958561)

Main code:

```C#
using Microsoft.Office.Interop.Word;
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace WordToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        BackgroundWorker[] backgroundWorkers;
        int totalJob;
        int completedJobNum;

        private void WordToPDF(object sender, DoWorkEventArgs e)
        {
            string sourcePath = ((string[]) e.Argument)[0];
            string targetPath = ((string[]) e.Argument)[1];

            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null;
            try
            {
                application.Visible = false;
                document = application.Documents.Open(sourcePath);
                document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF, OpenAfterExport.Checked);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                document.Close();
                application.Quit();
            }
        }

        //完成後會執行的事件
        private void WordToPDFCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.PerformStep();
            completedJobNum++;
            processLabel.Text = completedJobNum + " / " + totalJob;
        }

        private void SelectMultFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "請選擇需要轉換為 pdf 的 docx 文件",
                Filter = "docx文件(*.docx)|*.docx"
            };

            if (fileDialog.ShowDialog() != DialogResult.OK) { fileDialog.Dispose(); return; }

            string[] names = fileDialog.FileNames;

            completedJobNum = 0;
            totalJob = names.Length;

            progressBar.Value = 0;
            progressBar.Step = progressBar.Maximum / names.Length;

            backgroundWorkers = new BackgroundWorker[totalJob];

            for (int i = 0; i < names.Length; i++)
            {
                string file = names[i];
                string[] path = { file, file.Substring(0, file.Length - 4) + ".pdf" };

                backgroundWorkers[i] = new BackgroundWorker();
                backgroundWorkers[i].DoWork += new DoWorkEventHandler(WordToPDF);
                backgroundWorkers[i].RunWorkerCompleted += new RunWorkerCompletedEventHandler(WordToPDFCompleted);
                backgroundWorkers[i].RunWorkerAsync(path);
            }

            fileDialog.Dispose();
        }

    }
}
```
