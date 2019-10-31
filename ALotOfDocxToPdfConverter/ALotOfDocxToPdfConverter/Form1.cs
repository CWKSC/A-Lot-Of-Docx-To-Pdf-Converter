using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

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
            using (OpenFileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "請選擇需要轉換為 pdf 的 docx 文件",
                Filter = "docx文件(*.docx)|*.docx"
            })
            {
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
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

                        //Thread thread = new Thread(SafeCallWordToPDF);
                        //thread.Start(path);

                        backgroundWorkers[i] = new BackgroundWorker();
                        backgroundWorkers[i].DoWork += new DoWorkEventHandler(WordToPDF);
                        backgroundWorkers[i].RunWorkerCompleted += new RunWorkerCompletedEventHandler(WordToPDFCompleted);
                        backgroundWorkers[i].RunWorkerAsync(path);

                    }
                }
            }
        }

    }
}
