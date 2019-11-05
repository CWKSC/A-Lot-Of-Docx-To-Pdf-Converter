# A-Lot-Of-Docx-To-Pdf-Converter

a lot of docx to pdf converter

Since I don't know how to export the C# Project and change the path. The C# project has some problems.

由于我不知道如何导出 C# 项目并更改路径。C# 项目存在一些问题。

You can look directly at the source code to find out how it works.

你可以直接查看源代码去了解如何工作。

Have a code guide to help you understand below.

下面会有代码导读去帮助你的理解。

[C# 笔记 - 批量 docx 到 pdf 转换器](https://zhuanlan.zhihu.com/p/89958561 ) 

![Start]( https://raw.githubusercontent.com/CWKSC/A-Lot-Of-Docx-To-Pdf-Converter/master/Image/Start.png )

![Finish]( https://raw.githubusercontent.com/CWKSC/A-Lot-Of-Docx-To-Pdf-Converter/master/Image/Finish.png)

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

### Member 成员：

```csharp
BackgroundWorker[] backgroundWorkers;
int totalJob;
int completedJobNum;
```

### Method 方法：

```csharp
private void WordToPDF(object sender, DoWorkEventArgs e)
private void WordToPDFCompleted(object sender, RunWorkerCompletedEventArgs e)
private void SelectMultFileButton_Click(object sender, EventArgs e)
```

## Code reading / Guide 代码导读：

1. Start by pressing the button 首先从按下按钮開始

```csharp
SelectMultFileButton_Click()
```

2. Open the file selection window that can be multi-selected 打开可以多重选择的档案选择视窗

```csharp
OpenFileDialog fileDialog = new OpenFileDialog
{
    Multiselect = true,
    Title = "請選擇需要轉換為 pdf 的 docx 文件",
    Filter = "docx文件(*.docx)|*.docx"
};
```

`Multiselect = true` represents multiple choices. 代表可以多重选择。

3. If not press OK, free resources and leave. 如果不是按确定，释放资源并离开。

```csharp
if (fileDialog.ShowDialog() != DialogResult.OK) { fileDialog.Dispose(); return; }
```

4. Get the selected file path and put it in the string array `string[]` 

   获取选择了的档案路径，放到字串阵列 `string[]`

```csharp
string[] names = fileDialog.FileNames;
```

5. Set variables about the progress bar 设定有关进度条的变量

```csharp
completedJobNum = 0;
totalJob = names.Length;

progressBar.Value = 0;
progressBar.Step = progressBar.Maximum / names.Length;
```

6. Create a BackgroundWorker array 创建 BackgroundWorker 陣列

```csharp
backgroundWorkers = new BackgroundWorker[totalJob];
```

[BackgroundWorker Class (System.ComponentModel) | Microsoft Docs](https://link.zhihu.com/?target=https%3A//docs.microsoft.com/en-us/dotnet/api/system.componentmodel.backgroundworker%3Fview%3Dnetframework-4.8)

7. Traverse, generate the path needed by the transformation API, register BackgroundWorker, and run. 

   遍历，生成转换 API 需要的路徑，註冊 BackgroundWorker，運行。

```csharp
for (int i = 0; i < names.Length; i++)
{
    string file = names[i];
    string[] path = { file, file.Substring(0, file.Length - 4) + ".pdf" };

    backgroundWorkers[i] = new BackgroundWorker();
    backgroundWorkers[i].DoWork += new DoWorkEventHandler(WordToPDF);
    backgroundWorkers[i].RunWorkerCompleted += new RunWorkerCompletedEventHandler(WordToPDFCompleted);
    backgroundWorkers[i].RunWorkerAsync(path);
}
```

The main work of BackgroundWorker is in this sentence: / BackgroundWorker 主要的工作在这句：

```csharp
backgroundWorkers[i].DoWork += new DoWorkEventHandler(WordToPDF);
```

The events that will be executed after the BackgroundWorker completes are in this sentence: / BackgroundWorker 完成后会执行的事件在这句：

```csharp
backgroundWorkers[i].RunWorkerCompleted += new RunWorkerCompletedEventHandler(WordToPDFCompleted);
```

BackgroundWorker runs in this sentence: / BackgroundWorker 運行在这句：

```csharp
backgroundWorkers[i].RunWorkerAsync(path);
```

The latter variable is the parameter that it brings to DoWork, and the type will be object. 

后面的变量是它带入去 DoWork 的参数，类型会是 object。

8. Go to the main work part of DoWork - WordToPDF() 

   到达主要工作 DoWork 的部分 —— WordToPDF()

```csharp
private void WordToPDF(object sender, DoWorkEventArgs e)
```

The parameters that RunWorkerAsync takes in are placed in e.Argument instead of directly in the parameter list. 

RunWorkerAsync 带入去的参数会放在 e.Argument ，而不是直接在参数列表。

9. Read just put in the variable `path` 读取刚刚放进去变量 `path`

```csharp
string sourcePath = ((string[]) e.Argument)[0];
string targetPath = ((string[]) e.Argument)[1];
```

10. Create variables for Word.Application and Document. 创建 Word.Application 和 Document 的变量。

```csharp
Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
Document document = null;
```

11. Set to invisible, Word.Application opens, document accepts, and docx to pdf conversion. 

    设定为不可见，Word.Application 打开，document 接受，进行 docx 到 pdf 的转换。

```csharp
try
{
    application.Visible = false;
    document = application.Documents.Open(sourcePath);
    document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF, OpenAfterExport.Checked);
}
```

[Document.ExportAsFixedFormat method (Word) | Microsoft Docs](https://link.zhihu.com/?target=https%3A//docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat)

```csharp
document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF, OpenAfterExport.Checked);
```

There are two required parameters, OutputFileName and ExportFormat. 

有两个必须的参数，OutputFileName 和 ExportFormat 。

I have added an Optional parameter here: OpenAfterExport, which is determined by the tick option on WinForm.

我这里加了一个 Optional 的参数：OpenAfterExport，由 WinForm 上的勾选项决定。

12. Catch Exception 抓取错误

```csharp
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
```

13. Turn off document and application to free resources 关掉 document 和 application 以释放资源

```csharp
finally
{
    document.Close();
    application.Quit();
}
```

14. Execute events that will be executed after BackgroundWorker completes 

    执行 BackgroundWorker 完成后会执行的事件

```csharp
private void WordToPDFCompleted(object sender, RunWorkerCompletedEventArgs e)
{
    progressBar.PerformStep();
    completedJobNum++;
    processLabel.Text = completedJobNum + " / " + totalJob;
}
```

This is related to the progress bar, not much to say. 这个跟进度条有关，不多说。
