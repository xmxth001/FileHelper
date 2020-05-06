using Aspose.Words;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace FileHelper
{
    /// <summary>
    /// word 文档操作类
    /// </summary>
    public class WordHelper : IDisposable
    {

        private Microsoft.Office.Interop.Word._Application _app;
        private Microsoft.Office.Interop.Word._Document _doc;
        private Word.Application app;
        object _nullobj = System.Reflection.Missing.Value;

        /// <summary>
        /// 关闭Word进程
        /// </summary>
        public void KillWinword()
        {
            var p = Process.GetProcessesByName("WINWORD");
            if (p.Any())
                p[0].Kill();
        }

        /// <summary>
        /// 打开word文档
        /// </summary>
        /// <param name="filePath"></param>
        public void Open(string filePath)
        {
            _app = new Microsoft.Office.Interop.Word.ApplicationClass();
            object file = filePath;
            _doc = _app.Documents.Open(
              ref file, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj, ref _nullobj);
        }

        /// <summary>
        /// 替换word中的文字
        /// </summary>
        /// <param name="strOld">查找的文字</param>
        /// <param name="strNew">替换的文字</param>
        public void Replace(string strOld, string strNew)
        {
            _app.Selection.Find.ClearFormatting();
            _app.Selection.Find.Replacement.ClearFormatting();
            _app.Selection.Find.Text = strOld;
            _app.Selection.Find.Replacement.Text = strNew;
            object objReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            _app.Selection.Find.Execute(ref _nullobj, ref _nullobj, ref _nullobj,
                   ref _nullobj, ref _nullobj, ref _nullobj,
                   ref _nullobj, ref _nullobj, ref _nullobj,
                   ref _nullobj, ref objReplace, ref _nullobj,
                   ref _nullobj, ref _nullobj, ref _nullobj);
        }
        /// <summary>
        /// 通配符替换
        /// </summary>
        /// <param name="strOld"></param>
        /// <param name="strNew"></param>
        public void ReplaceWildcard(string strOld, string strNew)
        {
            _app.Selection.Find.ClearFormatting();
            _app.Selection.Find.Replacement.ClearFormatting();
            object objReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            object objWildcard = true;
            object FindText = strOld;
            object NewText = "";
            bool flag = _app.Selection.Find.Execute(ref FindText, ref _nullobj, ref _nullobj,
                   ref objWildcard, ref _nullobj, ref _nullobj,
                   ref _nullobj, ref _nullobj, ref _nullobj,
                   ref NewText, ref objReplace, ref _nullobj,
                   ref _nullobj, ref _nullobj, ref _nullobj);
        }
        /// <summary>
        /// 另存为
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string SaveAs(string filePath)
        {
            object file = filePath;
            object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;
            _doc.SaveAs(
              ref file, format, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj,
              ref _nullobj, ref _nullobj, ref _nullobj, ref _nullobj);
            return filePath;
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            _doc.Save();
        }
        /// <summary>
        /// 保存为PDF格式
        /// </summary>
        /// <param name="FileName"></param>
        public string SavePDF(string filePath)
        {
            bool paramOpenAfterExport = false;
            WdExportOptimizeFor paramExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
            WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;

            int paramStartPage = 0;
            int paramEndPage = 0;

            WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;

            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            _doc.ExportAsFixedFormat("abc.pdf",
                    WdExportFormat.wdExportFormatPDF,
                    paramOpenAfterExport,
                    paramExportOptimizeFor, paramExportRange, paramStartPage,
                    paramEndPage, paramExportItem, paramIncludeDocProps,
                    paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                    paramBitmapMissingFonts, paramUseISO19005_1,
                    ref _nullobj);

            return filePath;
        }

        /// <summary>
        /// Aspose组件
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        public void WordToPdf(string sourcePath, string targetPath)
        {

            Aspose.Words.Document doc = new Aspose.Words.Document(sourcePath);
            //Aspose.Words.DocumentBuilder builder = new DocumentBuilder(doc);
            //builder.Font.Name = "宋体";
            doc.Save(targetPath, SaveFormat.Pdf);

        }
        /// <summary>
        /// WPS组件
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        public void WPSWordToPdf(string sourcePath, string targetPath)
        {
            app = new Word.ApplicationClass();
            Word.Document doc = app.Documents.Open(sourcePath, Visible: false);
            doc.ExportAsFixedFormat(targetPath, Word.WdExportFormat.wdExportFormatPDF);
            doc.Close();
            app.Quit();
        }

        /// <summary>
        /// 读取文件流
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public FileStream ReadFileStream(string fileName, string ext = ".doc")
        {
            fileName += ext;
            if (File.Exists(@"" + fileName))
            {
                FileStream fileStream = new FileStream(fileName, FileMode.Open);
                return fileStream;
            }
            return null;
        }
        /// <summary>
        /// 退出
        /// </summary>
        public void Dispose()
        {
            _doc.Close(ref _nullobj, ref _nullobj, ref _nullobj);
            _app.Quit(ref _nullobj, ref _nullobj, ref _nullobj);
        }

        /// <summary>
        /// 打印预览
        /// </summary>
        /// <param name="filePath"></param>
        public void PrintPreview()
        {
            _doc.PrintPreview();
        }
        /// <summary>
        /// 打印
        /// </summary>
        public void PrintOut()
        {
            //_doc.PrintOut();
        }
    }
}
