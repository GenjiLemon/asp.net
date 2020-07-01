
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
namespace ADODemo
{
    class word
    {

        public string test()
        {
            string strServerPath = @"D:\1.docx";  //模板路径
            string strSavePath = @"D:\2.docx";  //另存为的路径


            List<string> findText = new List<string>();
            List<string> replaceText = new List<string>();

            findText.Add("{title1}");
            findText.Add("{title2}");
            findText.Add("{time}");
            findText.Add("{city}");
            findText.Add("{p1}");
            findText.Add("{p2}");
            replaceText.Add("民用核安全设备无损检验人员资格鉴定考试");
            replaceText.Add("专业泄漏检验技术Ⅱ级试题及答案");
            replaceText.Add("2013-5");
            replaceText.Add("上海");
            replaceText.Add("张三");
            replaceText.Add("李四");
            ReplaceWordDocAndSave(copyWordDoc(strServerPath), strSavePath, findText, replaceText);
            return strSavePath;
        }
        /// <summary>
        /// 从源DOC文档复制内容返回一个Document类
        /// </summary>
        /// <param name="sorceDocPath">源DOC文档路径</param>
        /// <returns>Document</returns>
        protected Document copyWordDoc(object sorceDocPath)
        {
            object objDocType = WdDocumentType.wdTypeDocument;
            object type = WdBreakType.wdSectionBreakContinuous;

            //Word应用程序变量   
            Word._Application wordApp;
            //Word文档变量
            Document newWordDoc;

            object readOnly = false;
            object isVisible = false;

            //初始化
            //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
            wordApp = new Application();

            Object Nothing = System.Reflection.Missing.Value;

            newWordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            Word._Document openWord;
            openWord = wordApp.Documents.Open(ref sorceDocPath, ref Nothing, ref readOnly, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref isVisible, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            openWord.Select();
            openWord.Sections[1].Range.Copy();

            object start = 0;
            Range newRang = newWordDoc.Range(ref start, ref start);

            //插入换行符   
            //newWordDoc.Sections[1].Range.InsertBreak(ref type);
            newWordDoc.Sections[1].Range.PasteAndFormat(WdRecoveryType.wdPasteDefault);
            openWord.Close(ref Nothing, ref Nothing, ref Nothing);
            return newWordDoc;
        }
        /// <summary>
        /// 替换指定Document的内容，并保存到指定的路径
        /// </summary>
        /// <param name="docObject">Document</param>
        /// <param name="savePath">保存到指定的路径</param>
        protected void ReplaceWordDocAndSave(Document docObject, object savePath, List<string> findText, List<string> replaceText)
        {
            object format = WdSaveFormat.wdFormatDocument;
            object readOnly = false;
            object isVisible = false;

            //string strOldText = "{WORD}";
            //string strNewText = "替换后的文本";

            List<string> IListOldStr = findText;
            List<string> IListNewStr = replaceText;

            string[] newStr = IListNewStr.ToArray();
            int i = 0;

            Object Nothing = System.Reflection.Missing.Value;

            Word._Application wordApp = new Application();
            Word._Document oDoc = docObject;

            object FindText, ReplaceWith, Replace;
            object MissingValue = Type.Missing;

            foreach (string str in IListOldStr)
            {
                oDoc.Content.Find.Text = str;
                //要查找的文本
                FindText = str;
                //替换文本
                //ReplaceWith = strNewText;
                ReplaceWith = newStr[i];
                i++;

                //wdReplaceAll - 替换找到的所有项。
                //wdReplaceNone - 不替换找到的任何项。
                //wdReplaceOne - 替换找到的第一项。
                Replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;

                //移除Find的搜索文本和段落格式设置
                oDoc.Content.Find.ClearFormatting();
                //执行替换
                oDoc.Content.Find.Execute(ref FindText, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref ReplaceWith, ref Replace, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue);
               
            }

            oDoc.SaveAs(ref savePath, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //关闭wordDoc文档对象    
            oDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象    
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }
    }
}
