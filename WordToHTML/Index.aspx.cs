using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;

namespace WordToHTML
{
    public partial class Index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        //word转换成HTML
        protected void btnChange_Click(object sender, EventArgs e)
        {
            string wordPath = Server.MapPath("/word/test.doc");
            string htmlPath = Server.MapPath("/html/test.html");
            
            //上传word文件, 由于只是做示例，在这里不多做文件类型、大小、格式以及是否存在的判断
            FileUpload1.SaveAs(wordPath);

            #region 文件格式转换
            //请引用Microsoft.Office.Interop.Word
            ApplicationClass word = new ApplicationClass();
            Type wordType = word.GetType();
            Documents docs = word.Documents;

            // 打开文件
            Type docsType = docs.GetType();
            object fileName = wordPath; //"f:\\cc.doc";
            Document doc = (Document)docsType.InvokeMember("Open", BindingFlags.InvokeMethod, null, (object)docs, new Object[] { fileName, true, true });

            //判断与文件转换相关的文件是否存在，存在则删除。（这里，最好还判断一下存放文件的目录是否存在，不存在则创建）
            if (File.Exists(htmlPath)) { File.Delete(htmlPath); }
            //每一个html文件，有一个对应的存放html相关元素的文件夹(html文件名.files)
            if (Directory.Exists(htmlPath.Replace(".html", ".files")))   
            { 
                Directory.Delete(htmlPath.Replace(".html", ".files"), true);
            };

            //转换格式，调用word的“另存为”方法
            Type docType = doc.GetType();
            object saveFileName = htmlPath; //"f:\\aaa.html";
            docType.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, doc, new object[] { saveFileName, WdSaveFormat.wdFormatHTML });

            // 退出 Word
            wordType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, word, null);

            #endregion

            ifrm.Attributes.Add("src", "/html/test.html");
        }
    }
}