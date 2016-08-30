using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Web.Script.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Microsoft.Office.Interop.Word;
using System.Net;
using System.Diagnostics;
using System.Management.Instrumentation;
using System.Management;


namespace WebApplication2
{
    /// <summary>
    /// WebService2 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class WebService2 : System.Web.Services.WebService
    {
        public string in_column;
        public static ManagementObjectCollection mn = (new ManagementClass("Win32_Processor")).GetInstances();
        public static int doc_handler = 8;//总句柄
        public static int slice = 16;//线程数目
        int key_line = 25;//最大rs记录的block块长
        _Document[] docs_list = new _Document[25];
        _Application[] apps_list = new Microsoft.Office.Interop.Word.Application[25];
        public int pcount;
        public int tc_count;
        public string[] columns;
        public string[] regrex=null;
        public Boolean IsMerge=false;
       // public string root = "D:\\files\\words\\";//文件保存的路径;
        public string root = "E:\\files\\words\\";
        public string OfficeFilePath = "E:\\files\\office\\"; // "D:\\pdf\\office\\"; //
        public string PdfFilePath = "E:\\files\\pdfs\\";//  "D:\\pdf\\pdf\\";//
        public string SWFFilePath = "E:\\files\\swfs\\";    //  "D:\\pdf\\swf\\"; //
        public class finaljson
        {

            //估计有上锁机制,导致会吧inman
            public List<Dictionary<string, string>> finalstrings = new List<Dictionary<string, string>>();
            //  public  List<Dictionary<string, string>> [] finals = new List<Dictionary<string, string>> [25];
            public List<Dictionary<string, object>> final_tc = new List<Dictionary<string, object>>();
        }
        public finaljson aaa = new finaljson();

        public class JsonTools
        {
            // 从一个对象信息生成Json串
            public static string ObjectToJson(object obj)
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                MemoryStream stream = new MemoryStream();
                serializer.WriteObject(stream, obj);
                byte[] dataBytes = new byte[stream.Length];
                stream.Position = 0;
                stream.Read(dataBytes, 0, (int)stream.Length);
                return Encoding.UTF8.GetString(dataBytes);
            }
            // 从一个Json串生成对象信息
            public static object JsonToObject(string jsonString, object obj)
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                MemoryStream mStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
                return serializer.ReadObject(mStream);
            }
        }
        public string swfnanme;
        private string OfficeToPdf(string OfficePath, string OfficeName, string destPath)
        {
            string fullPathName = OfficePath + OfficeName;//包含 路径 的全称
            FileInfo fi1 = new FileInfo(fullPathName);
            fi1.Attributes = ~FileAttributes.ReadOnly;
            string fileNameWithoutEx = System.IO.Path.GetFileNameWithoutExtension(OfficeName);//不包含路径，不包含扩展名
            string extendName = System.IO.Path.GetExtension(OfficeName).ToLower();//文件扩展名
            string saveName = destPath + fileNameWithoutEx + ".pdf";
            string returnValue = fileNameWithoutEx + ".pdf";
            Util.WordToPDF(fullPathName, saveName);
            return returnValue;
        }
        private string PdfToSwf(string pdf2swfPath, string PdfPath, string PdfName, string destPath)
        {
            string fullPathName = PdfPath + PdfName;//包含 路径 的全称
            string fileNameWithoutEx = System.IO.Path.GetFileNameWithoutExtension(PdfName);//不包含路径，不包含扩展名
            string extendName = System.IO.Path.GetExtension(PdfName).ToLower();//文件扩展名
            string saveName = destPath + fileNameWithoutEx + ".swf";
            string returnValue = fileNameWithoutEx + ".swf";
            Util.PDFToSWF(pdf2swfPath, fullPathName, saveName);
            return returnValue;
        }
        public string showwordfiles(string filename)
        {
            string pdf2swfToolPath = System.Web.HttpContext.Current.Server.MapPath("~/FlexPaper/pdf2swf.exe");
            string SwfFileName = String.Empty;
            string UploadFileName = System.IO.Path.GetFileNameWithoutExtension(filename) + ".doc";
            string UploadFileType = System.IO.Path.GetExtension(UploadFileName).ToLower();
            string UploadFileNameFileFullName = String.Empty;
            UploadFileNameFileFullName = OfficeFilePath + UploadFileName;
            File.Copy(filename, UploadFileNameFileFullName);
            string PdfFileName = OfficeToPdf(OfficeFilePath, UploadFileName, PdfFilePath);
            SwfFileName = PdfToSwf(pdf2swfToolPath, PdfFilePath, PdfFileName, SWFFilePath);
            return SWFFilePath;
        }
        public string delet_tables(string filename)
        {
            _Application appdelet_tables = new Microsoft.Office.Interop.Word.Application();
            _Document docdelet_tables;
            object fileName = filename;
            object unknow = System.Type.Missing;
            docdelet_tables = appdelet_tables.Documents.Open(ref fileName,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
            int table_num = docdelet_tables.Tables.Count;
            try
            {
                for (int i = 1; i <= table_num; i++)
                {
                    docdelet_tables.Tables[i].Delete();
                }
            }
            catch { }
            docdelet_tables.Close(ref unknow, ref unknow, ref unknow);
            appdelet_tables.Quit(ref unknow, ref unknow, ref unknow);
            return null;
        }

        public void thread1(Document doc, int start, int end)//,ref List<Dictionary<string, string>> finalstrings)
        {
            string temp = null;
            // int start = 1;
            //int end = pcount
            if (start <= 0) start = 1;
            if (end >= pcount) end = pcount;
            Dictionary<string, string> map = null;
            //  finalstrings=new List<Dictionary<string,string>>();
            Debug.WriteLine("inside {0}=>{1},{2}", doc.GetHashCode(), start, end);
            while (start <= end)
            {
                temp = doc.Paragraphs[start].Range.Text.Trim();


                if (map == null && Regex.Matches(temp, @"(^\[(?!End).*\]$)", RegexOptions.IgnoreCase).Count > 0)
                {
                    map = new Dictionary<string, string>(); map.Add("tag", temp);
                }
                else if (map != null)
                {
                    if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                    {
                        //大小写的问题尚未解决呢
                        Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                        // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                        if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                        {
                            string key=mark.Groups[1].ToString().Trim();string value=mark.Groups[2].Value.Trim();
                            if (!map.ContainsKey(key)) map.Add(key,value);
                            else { map[key]+= (value + ","); }//
                        }

                    }
                    else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                    {
                        if (!aaa.finalstrings.Contains(map))
                        {
                            aaa.finalstrings.Add(map);
                            map = null;
                        }//销毁对象}
                    }
                    else
                    {
                        //description字段哦,如果用户没有输入
                        if (in_column.Contains("description"))
                        {
                            if (map.ContainsKey("description")) { map["description"] += temp; }//1/(11):(map1.Add("description",temp));
                            else { map.Add("description", temp); }
                        }
                    }

                }//else map1==null
                else { }
                start++;
            }//while



        }//thread1


        [WebMethod]
        public string downfile(string doc_url)
        {

            string LocalPath = root + doc_url.Substring(doc_url.LastIndexOf('/'));
            FileStream fs = null;
            Stream sIn = null;
            HttpWebResponse wr = null;
            Uri u = new Uri(doc_url);

            HttpWebRequest mRequest = (HttpWebRequest)WebRequest.Create(u);
            mRequest.Method = "GET";
            mRequest.ContentType = "application/x-www-form-urlencoded";
            wr = (HttpWebResponse)mRequest.GetResponse();
            sIn = wr.GetResponseStream();
            fs = new FileStream(LocalPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            byte[] bytes = new byte[4096];
            int start = 0;
            int length;
            while ((length = sIn.Read(bytes, 0, 4096)) > 0)
            {
                fs.Write(bytes, 0, length);
                start += length;
            }



            if (sIn != null) sIn.Close();
            if (wr != null) wr.Close();
            if (fs != null) fs.Close();

            return LocalPath;

        }

        [WebMethod(Description = "readtc_void")]
        public void readtc(Document doc, int start_line, int end_line)
        {

            start_line = start_line <= 0 ? 1 : start_line;
            end_line = end_line >= pcount ? pcount : end_line;
            Debug.WriteLine("inside {0}=>{1},{2}", doc, start_line, end_line);
            Microsoft.Office.Interop.Word.Table nowTable;
            Dictionary<string, object> h = null;
            for (int tablePos = start_line; tablePos <= end_line; tablePos++)
            {
                nowTable = doc.Tables[tablePos];
                Regex tsp_mark = new Regex(@"^(\[.*?[^\r\n]*).*?");
                if (!tsp_mark.Match(nowTable.Cell(1, 2).Range.Text.Trim()).Success)
                {
                    // aaa.finalstrings.Add(nowTable.Cell(1, 2).Range.Text.Trim());
                    continue;
                }
                h = new Dictionary<string, object>();
                h.Add("tag", tsp_mark.Match(nowTable.Cell(1, 2).Range.Text.Trim()).Groups[1].Value);
                Dictionary<string, string> tc_2step = null; h["test steps"] = null;
                Range range = nowTable.Range;
                Debug.WriteLine("fuck you 4,2" + nowTable.Columns.Count);
                if (IsMerge)
                {       
                      
                        Cell   tmp=range.Cells[1];
                        while (tmp != null)
                        {   
                            Debug.WriteLine("content is "+tmp.Range.Text);
                            string flag = Regex.Match(tmp.Range.Text.Trim().ToLower().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase).Groups[1].Value;//,@"\w",RegexOptions.IgnoreCase).Groups[0].Value.Trim();
                            
                            string[] flags = flag.Split(new char[] { ' ' });
                            flag = string.Join(" ", flags).Trim();
                            Debug.WriteLine("flag is the"+flag);
                            bool mark = false;
                            foreach (string a in columns)
                            {
                                if (a.ToLower().Trim() == flag) { mark = true; break; }

                            }
                            if(mark&&((flag=="test case")||(flag=="test steps"))){

                                h["test steps"] = "";
                                Cell a = tmp.Next; int start_row = tmp.RowIndex; string[] column_name = new string[nowTable.Columns.Count];int k=1;
                                int running_row = start_row; Dictionary<string, string> tc_step = null;
                                while (a != null && a.ColumnIndex != 1)
                                {
                                    //Debug.WriteLine("当前的Text"+a.RowIndex+a.Range.Text);
                                    if (a.RowIndex == start_row)//作为key列
                                    {
                                        Match temp = Regex.Match(a.Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase);
                                        column_name[k++] = temp.Groups[1].Value.ToLower();
                                        
                                    }
                                    else
                                    {
                                        //分为两列和三列的情形吧,所以必须是换行的时候添加进去
                                        if (running_row != a.RowIndex&&tc_step!=null)
                                        {
                                            //此时也要把上一次tc_step加入进来
                                            h["test steps"] += (new JavaScriptSerializer().Serialize(tc_step)) + ",";
                                            tc_step = new Dictionary<string, string>();
                                            tc_step.Add(column_name[a.ColumnIndex-1], a.Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                            running_row = a.RowIndex;
                                        }
                                       
                                        else if (running_row != a.RowIndex && tc_step == null)
                                        {
                                            tc_step = new Dictionary<string, string>();
                                            tc_step.Add(column_name[a.ColumnIndex-1], a.Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                            running_row = a.RowIndex;
                                        }
                                        else if (tc_step != null )
                                        {
                                            Debug.WriteLine("当前的col{0},{1}", column_name[a.ColumnIndex - 1], a.Range.Text);
                                            tc_step.Add(column_name[a.ColumnIndex - 1], a.Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                            
                                        }

                                    }//else

                                    a = a.Next;

                                }//while 循环结束的情况哦
                                  if(tc_step!=null)h["test steps"] += (new JavaScriptSerializer().Serialize(tc_step)) + ",";      
                                 if(h["test steps"].ToString().Length>0)h["test steps"] = "[" + h["test steps"].ToString().Substring(0, h["test steps"].ToString().Length - 1) + "]";
                               


                            }
                            else if (mark)
                            {
                                 Debug.WriteLine("两列格式读取ing,flag is the"+flag);
                                if (tmp.RowIndex != 1)
                                {
                                    h.Add(flag, tmp.Next.Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                    tmp = tmp.Next.Next;
                                    continue;
                                }
                                //否则要匹配regx喽
                                //Debug.WriteLine("the regrex  length is " + (regrex.Length));
                                if (regrex != null)
                                {
                                    foreach (string aa in regrex)
                                    {
                                        string a = aa.ToLower().Trim(); string key = null, value = null;
                                        MatchCollection matches = Regex.Matches(@tmp.Range.Text, @a, RegexOptions.IgnoreCase);
                                        Debug.WriteLine("reg : {0} ;  text :{1}", a, tmp.Range.Text);// ,  all : {1}", a);, new JavaScriptSerializer().Serialize(regrex));
                                        foreach (Match match in matches)
                                        {
                                            GroupCollection groups = match.Groups;//h.Add(a, null);
                                            key = groups[1].ToString().Trim(); value = groups[2].ToString().Trim();
                                            Debug.WriteLine("key: {0} , value: {1}", key, value);
                                            h[key] = h.ContainsKey(key) ? (h[key] += (value + ",")) : (value += ",");

                                        }
                                        //最后一次的
                                        if (key != null && h.ContainsKey(key) && h[key].ToString().Length != 0 && h[key].ToString().LastIndexOf(',') == h[key].ToString().Length - 1) h[key] = h[key].ToString().Substring(0, h[key].ToString().Length - 1);
                                    }//foreach
                                }//if regrex!=null
                                //description的key有点特殊哦
                                Match match_desc = Regex.Match(@tmp.Next.Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""), @"\]([^\[\]#]*)", RegexOptions.IgnoreCase);
                                if (match_desc.Success) { h.Add(flag, match_desc.Groups[1].Value); }

                            }
                            else if (!mark && (flag == "test case") || (flag == "test steps"))
                            {
                                //column里面没有test steps
                                    while (tmp.Next != null && tmp.Next.ColumnIndex != 1)
                                        tmp = tmp.Next;
                            }
                            
                            //h.steps 要走的,没有mark的也要走
                          tmp = tmp.Next;
                        }//while

















                }
                else
                {
                    for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
                    {  //还要把中文给去掉,以及纵向合并的单元格


                        string flag = Regex.Match(nowTable.Cell(rowPos, 1).Range.Text.Trim().ToLower().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase).Groups[1].Value;//,@"\w",RegexOptions.IgnoreCase).Groups[0].Value.Trim();
                        string[] flags = flag.Split(new char[] { ' ' });
                        flag = string.Join(" ", flags).Trim();
                        //   aaa.finalstrings.Add(flag);
                        bool mark = false;
                        foreach (string a in columns)
                        {
                            if (a.ToLower().Trim() == flag) { mark = true; break; }

                        }
      
                        if (mark && nowTable.Rows[rowPos].Cells.Count == 2)
                        {
                            string text = nowTable.Rows[rowPos].Cells[2].Range.Text.Trim().Replace("\r", "").Replace("\u0007", "");
                            //  aaa.finalstrings.Add(flag + "haha here" + nowTable.Rows[rowPos].Cells[1].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                            // Dictionary<string, string> tc_step = null; h["test steps"] = null;
                            switch (flag)
                            {
                                case "execution step":
                                    tc_2step = new Dictionary<string, string>();
                                    tc_2step.Add("num", "1");
                                    tc_2step.Add("actions", text);
                                    Debug.WriteLine("{2},{0},{1}", tc_2step, h["test steps"], rowPos);
                                    break;
                                case "expected output":
                                    if (tc_2step != null)
                                    {
                                        tc_2step.Add("expected result", text);
                                        h["test steps"] += (new JavaScriptSerializer().Serialize(tc_2step)); tc_2step = null;
                                        h["test steps"] = "[" + h["test steps"].ToString() + "]";
                                    }
                                    break;
                                default:
                                    if (rowPos != 1) { h.Add(flag, text); continue; }
                                    string text1 = nowTable.Rows[rowPos].Cells[2].Range.Text.Trim().Replace("\u0007", "");
                                    if (regrex != null)
                                    {
                                        foreach (string aa in regrex)
                                        {
                                            string a = aa.ToLower().Trim(); string key = null, value = null;
                                            MatchCollection matches = Regex.Matches(@text1, @a, RegexOptions.IgnoreCase);
                                            Debug.WriteLine("reg : {0} ;  text :{1}", a, text1);// ,  all : {1}", a);, new JavaScriptSerializer().Serialize(regrex));
                                            foreach (Match match in matches)
                                            {
                                                GroupCollection groups = match.Groups;//h.Add(a, null);
                                                key = groups[1].ToString().Trim(); value = groups[2].ToString().Trim();
                                                Debug.WriteLine("key: {0} , value: {1}", key, value);
                                                h[key] = h.ContainsKey(key) ? (h[key] += (value + ",")) : (value += ",");

                                            }
                                            //最后一次的
                                            if (key != null && h.ContainsKey(key) && h[key].ToString().Length != 0 && h[key].ToString().LastIndexOf(',') == h[key].ToString().Length - 1) h[key] = h[key].ToString().Substring(0, h[key].ToString().Length - 1);


                                        }//foreach
                                    }//if
                                    Match match_desc = Regex.Match(text, @"\]([^\[\]#]*)", RegexOptions.IgnoreCase);
                                    if (match_desc.Success) { h.Add(flag, match_desc.Groups[1].Value); }

                                    break;

                            }//switch
                            //此时要解析出description,source,safety
                        }
                        else if (mark && nowTable.Rows[rowPos].Cells.Count > 2)//不要后面因为合并单元格
                        {
                            //处理tc_test的情况了的
                            //收集列名字
                            /*   aaa.finalstrings.Add(flag + "tc_step here" + in_column.Trim());
                              for (int start = rowPos + 1; start < nowTable.Rows.Count; start++)
                            {

                                if (nowTable.Rows[start].Cells.Count != num)
                                {
                                    rowPos = start - 1; break;
                                }
                              }
                            nowTable.Rows[rowPos].Cells[i].M
                             * */
                            int num = nowTable.Rows[rowPos].Cells.Count;
                            string[] column_name = new string[num + 1];

                            for (int i = 1; i <= nowTable.Rows[rowPos].Cells.Count; i++)
                            {

                                Match temp = Regex.Match(nowTable.Rows[rowPos].Cells[i].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase);
                                column_name[i] = temp.Groups[1].Value.ToLower();// +nowTable.Rows[rowPos].Cells[i].Range.Text.Trim();

                            }
                            column_name[0] = nowTable.Rows[rowPos].Range.Text.ToString();
                            //column_name[1] = "Actions"; column_name[2] = "Expected Result";
                            Console.WriteLine(nowTable.Rows[rowPos].Cells[num].Range.Text);
                            //  column_name[1]="num";
                            h["test steps"] = null;
                            //往下收集列即可
                            int k = 0;
                            for (int start = rowPos + 1; start <= nowTable.Rows.Count; start++)
                            {

                                if (nowTable.Rows[start].Cells.Count != num)
                                {
                                    rowPos = start - 1; break;
                                }

                                Dictionary<string, string> tc_step = new Dictionary<string, string>();
                                int j;
                                for (j = 1; j <= num; j++)
                                {

                                    tc_step.Add(column_name[j], nowTable.Rows[start].Cells[j].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                    // Debug.WriteLine("hehe"+tc_step[column_name[j]]);


                                }


                                h["test steps"] += (new JavaScriptSerializer().Serialize(tc_step)) + ",";
                                //Debug.WriteLine("呵呵" + h["test steps"]);


                            }//处理tc_step
                            h["test steps"] = "[" + h["test steps"].ToString().Substring(0, h["test steps"].ToString().Length - 1) + "]";
                        }//else if 


                    }//for row
                }//else IsMerge

                aaa.final_tc.Add(h);

            }//for tables
         

        }












        [WebMethod(Description = "readtc")]
        public string resolve()
        {

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            string param = System.Web.HttpUtility.UrlDecode(HttpContext.Current.Request.Url.Query.ToString().Substring(1));
            string[] param_s = param.Split('&'); Hashtable key_values = new System.Collections.Hashtable(); ;
            foreach (string item in param_s)
            {
                string[] key_value = item.Split('=');
                //  this.GetType().GetField(key_value[0]).GetValue(key_value[1]).ToString();
                key_values.Add(key_value[0], key_value[1]);
            }


            String LocalPath = null;
            string column = key_values.ContainsKey("column") ? key_values["column"].ToString() : "", type = key_values.ContainsKey("type") ? key_values["type"].ToString() : "", doc_url = key_values.ContainsKey("doc_url") ? key_values["doc_url"].ToString() : "";
            regrex = (key_values.ContainsKey("regrex") && key_values["regrex"].ToString()!="") ? System.Web.HttpUtility.UrlDecode(key_values["regrex"].ToString()).Split(',') : null;
            IsMerge = (key_values.ContainsKey("ismerge") && key_values["ismerge"].ToString() != ""&& key_values["ismerge"].ToString() != "0") ? true : false;
            if (column.Equals("") || type.Equals("") || doc_url.Equals(""))
                throw new Exception("输入参数不合法");
            //  return column;  
            string message = null;
            try
            {
                String savePath = downfile(doc_url);
                //  string pdfpath = showwordfiles(savePath);
                //不需要此步判断了 if (!File.Exists(savePath)) throw new Exception("保存文件失败" ）; 

                in_column = column;
                //WORD  中数据都规整为一个空格隔开来
                columns = column.Split(',');
                object fileName = savePath;

                object unknow = System.Type.Missing;

                //目前一个线程再跑
                if (type == "rs") { delet_tables(savePath); }

                //count the paragraphs

                if (type == "rs")
                {
                    //然后完成对文档的解析

                    List<System.Threading.Tasks.Task> TaskList = new List<System.Threading.Tasks.Task>();
                    // 开启线程池,线程分配算法
                    System.Threading.Tasks.Task t = null;
                    int k = (int)Math.Ceiling((Double)slice / doc_handler);
                    for (int i = 0; i < doc_handler; i++)
                    {
                        Microsoft.Office.Interop.Word.Application app_in = new Microsoft.Office.Interop.Word.Application();
                        var doc1 = app_in.Documents.Open(ref fileName, ref unknow, true, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
                        pcount = doc1.Paragraphs.Count;
                        docs_list[i] = (doc1);
                        apps_list[i] = (app_in);

                        int block = (int)Math.Ceiling((Double)pcount / slice);
                        for (int j = i * k + 1; j <= (i + 1) * k && j <= slice; j++)
                        {

                            //Debug.WriteLine("传入{0},{1}", pcount * (i) / slice + 1 - key_line, pcount * (i + 1) / slice + key_line);
                            int start_in = block * (j - 1) + 1 - key_line <= 0 ? 1 : block * (j - 1) + 1 - key_line, end_in = block * (j) + key_line >= pcount ? pcount : block * (j) + key_line;
                            Debug.WriteLine("outside{0}=>{1},{2}", doc1.GetHashCode(), start_in, end_in);
                            t = new System.Threading.Tasks.Task(() => thread1(doc1, start_in, end_in));//, ref aaa.finals[i]));
                            t.Start();
                            TaskList.Add(t);
                        }
                    }



                    /*     
                         var t1 = new System.Threading.Tasks.Task(() => thread1(1,pcount/8+key_line, map[0]));
                         t8.Start(); TaskList.Add(t8);
                 
                 */

                    System.Threading.Tasks.Task.WaitAll(TaskList.ToArray());//t1, t2, t3, t4, t5, t6, t7, t8);
                    var json = new JavaScriptSerializer().Serialize(aaa.finalstrings.Where((x, i) => aaa.finalstrings.FindIndex(z => z["tag"] == x["tag"]) == i).ToList());
                    message = json;// "{\"success\":true,\"msg\":" + (new JavaScriptSerializer().Serialize(json)) + "}";

                }//rs
                else if (type == "tc")
                {

                    //tc并发并没有什么问题


                    List<System.Threading.Tasks.Task> TaskList = new List<System.Threading.Tasks.Task>();

                    int k = (int)Math.Ceiling((Double)slice / doc_handler);
                    Debug.WriteLine("buchang {0}", k);
                    for (int i = 0; i < doc_handler; i++)
                    {
                        Microsoft.Office.Interop.Word.Application app_in = new Microsoft.Office.Interop.Word.Application();
                        var doc1 = app_in.Documents.Open(ref fileName, ref unknow, true, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
                        pcount = doc1.Tables.Count;
                        docs_list[i] = (doc1);
                        apps_list[i] = (app_in);
                        int block = (int)Math.Ceiling((Double)pcount / slice);

                        for (int j = i * k + 1; j <= (i + 1) * k && j <= slice; j++)
                        {
                            //Debug.WriteLine("传入{0},{1}", pcount * (i) / slice + 1 - key_line, pcount * (i + 1) / slice + key_line);

                            int start_in = block * (j - 1) + 1 <= 0 ? 1 : block * (j - 1) + 1, end_in = block * (j) >= pcount ? pcount : block * (j);
                            Debug.WriteLine("outside {0},{1}", start_in, end_in);
                            System.Threading.Tasks.Task t = new System.Threading.Tasks.Task(() => readtc(doc1, start_in, end_in));
                            t.Start();
                            TaskList.Add(t);
                        }
                    }




                    System.Threading.Tasks.Task.WaitAll(TaskList.ToArray());
                    var json = new JavaScriptSerializer().Serialize(aaa.final_tc);
                    //(new HashSet<Dictionary<string, object>>(aaa.final_tc)));
                    message = json;// "{\"success\":true,\"msg\":" + (new JavaScriptSerializer().Serialize(json)) + "}";



                }


            }

            catch (Exception e)
            {

                message = e.ToString();//
                // "{\"success\":false,\"msg\":\"" + e.Message + e.StackTrace + e.TargetSite + "\"}";

                return message;


            }

            finally
            {
                stopwatch.Stop();

                //  return  json;

                //异不异常到最后都关闭文档,避免word一直处于打开状态占用资源
                object unknows = System.Type.Missing;
                Debug.WriteLine("打开大小为{0}", docs_list.Length);
                //   if (doc != null) doc.Close();
                for (int i = 0; i < docs_list.Length; i++)
                {
                    if (docs_list[i] == null) { break; } Debug.WriteLine("开始close"); docs_list[i].Close(ref unknows, ref unknows, ref unknows); Debug.WriteLine("结束close");
                }
                for (int i = 0; i < apps_list.Length; i++)
                {

                    if (apps_list[i] == null) { break; } Debug.WriteLine("开始quit"); apps_list[i].Quit(ref unknows, ref unknows, ref unknows); Debug.WriteLine("结束quit");
                }

                GC.Collect();
                GC.Collect();
                Context.Response.ContentType = "text/json";
                // Context.Response.Write(stopwatch.Elapsed);
                Context.Response.Write(message);
                // Context.Response.Write(json);

                Context.Response.End();

                /*     app1.Quit(ref unknow, ref unknow, ref unknow);
               
                 */







            }//finally
            return null;//不是正常请求方式,不可见这结果
        }




        [WebMethod(Description = "readtitles")]
        public string readtitles(string filename)
        {
            _Application app = new Microsoft.Office.Interop.Word.Application();
            _Document doc;

            object fileName = filename;
            object unknow = System.Type.Missing;
            doc = app.Documents.Open(ref fileName,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
            object pcount = doc.Paragraphs.Count;//count the paragraphs
            object trydocfunc = doc.Tables.Count;
            object listnumbers = doc.Lists.Count;
            object listpnumbers = doc.ListParagraphs.Count;
            object listtbumbers = doc.ListTemplates.Count;
            Lists lists = doc.Lists;
            ListParagraphs listps = doc.ListParagraphs;
            ListTemplates listts = doc.ListTemplates;
            object list1 = lists[1];
            object list2 = lists[2];
            object list3 = listps[1];
            object list4 = listts[1];
            string k = listps[3].Range.Text.Trim();
            //object level = lists[1].ApplyListTemplateWithLevel
            string[] k3 = new string[3];
            for (int i = 0; i <= 1; i++)
            {
                k3[i] = lists[i + 1].Range.Text.Trim();
            }
            int[] num3 = new int[2];
            for (int i = 0; i <= 1; i++)
            {
                num3[i] = lists[i + 1].Range.Start;
            }
            string[] k4 = new string[3];
            for (int i = 0; i <= 2; i++)
            {
                k3[i] = listps[i + 1].Range.Text.Trim();
            }
            listpnumbers = doc.ListParagraphs.Count;
            app.Documents.Close(ref unknow, ref unknow, ref unknow);
            app.Quit(ref unknow, ref unknow, ref unknow);
            return null;
        }


    }



}


