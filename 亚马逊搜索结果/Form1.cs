using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Threading;

namespace 亚马逊搜索结果
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int searchnum = 0;
        int sign = 0;
        delegate void DoWork(string data);
        DoWork doWork;
        private void button1_Click(object sender, EventArgs e)
        { 
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            if (!openFileDialog.CheckFileExists)
            {
                MessageBox.Show("请选择一个文件操作");
                return;
            }
            if (!openFileDialog.CheckPathExists)
            {
                MessageBox.Show("请选择一个文件操作");
                return;
            }
            if (openFileDialog.FileName == "")
            {
                MessageBox.Show("请选择一个文件操作");
                return;
            }
            string path = openFileDialog.FileName;  //关键字的文件夹

            //重置
            progressBar1.Value = 0;
            sign = 0;

            string[] data = File.ReadAllLines(path);
            searchnum = data.Length;
            progressBar1.Maximum = searchnum;
            foreach (var gjz in data)
            {
                doWork = new DoWork(GetData);
                doWork.Invoke(gjz);
            }
        }
         
        public void GetData(string gjz)
        {
            //保存获取数据的文件名及地址
            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "//" + DateTime.Now.ToString("yyyyMMdd") + "市场占有率.xls";
            var url = new Uri("https://www.amazon.com/s/ref=nb_sb_noss_2?url=search-alias%3Daps&field-keywords=" + gjz);
            Task task = new Task(()=> {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                //请求的ContentType必须这样设置
                httpWebRequest.ContentType = "text/html;charset=UTF-8";
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                string result = "";
                using (StreamReader sr = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8))
                {
                    result = sr.ReadToEnd();
                }
                string t = "<span id=\"s-result-count\">(.|\\s)*?</span>";
                Regex regex = new Regex(t);
                result = regex.Match(result).Value;
                if (result != "")
                {
                    string[] data = DataChuLi(result);
                    data[0] = gjz;
                    if (!File.Exists(filepath))
                    {
                        CreateExcel(data, filepath);
                    }
                    else
                    {
                        WriteExcel(data, filepath);
                    }
                }
                httpWebResponse.Dispose();
                sign++;
                this.Invoke(new Action(() => {
                    progressBar1.Value = sign;
                    
                    if (sign == searchnum)
                    {
                        Thread.Sleep(500);
                        MessageBox.Show("获取完成");
                    }
                }));
                
            });
            task.Start();
        }
       private void CreateExcel(string[] data, string filepath)
        {
            IWorkbook workbook = new HSSFWorkbook();
            FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate);
            ISheet sheet = workbook.CreateSheet("sheet0");
            IRow row;
            ICell cell;
            row = sheet.CreateRow(0);
            cell = row.CreateCell(0);
            cell.SetCellValue("phrase");
            cell = row.CreateCell(1);
            cell.SetCellValue("volume");
            cell = row.CreateCell(2);
            cell.SetCellValue("amz resutls");
            row = sheet.CreateRow(1);
            cell = row.CreateCell(0);
            cell.SetCellValue(data[0]);
            cell = row.CreateCell(1);
            cell.SetCellValue(data[1]);
            workbook.Write(fs);
            fs.Close();
        }
        private void WriteExcel(string[] data, string filepath)
        {
            FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite);
            FileStream fw = new FileStream(filepath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            POIFSFileSystem ps = new POIFSFileSystem(fs);
            IWorkbook workbook = new HSSFWorkbook(ps);
            ISheet sheet = workbook.GetSheet("sheet0");
            IRow row = sheet.CreateRow(sheet.LastRowNum + 1);
            ICell cell;
            cell = row.CreateCell(0);
            cell.SetCellValue(data[0]);
            cell = row.CreateCell(1);
            cell.SetCellValue(data[1]);
            workbook.Write(fw);
            fw.Close();
        }
        /// <summary>
        /// 将获取的值处理成字符数组，第一个是关键字，第二个是数量
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string[] DataChuLi(string data)
        {
            string[] result = new string[2];
            string[] str;
            if (data.IndexOf("of") >= 0)
            {
                str = Regex.Split(data, "results for");
            }
            else
            {
                string[] zf = Regex.Split(data, "result");
                if (Convert.ToInt32(zf[0]) > 1)
                {
                    str = Regex.Split(data, "results for");
                }
                else
                {
                    str = Regex.Split(data, "result for");
                }
            }
            if (str[1].IndexOf(":") >= 0)
            {
                int a = str[1].LastIndexOf(":");
                string b = str[1].Substring(a);
                result[0] = b.Split('"')[1];
            }
            else
            {
                result[0] = str[1].Replace("\"", "");
            }
            
            if (str[0].IndexOf("of") >= 0)
            {
                result[1] = Regex.Split(str[0], "of")[1].Replace(",", "");
            }
            else
            {
                result[1] = str[0].Replace(",", "");
            }
            return result;
        }
    }
}
