using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace FileHelper
{
    /// <summary>
    /// 文件操作类
    /// </summary>
    public class FileHelper
    {
        /// <summary>
        /// 写入txt
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="Content"></param>
        public static void WriteTxt(string FilePath, string Content)
        {
            if (!string.IsNullOrEmpty(Content))
            {
                string DateNow = DateTime.Now.ToString();
                string OldText = ReadTxt(FilePath);
                FileStream fs = new FileStream(FilePath, FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs);
                //开始写入
                sw.WriteLine(OldText + "时间【" + DateNow + "】" + Content);
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string ReadTxt(string path)
        {
            string result = "";
            StreamReader sr = new StreamReader(path, Encoding.UTF8);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                result += line.ToString();
            }
            sr.Close();
            return result;
        }
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool DeleteFile(string path)
        {
            FileAttributes attr = File.GetAttributes(path);
            if (attr == FileAttributes.Directory)
            {
                //Directory.Delete(path , true);
                return false;
            }
            else
            {
                File.Delete(path);
                return true;
            }
        }
        /// <summary>
        /// 远程文件转Stream
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static Stream FileToStream(string url)
        {
            try
            {
                //设置参数 
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                //发送请求并获取相应回应数据 
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                //直到request.GetResponse()程序才开始向目标网页发送Post请求 
                Stream responseStream = response.GetResponseStream();

                byte[] byteFile = null;
                List<byte> bytes = new List<byte>();
                int temp = responseStream.ReadByte();
                while (temp != -1)
                {
                    bytes.Add((byte)temp);
                    temp = responseStream.ReadByte();
                }
                byteFile = bytes.ToArray();
                response.Close();
                responseStream.Close();

                Stream stream = new MemoryStream(byteFile);
                return stream;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        /// <summary>
        /// 上传文件到本地路径
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileType"></param>
        /// <param name="stream"></param>
        public static void UploadFileByStream(string fileName, string fileType, Stream stream)
        {
            //生成上传文件目录
            string uploadPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\upload";
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }

            FileStream pFileStream = null;
            fileName = fileName + "." + fileType;
            string pre = DateTime.Now.ToString("yyyyMMddHHssmm") + new Random().Next(10000);
            fileName = pre + "_" + fileName;
            string AttPath = uploadPath + "\\" + fileName;

            byte[] bArr = new byte[stream.Length];
            stream.Read(bArr, 0, bArr.Length);

            pFileStream = new FileStream(AttPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            pFileStream.Write(bArr, 0, bArr.Length);
            pFileStream.Close();
            stream.Close();
        }
    }
}
