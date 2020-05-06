using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            //在当前项目目录的指定文件夹写日志
            LogHelper.WriteProgramLogInFolder("测试", "test");

            //在本地磁盘路径下写日志
            FileHelper.WriteTxt(@"C:\Error.txt", "test");

            //远程路径文件转Stream
            Stream stream = FileHelper.FileToStream("");

            //上传文件到本地磁盘
            FileHelper.UploadFileByStream("test", "doc", stream);
        }
    }
}
