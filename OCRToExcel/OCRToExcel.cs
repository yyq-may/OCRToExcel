using Aspose.Cells;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace OCRToExcel
{
    public class OCRToExcel : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]        
        public InArgument<string> ActionUrl { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ImagePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> OutExcelPath { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string imagePath = ImagePath.Get(context);
            string actionUrl = ActionUrl.Get(context);
            string outPath = OutExcelPath.Get(context);
            string result = TaskAsync(actionUrl, imagePath).Result;
            
            Workbook workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            
            var json = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(result);
            int n = 1;
            foreach (var item in json)
            {
                worksheet.Cells["A"+n+""].PutValue(item);
                n++;
            }
            workbook.Save(outPath);
            worksheet.Dispose();
            workbook.Dispose();

        }

        public static async Task<string> TaskAsync(string actionUrl, string imageFilePath)
        {
            using (var client = new HttpClient())
            {
                Image img = new Image();
                img.image = ImgToBase64(imageFilePath);
                var str = Newtonsoft.Json.JsonConvert.SerializeObject(img);
                HttpContent content = new StringContent(str);
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
                HttpResponseMessage response = await client.PostAsync(actionUrl, content);
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();
                return responseBody;
            }
        }

        

        public static string ImgToBase64(string fileName)
        {
            FileStream filestream = new FileStream(fileName, FileMode.Open);
            byte[] arr = new byte[filestream.Length];
            filestream.Read(arr, 0, (int)filestream.Length);
            string baser64 = Convert.ToBase64String(arr);
            filestream.Close();
            return baser64;
        }
        public class Image
        {
            public string image;
        }
    }
}
