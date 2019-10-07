using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tesseract;

namespace E7_Gear_Optimizer.Ocr
{
    public class TesseractService
    {
        public static string ParseText(Bitmap bmp, Rect region)
        {
            string text = null;

            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default, "e7.config"))
            {
                using (var page = engine.Process(bmp, region, PageSegMode.SingleColumn))
                {
                    text = page.GetText();
                }
            }
            
            return text;
        }
    }
}
