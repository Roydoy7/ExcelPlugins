using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Imaging;

namespace ExcelToPaper.Components
{
    public static class BitmapToImgSrcMethods
    {
        public static BitmapImage ToImageSource(this Bitmap bitmap)
        {
            //using (var memory = new MemoryStream())
            //{
            //    bitmap.Save(memory, ImageFormat.Bmp);
            //    memory.Position = 0;
            //    BitmapImage bitmapimage = new BitmapImage();
            //    bitmapimage.BeginInit();
            //    bitmapimage.StreamSource = memory;
            //    bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
            //    bitmapimage.EndInit();

            //    return bitmapimage;
            //}

            using (MemoryStream stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png); // Was .Bmp, but this did not show a transparent background.

                stream.Position = 0;
                BitmapImage result = new BitmapImage();
                result.BeginInit();
                // According to MSDN, "The default OnDemand cache option retains access to the stream until the image is needed."
                // Force the bitmap to load right now so we can dispose the stream.
                result.CacheOption = BitmapCacheOption.OnLoad;
                result.StreamSource = stream;
                result.EndInit();
                result.Freeze();
                return result;
            }
        }
    }
}
