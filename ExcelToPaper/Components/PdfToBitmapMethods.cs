using Docnet.Core;
using Docnet.Core.Models;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ExcelToPaper.Components
{
    public static class PdfToBitmapMethods
    {
        public static IEnumerable<Bitmap> ToBitmaps(string filePath)
        {
            using (var library = DocLib.Instance)
            {
                using (var docReader = library.GetDocReader(filePath, new PageDimensions(540, 960)))
                {
                    for (int i = 0; i < docReader.GetPageCount(); i++)
                    {
                        using (var pageReader = docReader.GetPageReader(i))
                        {
                            var rawBytes = pageReader.GetImage();

                            var width = pageReader.GetPageWidth();
                            var height = pageReader.GetPageHeight();

                            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                            AddBytes(bmp, rawBytes);
                            //bmp.MakeTransparent();
                            yield return bmp;

                            //using (var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                            //{
                            //    AddBytes(bmp, rawBytes);
                            //    yield return bmp;
                            //}
                        }
                    }

                }
            }

        }

        private static void AddBytes(Bitmap bmp, byte[] rawBytes)
        {
            var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);

            var bmpData = bmp.LockBits(rect, ImageLockMode.WriteOnly, bmp.PixelFormat);
            var pNative = bmpData.Scan0;

            Marshal.Copy(rawBytes, 0, pNative, rawBytes.Length);
            bmp.UnlockBits(bmpData);
        }
    }
}
