using System;
using corelDraw = Corel.Interop.CorelDRAW;
using vgCore = Corel.Interop.VGCore;

namespace Export
{
    class Program
    {
        static void Main(string[] args)
        {
            vgCore.Application app = new corelDraw.Application();
            var activeDocument = app.CreateDocument();
            activeDocument.Unit = vgCore.cdrUnit.cdrPixel;

            AddGrayRectangle(activeDocument);

            //export like eps file
            activeDocument.Export($@"E:\{Guid.NewGuid().ToString("N")}.eps", vgCore.cdrFilter.cdrEPS);

            #region Export like pdf
            //activeDocument.Export($@"E:\{Guid.NewGuid().ToString("N")}.pdf", vgCore.cdrFilter.cdrPDF, vgCore.cdrExportRange.cdrAllPages);
            #endregion

            #region export like png
            // activeDocument.Export($@"E:\{Guid.NewGuid().ToString("N")}.png", vgCore.cdrFilter.cdrPNG);
            #endregion

            activeDocument.Close();
        }

        static void AddGrayRectangle(vgCore.Document activeDocument)
        {
            var activePage = activeDocument.ActivePage;
            var rectangle = activePage.ActiveLayer.CreateRectangle(10, 10, 1000, 1000);
            vgCore.Color fillColor = new vgCore.Color();
            fillColor.HexValue = "#CDCDCD";
            rectangle.Fill.UniformColor = fillColor;
        }
    }
}
