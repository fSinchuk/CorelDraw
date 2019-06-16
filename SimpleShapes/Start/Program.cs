using System;
using corelDraw = Corel.Interop.CorelDRAW;
using vgCore = Corel.Interop.VGCore;


namespace Start
{
    class Program
    {
        static void Main(string[] args)
        {
            vgCore.Application app = new corelDraw.Application();
            var activeDocument = app.CreateDocument();
            activeDocument.Unit = vgCore.cdrUnit.cdrPixel;

            AddGrayTriangle(activeDocument);

            activeDocument.SaveAs($@"E:\{Guid.NewGuid().ToString("N")}.cdr");
            activeDocument.Close();
        }

        static void AddGrayRectangle(vgCore.Document activeDocument) {
            var activePage = activeDocument.ActivePage;
            var rectangle = activePage.ActiveLayer.CreateRectangle(10, 10, 1000, 1000);
            vgCore.Color fillColor = new vgCore.Color();
            fillColor.HexValue = "#CDCDCD";
            rectangle.Fill.UniformColor = fillColor;
        }
        static void AddGrayCircle(vgCore.Document activeDocument)
        {
            int radius = 1000;

            var activePage = activeDocument.ActivePage;
            var rectangle = activePage.ActiveLayer.CreateRectangle(10, 10, 1000, 1000, radius, radius, radius, radius);
            vgCore.Color fillColor = new vgCore.Color();
            fillColor.HexValue = "#CDCDCD";
            rectangle.Fill.UniformColor = fillColor;
        }
        static void AddGrayTriangle(vgCore.Document activeDocument)
        {
            var activePage = activeDocument.ActivePage;
            var rectangle = activePage.ActiveLayer.CreatePolygon(100, 200, 200, 100,3);
            vgCore.Color fillColor = new vgCore.Color();
            fillColor.HexValue = "#CDCDCD";
            rectangle.Fill.UniformColor = fillColor;
        }
    }
}
