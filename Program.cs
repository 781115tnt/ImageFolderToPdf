using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using IOPath = System.IO.Path;

using iText.Kernel.Pdf;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;

using System.Runtime.InteropServices;
using System.Net;
using iText.IO.Source;
using System.ComponentModel;
using System.Reflection.Metadata;
using static Program;
using iText.Kernel.Pdf.Xobject;
using iText.Commons.Utils;
using Org.BouncyCastle.Math.EC.Rfc7748;
using System.Runtime.CompilerServices;
using System.Xml;
using Org.BouncyCastle.Bcpg.OpenPgp;

class NaturalSortComparer : IComparer<string>
{
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    static extern int StrCmpLogicalW(string x, string y);

    public int Compare(string? x, string? y)
    {
        return StrCmpLogicalW(x ?? "", y ?? "");
    }
}

class Program
{
    const string PDF_EXTENSION = ".pdf";
    static readonly string[] _ImageExtensions =
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", PDF_EXTENSION
    };

    static readonly Dictionary<string, (float WidthMm, float HeightMm)> _PaperSizeDict =
    new(StringComparer.OrdinalIgnoreCase)
    {
        { "A5", (148, 210) },
        { "A4", (210, 297) },
        { "A3", (297, 420) },
        { "A2", (420, 594) },
        { "A1", (594, 841) },
        { "A0", (841, 1189) },
        { "2A0",(1189,1682) }
    };

    static Dictionary<string, List<int>> _PageSizeReport =
    new(StringComparer.OrdinalIgnoreCase);

    static Dictionary<int, string> _PageSizeByPage =
        new Dictionary<int, string>();

    static string _folder = "";
    static string _outputPdf = "";
    static float _marginMm = 8f;
    static float _numberOffsetMm = 4f;
    static bool _stretch = true;
    static float _autoSizeToleranceMm = 5f;
    static string? _standardPageSize = null;
    static float? _fitPageWidthMm = 210f;
    static string _fontPath = @"C:\Windows\Fonts\arial.ttf";
    static float _fontSize = 5f;

    static bool _usePageWidth => _fitPageWidthMm.HasValue;
    static bool _isOneToOne => _standardPageSize?.Equals("1_1", StringComparison.OrdinalIgnoreCase) == true;
    static bool _isAuto => _standardPageSize?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;
    static bool _allowStretch => _stretch && !_isOneToOne;
    static float _marginPts => MmToPts(_marginMm);
    static float _numberOffsetPts => MmToPts(_numberOffsetMm);

    static float _fitPageWidthPts => MmToPts(_fitPageWidthMm ?? 210);
    const float STANDARD_SIZE_TOLERANCE_MM = 2f;

    static List<string> supportedFiles = new List<string>();

    public class TMatrix
    {
        float a;
        float b;
        float c;
        float d;
        float e;
        float f;

        public TMatrix(float a, float b, float c, float d, float e, float f)
        {
            this.a = a;
            this.b = b;
            this.c = c;
            this.d = d;
            this.e = e;
            this.f = f;
        }

        public void ApplyMatrix(PdfCanvas canvas, ImageData imagedata)
        {
            canvas.AddImageWithTransformationMatrix(imagedata, a, b, c, d, e, f);
        }
        public void ApplyMatrix(PdfCanvas canvas, PdfFormXObject pageCopy)
        {
            canvas.AddXObjectWithTransformationMatrix(pageCopy, a, b, c, d, e, f);
        }
    }


    public class OrigPageGroup
    {
        const int SIZE_GROUP_TOLERANCE_PTS = 5;
        public OrigPageGroup() { OrigPageList = new List<OrigPage>(); }
        public List<OrigPage> OrigPageList { get; set; }
        public List<List<OrigPage>> GenerateGroups()
        {
            var groups = new List<List<OrigPage>>();
            foreach (var item in OrigPageList)
            {
                bool added = false;
                foreach (var group in groups)
                {
                    var representative = group[0];

                    if (Math.Abs(item.WidthPts - representative.WidthPts) <= SIZE_GROUP_TOLERANCE_PTS &&
                        Math.Abs(item.HeightPts - representative.HeightPts) <= SIZE_GROUP_TOLERANCE_PTS)
                    {
                        group.Add(item);
                        added = true;
                        break;
                    }
                }
                if (!added) groups.Add(new List<OrigPage> { item });
            }
            return groups;
        }
        public void WriteDebug()
        {
            int groupIndex = 1;
            List<List<OrigPage>> groups = GenerateGroups();
            foreach (var group in groups)
            {
                Console.WriteLine($"Group {groupIndex++}: {GetSizeLabel(group[0].WidthPts, group[0].HeightPts)}");
                foreach (var item in group) { Console.WriteLine("   " + item); }
            }
        }
    }
    public class OrigPage
    {
        public float WidthPts { get; set; }
        private float WidthMm => PtsToMm(WidthPts);
        public float HeightPts { get; set; }
        private float HeightMm => PtsToMm(HeightPts);
        public int PageNumber { get; set; }
        public string FilePath { get; set; }
        public int OrigPdfPageNumber { get; set; }

        public float OrigWidthPts { get; set; }
        private float OrigWidthMm => PtsToMm(OrigWidthPts);
        public float OrigHeightPts { get; set; }
        private float OrigHeightMm => PtsToMm(OrigHeightPts);
        public TMatrix TransformMatrix { get; set; }

        public float llx { get; set; }
        public float lly { get; set; }
        public OrigPage(float OrigWidthPts, float OrigHeightPts, float WidthPts, float HeightPts, int PageNumber, string FilePath, int OrigPdfPageNumber, float llx, float lly, TMatrix TransformMatrix, OrigPageGroup group)
        {
            this.OrigWidthPts = OrigWidthPts;
            this.OrigHeightPts = OrigHeightPts;
            this.WidthPts = WidthPts;
            this.HeightPts = HeightPts;
            this.PageNumber = PageNumber;
            this.FilePath = FilePath;
            this.OrigPdfPageNumber = OrigPdfPageNumber;
            this.TransformMatrix = TransformMatrix;
            this.llx = llx;
            this.lly = lly;
            group.OrigPageList.Add(this);
        }
        public override string ToString()
        {
            return $"W={WidthPts}, H={HeightPts}, OW={OrigWidthMm}, OH={OrigHeightMm}, PN={PageNumber}, File={FilePath}, OrigPdfPN={OrigPdfPageNumber}";
        }
    }

    static OrigPageGroup AnalyzePages()
    {
        OrigPageGroup origPageGroup = new OrigPageGroup();

        int pageNumber = 0;

        foreach (var file in supportedFiles)
        {
            if (file.EndsWith(PDF_EXTENSION) == false)
            {
                //Images
                var imageData = ImageDataFactory.Create(file);

                float pixelX = imageData.GetWidth();
                float pixelY = imageData.GetHeight();

                float dpiX = imageData.GetDpiX() > 0 ? imageData.GetDpiX() : 72;
                float dpiY = imageData.GetDpiY() > 0 ? imageData.GetDpiY() : 72;

                float imgXPts = pixelX * 72f / dpiX;
                float imgYPts = pixelY * 72f / dpiY;

                float pageWidthPts;
                float pageHeightPts;
                TMatrix transformMatrix;
                CalculatePageSizeForImage(imgXPts, imgYPts, out pageWidthPts, out pageHeightPts, out transformMatrix);
                pageNumber++;
                new OrigPage(imgXPts, imgYPts, pageWidthPts, pageHeightPts, pageNumber, file, 1, 0f, 0f, transformMatrix, origPageGroup);
            }
            else
            {
                //PDF
                using var src = new PdfDocument(new PdfReader(file));
                for (int i = 1; i <= src.GetNumberOfPages(); i++)
                {
                    var srcPage = src.GetPage(i);

                    Rectangle bbox = srcPage.GetCropBox();

                    float imgXPts = bbox.GetWidth();
                    float imgYPts = bbox.GetHeight();
                    float llx = bbox.GetX();
                    float lly = bbox.GetY();

                    float pageWidthPts;
                    float pageHeightPts;
                    TMatrix transformMatrix;
                    CalculatePageSizeForPdf(imgXPts, imgYPts, out pageWidthPts, out pageHeightPts, out transformMatrix, llx, lly);
                    pageNumber++;

                    new OrigPage(imgXPts, imgYPts, pageWidthPts, pageHeightPts, pageNumber, file, i,llx, lly, transformMatrix, origPageGroup);
                }
            }
        }

        return origPageGroup;
    }

    static int Main(string[] args)
    {
        var options = ParseArguments(args);

        if (!options.ContainsKey("command") || options["command"] != "file2pdf")
        {
            Console.WriteLine("Usage: -command=file2pdf -input=folder -output=file.pdf");
            return 1;
        }

        if (!options.ContainsKey("input") || !options.ContainsKey("output"))
        {
            Console.WriteLine("Missing -input or -output");
            return 1;
        }

        _folder = options["input"];
        _outputPdf = options["output"];

        _marginMm = options.ContainsKey("margin") ? float.Parse(options["margin"]) : 8f;
        _numberOffsetMm = options.ContainsKey("numberoffset") ? float.Parse(options["numberoffset"]) : 4f;
        _stretch = options.ContainsKey("stretch") && options["stretch"].ToLower() == "y";
        _autoSizeToleranceMm = options.ContainsKey("autosizetolerance") ? float.Parse(options["autosizetolerance"]) : 5f;
        // If not specify any thing -> fallback to A4
        _standardPageSize = options.ContainsKey("pagesize") ? options["pagesize"] : null;
        _fitPageWidthMm = options.ContainsKey("pagewidth") ? float.Parse(options["pagewidth"]) : null;

        _fontPath = @"C:\Windows\Fonts\arial.ttf";
        _fontSize = 5f;

        // NonRecursive
        supportedFiles = Directory.EnumerateFiles(_folder)
            .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
            .ToList();
        // Recursive
        // supportedFiles = Directory.EnumerateFiles(_folder, "*.*", SearchOption.AllDirectories)
        //     .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
        //     .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
        //     .ToList();




        //CreatePdfFromFiles_NoFormat(false);

        OrigPageGroup origPageGroup = AnalyzePages();

        CreatePdfFromFiles(origPageGroup, true);

        //CreatePdfFromFiles_Group(origPageGroup, true);

        origPageGroup.WriteDebug();


        Console.WriteLine("Done.");
        return 0;
    }





    static void AddPageFromImage(OrigPage origPage, PdfDocument pdf, PdfFont font, bool addPageNumber = true)
    {

        var imageData = ImageDataFactory.Create(origPage.FilePath);

        var page = pdf.AddNewPage(new PageSize(origPage.WidthPts, origPage.HeightPts));

        //RegisterPageSize(pageWidth, pageHeight, pageNumber);

        var canvas = new PdfCanvas(page);
        origPage.TransformMatrix.ApplyMatrix(canvas, imageData);
        if (addPageNumber) AddPageNumber(canvas, font, _numberOffsetPts, origPage.PageNumber);
    }

    static void AddPageFromPdf(OrigPage origPage, PdfDocument pdf, PdfFont font, bool addPageNumber = true)
    {
        using var src = new PdfDocument(new PdfReader(origPage.FilePath));
        var srcPage = src.GetPage(origPage.OrigPdfPageNumber);


        int rot = srcPage.GetRotation();
        Rectangle crop = srcPage.GetCropBox();
        srcPage.SetRotation(0);
        var pageCopy = srcPage.CopyAsFormXObject(pdf);
        var newPage = pdf.AddNewPage(new PageSize(origPage.WidthPts, origPage.HeightPts));

        var canvas = new PdfCanvas(newPage);

        origPage.TransformMatrix.ApplyMatrix(canvas, pageCopy);

        if (addPageNumber) AddPageNumber(canvas, font, _numberOffsetPts, origPage.PageNumber);
    }
    static void CreatePdfFromFiles(OrigPageGroup origPageGroup, bool addPageNumber = true)
    {

        using var writer = new PdfWriter(_outputPdf);
        using var pdf = new PdfDocument(writer);

        PdfFont font = PdfFontFactory.CreateFont(
            _fontPath,
            PdfEncodings.WINANSI,
            PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED
        );

        var pageGroupLists = origPageGroup.GenerateGroups();
        foreach (var list in pageGroupLists)
        {
            foreach (var op in list)
            {
                if (IOPath.GetExtension(op.FilePath) != PDF_EXTENSION)
                {
                    AddPageFromImage(op, pdf, font, true);
                }
                else
                {
                    AddPageFromPdf(op, pdf, font, true);
                }
            }
        }
    }



    static TMatrix CreatePdfTransformMatrix(float origXPts, float origYPts, float x, float y, float w, float h, float llx, float lly)
    {
        //         y+h
        //  ↑
        //  │
        //  │        +-------------------+
        //  │        |                   |
        //  │        |       IMAGE       |  height = h
        //  │        |                   |
        //  │        +-------------------+
        //  │        (x, y)
        //  └────────────────────────────────→
        //                 x            x+w
        //Portrait
        if (origXPts < origYPts)
        {
            float sx = w / origXPts;
            float sy = h / origYPts;
float e = x - llx * sx;
float f = y - lly * sy;
            return new TMatrix(
                sx, 0,
                0, sy,
                e, f
            );
        }
        else
        {
            
            float sx = w / origYPts;
            float sy = h / origXPts;
            float e = x + w + (lly * sy);
float f = y - (llx * sx);

            return new TMatrix(
                0, sy,
                -sx, 0,
                e, f
            );
        }
    }
    static TMatrix CreateImageTransformMatrix(float x, float y, float w, float h, bool isImageLandscape)
    {
        //         y+h
        //  ↑
        //  │
        //  │        +-------------------+
        //  │        |                   |
        //  │        |       IMAGE       |  height = h
        //  │        |                   |
        //  │        +-------------------+
        //  │        (x, y)
        //  └────────────────────────────────→
        //                 x            x+w

        //Portrait
        if (!isImageLandscape)
        {
            return new TMatrix(
                w, 0,
                0, h,
                x, y
            );
        }
        else
        {
            return new TMatrix(
                0, h,
                -w, 0,
                x + w,
                y
            );
        }
    }



    static void CalculatePageSizeForPdf(float imgXPts, float imgYPts, out float pageWidthPts, out float pageHeightPts, out TMatrix TransformMatrix, float llx, float lly)
    {
        //Note: imgXPts, imgYPtx is X, Y lenghts
        //x,y placement location calculate from bottom left
        //w,h width and height of the placement area. w is horizontal and h is vertical.
        //All are in pts
        float x, y, w, h;

        float imgShortPts = Math.Min(imgXPts, imgYPts);
        float imgLongPts = Math.Max(imgXPts, imgYPts);
        float ratio = imgLongPts / imgShortPts;


        if (_usePageWidth)
        {
            pageWidthPts = _fitPageWidthPts;
            pageHeightPts = _fitPageWidthPts * ratio;

            // Calculate matrix
            x = _marginPts;
            w = pageWidthPts - 2 * _marginPts;
            if (_allowStretch)
            {
                y = _marginPts;
                h = pageHeightPts - 2 * _marginPts;
            }
            else
            {
                h = ratio * w;
                y = (pageHeightPts - h) / 2;
            }

        }
        else if (_isOneToOne)
        {
            pageWidthPts = imgShortPts + (_marginPts * 2);
            pageHeightPts = imgLongPts + (_marginPts * 2);

            x = _marginPts;
            y = _marginPts;
            w = imgShortPts;
            h = imgLongPts;
        }
        else if (!string.IsNullOrEmpty(_standardPageSize) && !_isAuto)
        {
            //Specify wrong standard paper size name
            if (!_PaperSizeDict.ContainsKey(_standardPageSize))
                throw new Exception($"Unsupported page size: {_standardPageSize}");
            //Get user specified paper
            var paper = _PaperSizeDict[_standardPageSize];

            pageWidthPts = MmToPts(Math.Min(paper.WidthMm, paper.HeightMm));
            pageHeightPts = MmToPts(Math.Max(paper.WidthMm, paper.HeightMm));

            x = _marginPts;
            w = pageWidthPts - 2 * _marginPts;
            if (_allowStretch)
            {
                y = _marginPts;
                h = pageHeightPts - 2 * _marginPts;
            }
            else
            {
                h = ratio * w;
                y = (pageHeightPts - h) / 2;
            }
        }
        else
        {
            //Not specified => pagesize is equal to original size
            pageWidthPts = Math.Min(imgXPts, imgYPts);
            pageHeightPts = Math.Max(imgXPts, imgYPts);

            //If auto => detect standard size and recalculate paper size
            if (_isAuto)
            {
                foreach (var paper in _PaperSizeDict)
                {
                    float paperW = MmToPts(paper.Value.WidthMm);
                    float paperH = MmToPts(paper.Value.HeightMm);

                    if (MatchStandardPageSize(pageWidthPts, pageHeightPts, paperW, paperH, _autoSizeToleranceMm))
                    {
                        pageWidthPts = Math.Min(paperW, paperH);
                        pageHeightPts = Math.Max(paperW, paperH);
                        break;
                    }
                }
            }
            //Until here paper size are re-calculated if autosize detected
            x = _marginPts;
            w = pageWidthPts - 2 * _marginPts;
            if (_allowStretch)
            {
                y = _marginPts;
                h = pageHeightPts - 2 * _marginPts;
            }
            else
            {
                h = ratio * w;
                y = (pageHeightPts - h) / 2;
            }
        }
        //Until here all x, y, w, h are calculated.  imgXPts > imgYPts => landscape
        TransformMatrix = CreatePdfTransformMatrix(imgXPts, imgYPts, x, y, w, h, llx, lly);
    }



    static void CalculatePageSizeForImage(float imgXPts, float imgYPts, out float pageWidthPts, out float pageHeightPts, out TMatrix TransformMatrix)
    {
        //Note: imgXPts, imgYPtx is X, Y lenghts
        //x,y placement location calculate from bottom left
        //w,h width and height of the placement area. w is horizontal and h is vertical.
        //All are in pts
        float x, y, w, h;

        float imgShortPts = Math.Min(imgXPts, imgYPts);
        float imgLongPts = Math.Max(imgXPts, imgYPts);
        float ratio = imgLongPts / imgShortPts;


        if (_usePageWidth)
        {
            pageWidthPts = _fitPageWidthPts;
            pageHeightPts = _fitPageWidthPts * ratio;

            // Calculate matrix
            x = _marginPts;
            w = pageWidthPts - 2 * _marginPts;
            if (_allowStretch)
            {
                y = _marginPts;
                h = pageHeightPts - 2 * _marginPts;
            }
            else
            {
                h = ratio * w;
                y = (pageHeightPts - h) / 2;
            }

        }
        else if (_isOneToOne)
        {
            pageWidthPts = imgShortPts + (_marginPts * 2);
            pageHeightPts = imgLongPts + (_marginPts * 2);

            x = _marginPts;
            y = _marginPts;
            w = imgShortPts;
            h = imgLongPts;
        }
        else if (!string.IsNullOrEmpty(_standardPageSize) && !_isAuto)
        {
            //Specify wrong standard paper size name
            if (!_PaperSizeDict.ContainsKey(_standardPageSize))
                throw new Exception($"Unsupported page size: {_standardPageSize}");
            //Get user specified paper
            var paper = _PaperSizeDict[_standardPageSize];

            pageWidthPts = MmToPts(Math.Min(paper.WidthMm, paper.HeightMm));
            pageHeightPts = MmToPts(Math.Max(paper.WidthMm, paper.HeightMm));

            x = _marginPts;
            w = pageWidthPts - 2 * _marginPts;
            if (_allowStretch)
            {
                y = _marginPts;
                h = pageHeightPts - 2 * _marginPts;
            }
            else
            {
                h = ratio * w;
                y = (pageHeightPts - h) / 2;
            }
        }
        else
        {
            //Not specified => pagesize is equal to original size
            pageWidthPts = Math.Min(imgXPts, imgYPts);
            pageHeightPts = Math.Max(imgXPts, imgYPts);

            //If auto => detect standard size and recalculate paper size
            if (_isAuto)
            {
                foreach (var paper in _PaperSizeDict)
                {
                    float paperW = MmToPts(paper.Value.WidthMm);
                    float paperH = MmToPts(paper.Value.HeightMm);

                    if (MatchStandardPageSize(pageWidthPts, pageHeightPts, paperW, paperH, _autoSizeToleranceMm))
                    {
                        pageWidthPts = Math.Min(paperW, paperH);
                        pageHeightPts = Math.Max(paperW, paperH);
                        break;
                    }
                }
            }
            //Until here paper size are re-calculated if autosize detected
            x = _marginPts;
            w = pageWidthPts - 2 * _marginPts;
            if (_allowStretch)
            {
                y = _marginPts;
                h = pageHeightPts - 2 * _marginPts;
            }
            else
            {
                h = ratio * w;
                y = (pageHeightPts - h) / 2;
            }
        }
        //Until here all x, y, w, h are calculated.  imgXPts > imgYPts => landscape
        TransformMatrix = CreateImageTransformMatrix(x, y, w, h, imgXPts > imgYPts);
    }








    static void AddPageNumber(PdfCanvas canvas, PdfFont font, float numberOffset, int pageNumber)
    {
        canvas.BeginText();
        canvas.SetFontAndSize(font, _fontSize);
        canvas.SetFillColor(ColorConstants.GRAY);
        canvas.MoveText(numberOffset, numberOffset);
        canvas.ShowText(pageNumber.ToString());
        canvas.EndText();
    }



    static string GetSizeLabel(float widthPts, float heightPts)
    {
        float wMm = widthPts * 25.4f / 72f;
        float hMm = heightPts * 25.4f / 72f;

        float shortSide = Math.Min(wMm, hMm);
        float longSide = Math.Max(wMm, hMm);

        string sizeLabel = "";

        foreach (var kv in _PaperSizeDict)
        {
            float paperShort = Math.Min(kv.Value.WidthMm, kv.Value.HeightMm);
            float paperLong = Math.Max(kv.Value.WidthMm, kv.Value.HeightMm);

            if (Math.Abs(shortSide - paperShort) <= STANDARD_SIZE_TOLERANCE_MM &&
                Math.Abs(longSide - paperLong) <= STANDARD_SIZE_TOLERANCE_MM)
            {
                sizeLabel = $"{kv.Key} ({paperShort}x{paperLong} mm)";
                break;
            }
            else sizeLabel = $"{Math.Round(shortSide)}x{Math.Round(longSide)} mm";
        }
        return sizeLabel;
    }

    static bool MatchStandardPageSize(float imageWPts, float imageHPts, float paperWPts, float paperHPts, float percentThreshold)
    {
        float imgShort = Math.Min(imageWPts, imageHPts);
        float imgLong = Math.Max(imageWPts, imageHPts);

        float papShort = Math.Min(paperWPts, paperHPts);
        float papLong = Math.Max(paperWPts, paperHPts);

        float shortDiff = Math.Abs(imgShort - papShort) / papShort * 100f;
        float longDiff = Math.Abs(imgLong - papLong) / papLong * 100f;

        return shortDiff <= percentThreshold &&
               longDiff <= percentThreshold;
    }

    static float MmToPts(float mm)
        => mm * 72f / 25.4f;
    static float PtsToMm(float points)
=> points * 25.4f / 72f;
    static Dictionary<string, string> ParseArguments(string[] args)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var arg in args)
        {
            if (!arg.StartsWith("-")) continue;
            var parts = arg.Substring(1).Split('=', 2);
            if (parts.Length == 2)
                dict[parts[0]] = parts[1].Trim('"');
        }

        return dict;
    }

}










