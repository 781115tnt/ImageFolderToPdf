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
    static string _standardPageSize = "A4";
    static float? _fitPageWidthMm = 210f;
    static string _fontPath = @"C:\Windows\Fonts\arial.ttf";
    static float _fontSize = 5f;

    static bool _usePageWidth => _fitPageWidthMm.HasValue;
    static bool _isOneToOne => _standardPageSize?.Equals("1_1", StringComparison.OrdinalIgnoreCase) == true;
    static bool _isAuto => _standardPageSize?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;
    static bool _allowStretch => _stretch && !_isOneToOne;
    static float _marginPts => MmToPoints(_marginMm);
    static float _numberOffsetPts => MmToPoints(_numberOffsetMm);

    static float _fitPageWidthPts => MmToPoints(_fitPageWidthMm ?? 210);

    static OrigPageGroup _origPageGroup = new OrigPageGroup();

    public class OrigPageGroup
    {
        const int SIZE_DELTA = 5;
        public OrigPageGroup() { OrigPageList = new List<OrigPage>(); }
        public List<OrigPage> OrigPageList { get; set; }
        public List<List<OrigPage>> GroupPageByTolerance => GenerateGroups();
        private List<List<OrigPage>> GenerateGroups()
        {
            var groups = new List<List<OrigPage>>();
            foreach (var item in OrigPageList)
            {
                bool added = false;
                foreach (var group in groups)
                {
                    var representative = group[0];

                    if (Math.Abs(item.WidthPts - representative.WidthPts) <= SIZE_DELTA &&
                        Math.Abs(item.HeightPts - representative.HeightPts) <= SIZE_DELTA)
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
                Console.WriteLine($"Group {groupIndex++}:");
                foreach (var item in group) { Console.WriteLine("   " + item); }
            }
        }
    }
    public class OrigPage
    {
        public int WidthPts { get; set; }
        public int WidthMm => (int) (WidthPts * 0.35278f);
        public int HeightPts { get; set; }
        public int HeightMm => (int) (HeightPts * 0.35278f);
        public int PageNumber { get; set; }
        public string FilePath { get; set; }
        public int OrigPdfPageNumber { get; set; }

        public OrigPage(int WidthPts, int HeightPts, int PageNumber, string FilePath, int OrigPdfPageNumber, OrigPageGroup group)
        {
            this.WidthPts = WidthPts;
            this.HeightPts = HeightPts;
            this.PageNumber = PageNumber;
            this.FilePath = FilePath;
            this.OrigPdfPageNumber = OrigPdfPageNumber;
            group.OrigPageList.Add(this);
        }
        public override string ToString()
        {
            return $"W={WidthMm}, H={HeightMm}, PN={PageNumber}, File={FilePath}, OrigPdfPN={OrigPdfPageNumber}";
        }
    }


    static void AnalyzePages()
    {
        var supportedFiles = Directory
            .EnumerateFiles(_folder)
            .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
            .ToList();

        int pageNumber = 0;

        foreach (var file in supportedFiles)
        {
            if (file.EndsWith(PDF_EXTENSION) == false)
            {
                //Images
                var imageData = ImageDataFactory.Create(file);

                float pixelWidth = imageData.GetWidth();
                float pixelHeight = imageData.GetHeight();

                float dpiX = imageData.GetDpiX() > 0 ? imageData.GetDpiX() : 72;
                float dpiY = imageData.GetDpiY() > 0 ? imageData.GetDpiY() : 72;

                float imageWidthPts = pixelWidth * 72f / dpiX;
                float imageHeightPts = pixelHeight * 72f / dpiY;

                float pageWidth;
                float pageHeight;
                CalculatePageSize(ref imageWidthPts, ref imageHeightPts, out pageWidth, out pageHeight);
                pageNumber++;
                new OrigPage((int)pageWidth, (int)pageHeight, pageNumber, file, 1, _origPageGroup);
            }
            else
            {
                //PDF
                using var src = new PdfDocument(new PdfReader(file));
                for (int i = 1; i <= src.GetNumberOfPages(); i++)
                {
                    var srcPage = src.GetPage(i);

                    Rectangle bbox = srcPage.GetCropBox();

                    float imageWidthPts = bbox.GetWidth();
                    float imageHeightPts = bbox.GetHeight();

                    float pageWidth;
                    float pageHeight;

                    CalculatePageSize(ref imageWidthPts, ref imageHeightPts, out pageWidth, out pageHeight);
                    pageNumber++;

                    new OrigPage((int)pageWidth, (int)pageHeight, pageNumber, file, i, _origPageGroup);
                }
            }
        }
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
        _standardPageSize = options.ContainsKey("pagesize") ? options["pagesize"] : "A4";
        _fitPageWidthMm = options.ContainsKey("pagewidth") ? float.Parse(options["pagewidth"]) : null;

        _fontPath = @"C:\Windows\Fonts\arial.ttf";
        _fontSize = 5f;



        AnalyzePages();

        //CreatePdfFromFiles(true);


        //CreatePdfFromFilesNoFormat(false);


        _origPageGroup.WriteDebug();



        Console.WriteLine("Done.");
        return 0;
    }

    static void RegisterPageSize(float widthPts, float heightPts, int pageNumber)
    {
        float wMm = widthPts * 25.4f / 72f;
        float hMm = heightPts * 25.4f / 72f;

        float shortSide = Math.Min(wMm, hMm);
        float longSide = Math.Max(wMm, hMm);

        const float toleranceMm = 2f;

        string? sizeLabel = null;

        foreach (var kv in _PaperSizeDict)
        {
            float paperShort = Math.Min(kv.Value.WidthMm, kv.Value.HeightMm);
            float paperLong = Math.Max(kv.Value.WidthMm, kv.Value.HeightMm);

            if (Math.Abs(shortSide - paperShort) <= toleranceMm &&
                Math.Abs(longSide - paperLong) <= toleranceMm)
            {
                sizeLabel = $"{kv.Key} ({paperShort}x{paperLong} mm)";
                break;
            }
        }

        if (sizeLabel == null)
            sizeLabel = $"{Math.Round(shortSide)}x{Math.Round(longSide)} mm";

        // Grouped storage
        if (!_PageSizeReport.ContainsKey(sizeLabel))
            _PageSizeReport[sizeLabel] = new List<int>();

        _PageSizeReport[sizeLabel].Add(pageNumber);

        // Page-by-page storage
        _PageSizeByPage[pageNumber] = sizeLabel;
    }


    static void WritePageSizeReport(string pdfPath)
    {
        string txtPath = IOPath.ChangeExtension(pdfPath, ".txt");

        using var writer = new StreamWriter(txtPath);

        writer.WriteLine("PAGE SIZE REPORT");
        writer.WriteLine("================");
        writer.WriteLine();

        // --- Grouped section ---
        foreach (var kv in _PageSizeReport.OrderBy(k => k.Key))
        {
            writer.WriteLine($"Page Size: {kv.Key}");
            writer.WriteLine($"Total Pages: {kv.Value.Count}");
            writer.WriteLine($"Pages: {string.Join(", ", kv.Value)}");
            writer.WriteLine();
        }

        // --- Detailed page mapping section ---
        writer.WriteLine("PAGE → SIZE");
        writer.WriteLine("===========");
        writer.WriteLine();

        foreach (var kv in _PageSizeByPage.OrderBy(k => k.Key))
        {
            writer.WriteLine($"{kv.Key} - {kv.Value}");
        }
    }


    static void CalculatePageSize(ref float imageWidthPts, ref float imageHeightPts, out float pageWidthPts, out float pageHeightPts)
    {
        if (_usePageWidth)
        {
            float imgShort = Math.Min(imageWidthPts, imageHeightPts);
            float imgLong = Math.Max(imageWidthPts, imageHeightPts);
            float ratio = imgLong / imgShort;

            pageWidthPts = _fitPageWidthPts;
            pageHeightPts = _fitPageWidthPts * ratio;
        }
        else if (_isOneToOne)
        {
            float imgShort = Math.Min(imageWidthPts, imageHeightPts);
            float imgLong = Math.Max(imageWidthPts, imageHeightPts);

            pageWidthPts = imgShort + (_marginPts * 2);
            pageHeightPts = imgLong + (_marginPts * 2);
        }
        else if (!string.IsNullOrEmpty(_standardPageSize) && !_isAuto)
        {

            if (!_PaperSizeDict.ContainsKey(_standardPageSize))
                throw new Exception($"Unsupported page size: {_standardPageSize}");

            var paper = _PaperSizeDict[_standardPageSize];

            pageWidthPts = MmToPoints(Math.Min(paper.WidthMm, paper.HeightMm));
            pageHeightPts = MmToPoints(Math.Max(paper.WidthMm, paper.HeightMm));
        }
        else
        {
            pageWidthPts = Math.Min(imageWidthPts, imageHeightPts);
            pageHeightPts = Math.Max(imageWidthPts, imageHeightPts);

            if (_isAuto)
            {
                foreach (var paper in _PaperSizeDict)
                {
                    float paperW = MmToPoints(paper.Value.WidthMm);
                    float paperH = MmToPoints(paper.Value.HeightMm);

                    if (IsPaperMatch(pageWidthPts, pageHeightPts, paperW, paperH, _autoSizeToleranceMm))
                    {
                        pageWidthPts = Math.Min(paperW, paperH);
                        pageHeightPts = Math.Max(paperW, paperH);
                        break;
                    }
                }
            }
        }

    }

    static void AddPagesFromPdf(string file, ref int pageNumber, PdfWriter writer, PdfDocument pdf, PdfFont font, bool addPageNumber = true)
    {
        using var src = new PdfDocument(new PdfReader(file));

        for (int i = 1; i <= src.GetNumberOfPages(); i++)
        {
            var srcPage = src.GetPage(i);

            // Copy page first so we know the real drawable size
            var pageCopy = srcPage.CopyAsFormXObject(pdf);

            Rectangle bbox = pageCopy.GetBBox().ToRectangle();

            float imageWidthPts = bbox.GetWidth();
            float imageHeightPts = bbox.GetHeight();

            float pageWidth;
            float pageHeight;

            CalculatePageSize(ref imageWidthPts, ref imageHeightPts, out pageWidth, out pageHeight);

            var newPage = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
            pageNumber++;


            RegisterPageSize(pageWidth, pageHeight, pageNumber);


            var canvas = new PdfCanvas(newPage);

            float availableWidth = pageWidth - (_marginPts * 2);
            float availableHeight = pageHeight - (_marginPts * 2);

            float drawWidth;
            float drawHeight;

            if (!_allowStretch)
            {
                float scale = Math.Min(
                    availableWidth / imageWidthPts,
                    availableHeight / imageHeightPts);

                drawWidth = imageWidthPts * scale;
                drawHeight = imageHeightPts * scale;
            }
            else
            {
                drawWidth = availableWidth;
                drawHeight = availableHeight;
            }

            float x = (pageWidth - drawWidth) / 2;
            float y = (pageHeight - drawHeight) / 2;

            canvas.AddXObjectFittedIntoRectangle(
                pageCopy,
                new Rectangle(x, y, drawWidth, drawHeight)
            );

            if (addPageNumber) AddPageNumber(canvas, font, _numberOffsetPts, pageNumber);
        }
    }

    static void AddPageFromImage(string image, ref int pageNumber, PdfWriter writer, PdfDocument pdf, PdfFont font, bool addPageNumber = true)
    {

        var imageData = ImageDataFactory.Create(image);

        float pixelWidth = imageData.GetWidth();
        float pixelHeight = imageData.GetHeight();

        float dpiX = imageData.GetDpiX() > 0 ? imageData.GetDpiX() : 72;
        float dpiY = imageData.GetDpiY() > 0 ? imageData.GetDpiY() : 72;

        float imageWidthPts = pixelWidth * 72f / dpiX;
        float imageHeightPts = pixelHeight * 72f / dpiY;



        float pageWidth;
        float pageHeight;
        CalculatePageSize(ref imageWidthPts, ref imageHeightPts, out pageWidth, out pageHeight);


        var page = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
        pageNumber++;



        RegisterPageSize(pageWidth, pageHeight, pageNumber);



        var canvas = new PdfCanvas(page);

        bool isLandscape = imageWidthPts > imageHeightPts;

        float availableWidth = pageWidth - (_marginPts * 2);
        float availableHeight = pageHeight - (_marginPts * 2);

        float drawWidth;
        float drawHeight;

        if (!isLandscape)
        {
            if (!_allowStretch)
            {
                float scale = Math.Min(
                    availableWidth / imageWidthPts,
                    availableHeight / imageHeightPts);

                drawWidth = imageWidthPts * scale;
                drawHeight = imageHeightPts * scale;
            }
            else
            {
                drawWidth = availableWidth;
                drawHeight = availableHeight;
            }

            float x = (pageWidth - drawWidth) / 2;
            float y = (pageHeight - drawHeight) / 2;

            canvas.AddImageFittedIntoRectangle(
                imageData,
                new Rectangle(x, y, drawWidth, drawHeight),
                false);
        }
        else
        {
            float rotatedW = imageHeightPts;
            float rotatedH = imageWidthPts;

            if (!_allowStretch)
            {
                float scale = Math.Min(
                    availableWidth / rotatedW,
                    availableHeight / rotatedH);

                drawWidth = rotatedW * scale;
                drawHeight = rotatedH * scale;
            }
            else
            {
                drawWidth = availableWidth;
                drawHeight = availableHeight;
            }

            float x = (pageWidth - drawWidth) / 2;
            float y = (pageHeight - drawHeight) / 2;

            canvas.SaveState();
            canvas.ConcatMatrix(0, 1, -1, 0, x + drawWidth, y);

            canvas.AddImageFittedIntoRectangle(
                imageData,
                new Rectangle(0, 0, drawHeight, drawWidth),
                false);

            canvas.RestoreState();
        }
        if (addPageNumber) AddPageNumber(canvas, font, _numberOffsetPts, pageNumber);
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

    static void CreatePdfFromFiles(bool addPageNumber = true)
    {
        var supportedFiles = Directory
            .EnumerateFiles(_folder)
            .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
            .ToList();

        using var writer = new PdfWriter(_outputPdf);
        using var pdf = new PdfDocument(writer);

        PdfFont font = PdfFontFactory.CreateFont(
            _fontPath,
            PdfEncodings.WINANSI,
            PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED
        );

        int pageNumber = 0;

        foreach (var file in supportedFiles)
        {
            if (file.EndsWith(PDF_EXTENSION) == false) AddPageFromImage(file, ref pageNumber, writer, pdf, font, addPageNumber);
            else AddPagesFromPdf(file, ref pageNumber, writer, pdf, font, addPageNumber);
        }
        WritePageSizeReport(_outputPdf);
    }

    static void CreatePdfFromFilesNoFormat(bool autoPortrait = false)
    {
        var supportedFiles = Directory
            .EnumerateFiles(_folder)
            .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
            .ToList();

        string outputNoFormatPdf = IOPath.ChangeExtension(_outputPdf, null) + "_NO_FORMAT.pdf";
        using var writer = new PdfWriter(outputNoFormatPdf);
        using var pdf = new PdfDocument(writer);

        foreach (var file in supportedFiles)
        {
            if (!file.EndsWith(PDF_EXTENSION))
            {
                var imageData = ImageDataFactory.Create(file);

                float pixelWidth = imageData.GetWidth();
                float pixelHeight = imageData.GetHeight();

                float dpiX = imageData.GetDpiX() > 0 ? imageData.GetDpiX() : 72;
                float dpiY = imageData.GetDpiY() > 0 ? imageData.GetDpiY() : 72;

                float imageWidthPts = pixelWidth * 72f / dpiX;
                float imageHeightPts = pixelHeight * 72f / dpiY;

                float pageWidth = imageWidthPts;
                float pageHeight = imageHeightPts;

                if (autoPortrait)
                {
                    pageWidth = Math.Min(imageWidthPts, imageHeightPts);
                    pageHeight = Math.Max(imageWidthPts, imageHeightPts);
                }

                var page = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
                var canvas = new PdfCanvas(page);

                bool rotate = autoPortrait && imageWidthPts > imageHeightPts;

                if (!rotate)
                {
                    canvas.AddImageFittedIntoRectangle(
                        imageData,
                        new Rectangle(0, 0, pageWidth, pageHeight),
                        false);
                }
                else
                {
                    canvas.SaveState();

                    canvas.ConcatMatrix(0, 1, -1, 0, pageWidth, 0);

                    canvas.AddImageFittedIntoRectangle(
                        imageData,
                        new Rectangle(0, 0, pageHeight, pageWidth),
                        false);

                    canvas.RestoreState();
                }
            }
            else
            {
                using var src = new PdfDocument(new PdfReader(file));

                for (int i = 1; i <= src.GetNumberOfPages(); i++)
                {
                    var srcPage = src.GetPage(i);
                    var pageCopy = srcPage.CopyAsFormXObject(pdf);

                    Rectangle bbox = pageCopy.GetBBox().ToRectangle();

                    float w = bbox.GetWidth();
                    float h = bbox.GetHeight();

                    float pageWidth = w;
                    float pageHeight = h;

                    if (autoPortrait)
                    {
                        pageWidth = Math.Min(w, h);
                        pageHeight = Math.Max(w, h);
                    }

                    var newPage = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
                    var canvas = new PdfCanvas(newPage);

                    bool rotate = autoPortrait && w > h;

                    if (!rotate)
                    {
                        canvas.AddXObjectFittedIntoRectangle(
                            pageCopy,
                            new Rectangle(0, 0, pageWidth, pageHeight));
                    }
                    else
                    {
                        canvas.SaveState();

                        canvas.ConcatMatrix(0, 1, -1, 0, pageWidth, 0);

                        canvas.AddXObjectFittedIntoRectangle(
                            pageCopy,
                            new Rectangle(0, 0, pageHeight, pageWidth));

                        canvas.RestoreState();
                    }
                }
            }

        }
    }
    static bool IsPaperMatch(
        float imageW,
        float imageH,
        float paperW,
        float paperH,
        float percentThreshold)
    {
        float imgShort = Math.Min(imageW, imageH);
        float imgLong = Math.Max(imageW, imageH);

        float papShort = Math.Min(paperW, paperH);
        float papLong = Math.Max(paperW, paperH);

        float shortDiff = Math.Abs(imgShort - papShort) / papShort * 100f;
        float longDiff = Math.Abs(imgLong - papLong) / papLong * 100f;

        return shortDiff <= percentThreshold &&
               longDiff <= percentThreshold;
    }

    static float MmToPoints(float mm)
        => mm * 72f / 25.4f;

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