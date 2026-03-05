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
using iText.Kernel.Pdf.Xobject;
using System.Text;

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

    const float A2WMM = 420f;
    const float A2HMM = 594f;
    static readonly Dictionary<string, (float WidthMm, float HeightMm)> _PaperSizeDict =
    new(StringComparer.OrdinalIgnoreCase)
    {
        { "A5", (148,   210) },
        { "A4", (210,   297) },
        { "A3", (297,   420) },
        { "A2", (A2WMM, A2HMM) },
        { "A1", (594,   841) },
        { "A0", (841,   1189) },
        { "2A0",(1189,  1682) }
    };

    public readonly static float A2WPTS = MmToPts(A2WMM);
    public readonly static float A2HPTS = MmToPts(A2HMM);

    const float STANDARD_SIZE_TOLERANCE_MM = 2f;
    const string FONT_PATH = @"C:\Windows\Fonts\arial.ttf";

    public class PageInfoConllection
    {
        const int SIZE_GROUP_TOLERANCE_PTS = 5;

        public List<PageInfo>? PageInfoList { get; set; } = new();
        public List<List<PageInfo>> GenerateGroupBySize()
        {
            var groups = new List<List<PageInfo>>();
            foreach (var item in PageInfoList!)
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
                if (!added) groups.Add(new List<PageInfo> { item });
            }
            return groups;
        }
    }
    public class PageInfo
    {
        public float WidthPts { get; set; }
        public float HeightPts { get; set; }
        public int PageNumber { get; set; }
        public int FileIndex { get; set; }
        //If OrigPageNumber = 0 => image, otherwise pdf
        public bool IsImage => OrigPageNumber == 0;
        public int OrigPageNumber { get; set; }
        public float OrigWidthPts { get; set; }
        private float OrigWidthMm => PtsToMm(OrigWidthPts);
        public float OrigHeightPts { get; set; }
        private float OrigHeightMm => PtsToMm(OrigHeightPts);
        public TMatrix TransformMatrix { get; set; }
        public float llx { get; set; }
        public float lly { get; set; }
        public PageInfo(float OrigWidthPts, float OrigHeightPts, float WidthPts, float HeightPts, int PageNumber, int FileIndex, int OrigPageNumber, float llx, float lly, TMatrix TransformMatrix, PageInfoConllection group)
        {
            this.OrigWidthPts = OrigWidthPts;
            this.OrigHeightPts = OrigHeightPts;
            this.WidthPts = WidthPts;
            this.HeightPts = HeightPts;
            this.PageNumber = PageNumber;
            this.FileIndex = FileIndex;
            //If OrigPageNumber = 0 => image, otherwise pdf
            this.OrigPageNumber = OrigPageNumber;
            this.TransformMatrix = TransformMatrix;
            this.llx = llx;
            this.lly = lly;
            group.PageInfoList!.Add(this);
        }
    }

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

    public class Configuration
    {
        public string Folder { get; set; } = "";
        public string OutputPdf { get; set; } = "";
        public float MarginMm { get; set; } = 8f;
        public float NumberOffsetMm { get; set; } = 4f;
        public bool Stretch { get; set; } = true;
        public float AutoSizeToleranceMm { get; set; } = 5f;
        public string? StandardPageSize { get; set; } = null;
        public float? FitPageWidthMm { get; set; } = 210f;
        public bool CreateGroups { get; set; } = false;
        public bool CreateSingle { get; set; } = false;
        public bool Recursive { get; set; } = true;
        public bool AddPageNumber { get; set; } = true;

        public bool CreateReport { get; set; } = false;
        public List<string>? SupportedFiles { get { return supportedFiles; } }
        public List<string>? supportedFiles = null;


        // ---------- Derived values ----------

        public bool UsePageWidth => FitPageWidthMm.HasValue;

        public bool IsOneToOne => StandardPageSize?.Equals("1_1", StringComparison.OrdinalIgnoreCase) == true;

        public bool IsAuto => StandardPageSize?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;

        public bool AllowStretch => Stretch && !IsOneToOne;

        public float MarginPts => Program.MmToPts(MarginMm);

        public float NumberOffsetPts => Program.MmToPts(NumberOffsetMm);

        public float FitPageWidthPts => Program.MmToPts(FitPageWidthMm ?? 210);

        public float FontSize { get; set; }
        public string? FontPath { get; set; }

        public Configuration(
            string folder,
            string outputPdf,
            bool creategroups,
            bool createsingle,
            bool createreport,
            bool recursive,
            bool addpagenumber,
            float marginMm,
            float numberOffsetMm,
            bool stretch,
            float autoSizeToleranceMm,
            string? standardPageSize,
            float? fitPageWidthMm,
            string fontPath = FONT_PATH,
            float fontSize = 5f)
        {
            Folder = folder.Trim('"');
            OutputPdf = outputPdf;
            CreateGroups = creategroups;
            CreateSingle = createsingle;
            CreateReport = createreport;
            Recursive = recursive;
            AddPageNumber = addpagenumber;
            MarginMm = marginMm;
            NumberOffsetMm = numberOffsetMm;
            Stretch = stretch;
            AutoSizeToleranceMm = autoSizeToleranceMm;
            StandardPageSize = standardPageSize;
            FitPageWidthMm = fitPageWidthMm;
            FontPath = fontPath;
            FontSize = fontSize;

            if (!Recursive)
            {
                // NonRecursive
                supportedFiles = Directory.EnumerateFiles(folder)
                    .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
                    .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
                    .ToList();
            }
            else
            {
                // Recursive
                supportedFiles = Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                    .Where(f => _ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
                    .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
                    .ToList();
            }
        }


    }

    public class PageAnalyzer
    {
        public Configuration config;
        public List<List<PageInfo>>? PageInfosGroupBySize { get; set; }
        public PageInfoConllection? PageInfos { get; set; }
        public List<int> A2Pages = new List<int>();
        public PageAnalyzer(Configuration config)
        {
            this.config = config;
        }

        public void AnalyzePages()
        {
            PageInfos = new PageInfoConllection();
            int fileIndex = 0;
            int pageNumber = 0;
            float llx = 0f;
            float lly = 0f;
            foreach (var file in config.SupportedFiles!)
            {
                if (file.EndsWith(PDF_EXTENSION) == false)
                {
                    //Images
                    var imageData = ImageDataFactory.Create(file);

                    float pixelX = imageData.GetWidth();
                    float pixelY = imageData.GetHeight();

                    float dpiX = imageData.GetDpiX() > 0 ? imageData.GetDpiX() : 72;
                    float dpiY = imageData.GetDpiY() > 0 ? imageData.GetDpiY() : 72;

                    float origXPts = pixelX * 72f / dpiX;
                    float origYPts = pixelY * 72f / dpiY;

                    float pageWidthPts;
                    float pageHeightPts;
                    TMatrix transformMatrix;
                    CalculatePageSize(origXPts, origYPts, out pageWidthPts, out pageHeightPts, out transformMatrix, true, 0f, 0f);
                    pageNumber++;

                    if (MatchStandardPageSize(pageWidthPts, pageHeightPts, A2WPTS, A2HPTS, config.AutoSizeToleranceMm)) A2Pages.Add(pageNumber);
                    //If OrigPageNumber = 0 => image, otherwise pdf.  This case = 0
                    new PageInfo(origXPts, origYPts, pageWidthPts, pageHeightPts, pageNumber, fileIndex, 0, 0f, 0f, transformMatrix, PageInfos);
                }
                else
                {
                    //PDF
                    using var src = new PdfDocument(new PdfReader(file));
                    for (int i = 1; i <= src.GetNumberOfPages(); i++)
                    {
                        var srcPage = src.GetPage(i);

                        Rectangle bbox = srcPage.GetCropBox();

                        float origXPts = bbox.GetWidth();
                        float origYPts = bbox.GetHeight();
                        //Get location x,y of the cropped box
                        llx = bbox.GetX();
                        lly = bbox.GetY();

                        float pageWidthPts;
                        float pageHeightPts;
                        TMatrix transformMatrix;
                        CalculatePageSize(origXPts, origYPts, out pageWidthPts, out pageHeightPts, out transformMatrix, false, llx, lly);
                        pageNumber++;

                        if (MatchStandardPageSize(pageWidthPts, pageHeightPts, A2WPTS, A2HPTS, config.AutoSizeToleranceMm)) A2Pages.Add(pageNumber);
                        //If OrigPageNumber = 0 => image, otherwise pdf.  This case = original pdf page number
                        new PageInfo(origXPts, origYPts, pageWidthPts, pageHeightPts, pageNumber, fileIndex, i, llx, lly, transformMatrix, PageInfos);
                    }
                }
                fileIndex++;
            }

            PageInfosGroupBySize = PageInfos.GenerateGroupBySize();
        }

        public void CalculatePageSize(float origXPts, float origYPts, out float pageWidthPts, out float pageHeightPts, out TMatrix TransformMatrix, bool isImage, float llx, float lly)
        {
            //Note: origXPts, origYPts is X, Y lenghts
            //x,y placement location calculate from bottom left
            //w,h width and height of the placement area. w is horizontal and h is vertical.
            //llx, lly is the location of cropped box in the original pdf page

            //x, y, w, h is parameters for placement on new pdf page.  They are calculated
            //from user options

            //All are in pts
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

            float x, y, w, h;

            float imgShortPts = Math.Min(origXPts, origYPts);
            float imgLongPts = Math.Max(origXPts, origYPts);
            float lsRatio = imgLongPts / imgShortPts;

            if (config.UsePageWidth)
            {
                pageWidthPts = config.FitPageWidthPts;
                x = config.MarginPts;
                y = config.MarginPts;
                w = pageWidthPts - 2 * config.MarginPts;
                if (config.AllowStretch)
                {
                    //Page size is calculated from image size, when placing image with config.MarginPts
                    //will make it stretched
                    pageHeightPts = config.FitPageWidthPts * lsRatio;
                    h = pageHeightPts - 2 * config.MarginPts;
                }
                else
                {
                    //Not allow stretched:
                    //Have w => h from lsRatio
                    h = w * lsRatio;
                    //Then calculate page height
                    pageHeightPts = h + 2 * config.MarginPts;
                }
            }
            else if (config.IsOneToOne)
            {
                pageWidthPts = imgShortPts + (config.MarginPts * 2);
                pageHeightPts = imgLongPts + (config.MarginPts * 2);

                x = config.MarginPts;
                y = config.MarginPts;
                w = imgShortPts;
                h = imgLongPts;
            }
            else
            {
                // If use standard paper
                if (!string.IsNullOrEmpty(config.StandardPageSize) && !config.IsAuto)
                {
                    //Specify wrong standard paper size name
                    if (!_PaperSizeDict.ContainsKey(config.StandardPageSize))
                        throw new Exception($"Unsupported page size: {config.StandardPageSize}");
                    //Get user specified paper
                    var paper = _PaperSizeDict[config.StandardPageSize];

                    pageWidthPts = MmToPts(Math.Min(paper.WidthMm, paper.HeightMm));
                    pageHeightPts = MmToPts(Math.Max(paper.WidthMm, paper.HeightMm));
                }
                else
                {
                    //Not specified => pagesize is equal to original size
                    pageWidthPts = Math.Min(origXPts, origYPts);
                    pageHeightPts = Math.Max(origXPts, origYPts);

                    //If auto => detect standard size and recalculate paper size
                    if (config.IsAuto)
                    {
                        foreach (var paper in _PaperSizeDict)
                        {
                            float paperW = MmToPts(paper.Value.WidthMm);
                            float paperH = MmToPts(paper.Value.HeightMm);
                            // Paper size is auto detected
                            if (MatchStandardPageSize(pageWidthPts, pageHeightPts, paperW, paperH, config.AutoSizeToleranceMm))
                            {
                                pageWidthPts = Math.Min(paperW, paperH);
                                pageHeightPts = Math.Max(paperW, paperH);
                                break;
                            }
                        }
                    }
                }
                //Until here paper size are calculated correctly:
                //  - Specified has highest priority
                //  - Not specify -> use original size
                //  - Set with auto -> auto detect.
                //      If auto can't find => not change value -> fallback to not specified 

                // Stretch -> just fill the area with margin
                if (config.AllowStretch)
                {
                    x = config.MarginPts;
                    y = config.MarginPts;
                    w = pageWidthPts - 2 * config.MarginPts;
                    h = pageHeightPts - 2 * config.MarginPts;
                }
                //Calculate the fill area
                else
                {
                    //Maximum ration of h/w is the area inside page after page area reduce the margin
                    float maxw = pageWidthPts - 2 * config.MarginPts;
                    float maxh = pageHeightPts - 2 * config.MarginPts;
                    float maxhwRatio = maxh / maxw;
                    //This will fit to w
                    if (maxhwRatio > lsRatio)
                    {
                        x = config.MarginPts;
                        w = pageWidthPts - 2 * config.MarginPts;
                        h = w * lsRatio;
                        y = (pageHeightPts - h) / 2;
                    }
                    else //This will fit to h
                    {
                        y = config.MarginPts;
                        h = pageHeightPts - 2 * config.MarginPts;
                        w = h / lsRatio;
                        x = (pageWidthPts - w) / 2;
                    }
                }
            }
            //Until here all x, y, w, h are calculated.  imgXPts > imgYPts => landscape
            if (!isImage) TransformMatrix = CreatePdfTransformMatrix(origXPts, origYPts, x, y, w, h, llx, lly);
            else TransformMatrix = CreateImageTransformMatrix(x, y, w, h, origXPts > origYPts);
        }

        TMatrix CreatePdfTransformMatrix(float origXPts, float origYPts, float x, float y, float w, float h, float llx, float lly)
        {
            // For pdf, we get cropped and rotate area => this will have origin at llx, lly
            // And also have to scale original
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
        TMatrix CreateImageTransformMatrix(float x, float y, float w, float h, bool isImageLandscape)
        {
            //Ammazingly image don't need to scale 
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

        public void WriteConsoleGroupBySize()
        {
            StringBuilder text = new StringBuilder();
            GroupsToText(text);
            Console.WriteLine(text);
        }
        public void WriteConsoleNoGroup(bool detailFilePath)
        {
            StringBuilder text = new StringBuilder();
            NonGroupToText(text, detailFilePath);
            Console.WriteLine(text);
        }

        void GroupsToText(StringBuilder text)
        {
            if (text == null) return;
            text.Append($"TOTAL {PageInfos!.PageInfoList!.Count} PAGE(S):{Environment.NewLine}");
            foreach (var group in PageInfosGroupBySize!)
            {
                text.Append($"\t{GetSizeLabel(group[0].WidthPts, group[0].HeightPts)} - Total: {group.Count} page(s){Environment.NewLine}\t\t");
                foreach (var item in group) { text.Append(item.PageNumber).Append(" "); }
                text.Append(Environment.NewLine);
            }
        }
        void NonGroupToText(StringBuilder text, bool withFilePath)
        {
            if (text == null) return;
            text.Append($"PAGE INFO:{Environment.NewLine}");
            int index = 0;
            foreach (var origPageInfo in PageInfos!.PageInfoList!)
            {
                text.Append($"\t{origPageInfo.PageNumber} - {GetSizeLabel(origPageInfo.WidthPts, origPageInfo.HeightPts)}");
                if (withFilePath) text.Append($"\t{config.SupportedFiles![origPageInfo.FileIndex]}{Environment.NewLine}");
                else text.Append(Environment.NewLine);
                index++;
            }
        }
        public void WriteReportTextFile(bool detailFilePath)
        {
            StringBuilder text = new StringBuilder();
            GroupsToText(text);
            NonGroupToText(text, detailFilePath);
            File.WriteAllText(IOPath.ChangeExtension(config.OutputPdf, ".txt"), text.ToString());
        }

    }

    public class PdfCreator
    {
        PageAnalyzer PageAnalyzer { get; set; }
        public PdfCreator(PageAnalyzer pageAnalyzer)
        {
            PageAnalyzer = pageAnalyzer;
        }

        Configuration config => PageAnalyzer.config;
        public void CreateGroupBySizePdfFromFiles(bool addPageNumber = true)
        {
            foreach (var list in PageAnalyzer.PageInfosGroupBySize!)
            {
                using var writer = new PdfWriter(IOPath.ChangeExtension(config.OutputPdf, null) + "_" + GetSizeLabel(list[0].WidthPts, list[0].HeightPts) + PDF_EXTENSION);
                using var pdf = new PdfDocument(writer);

                foreach (var origPageInfo in list)
                {
                    PdfFont font = PdfFontFactory.CreateFont(
                        config.FontPath,
                        PdfEncodings.WINANSI,
                        PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED
                    );
                    if (origPageInfo.IsImage) AddPageFromImage(origPageInfo, pdf, font, config.AddPageNumber);
                    else AddPageFromPdf(origPageInfo, pdf, font, true);
                }
            }
        }

        public void CreateOnePdfFromFiles(bool addPageNumber = true)
        {
            using var writer = new PdfWriter(IOPath.ChangeExtension(config.OutputPdf, null) + "_ALL" + PDF_EXTENSION);
            using var pdf = new PdfDocument(writer);
            foreach (var origPageInfo in PageAnalyzer.PageInfos!.PageInfoList!)
            {
                PdfFont font = PdfFontFactory.CreateFont(
                    config.FontPath,
                    PdfEncodings.WINANSI,
                    PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);

                if (origPageInfo.IsImage) AddPageFromImage(origPageInfo, pdf, font, config.AddPageNumber);
                else AddPageFromPdf(origPageInfo, pdf, font, true);
            }
        }


        void AddPageFromImage(PageInfo origPageInfo, PdfDocument pdf, PdfFont font, bool addPageNumber = true)
        {
            var imageData = ImageDataFactory.Create(config.SupportedFiles![origPageInfo.FileIndex]);
            var page = pdf.AddNewPage(new PageSize(origPageInfo.WidthPts, origPageInfo.HeightPts));
            var canvas = new PdfCanvas(page);
            origPageInfo.TransformMatrix.ApplyMatrix(canvas, imageData);
            if (addPageNumber) AddPageNumber(canvas, font, config.NumberOffsetPts, origPageInfo.PageNumber);
        }

        void AddPageFromPdf(PageInfo origPageInfo, PdfDocument pdf, PdfFont font, bool addPageNumber = true)
        {
            using var src = new PdfDocument(new PdfReader(config.SupportedFiles![origPageInfo.FileIndex]));
            var srcPage = src.GetPage(origPageInfo.OrigPageNumber);

            int rot = srcPage.GetRotation();
            Rectangle crop = srcPage.GetCropBox();
            srcPage.SetRotation(0);
            var pageCopy = srcPage.CopyAsFormXObject(pdf);
            var newPage = pdf.AddNewPage(new PageSize(origPageInfo.WidthPts, origPageInfo.HeightPts));

            var canvas = new PdfCanvas(newPage);

            origPageInfo.TransformMatrix.ApplyMatrix(canvas, pageCopy);

            if (addPageNumber) AddPageNumber(canvas, font, config.NumberOffsetPts, origPageInfo.PageNumber);
        }
        void AddPageNumber(PdfCanvas canvas, PdfFont font, float numberOffset, int pageNumber)
        {
            canvas.BeginText();
            canvas.SetFontAndSize(font, config.FontSize);
            canvas.SetFillColor(ColorConstants.GRAY);
            canvas.MoveText(numberOffset, numberOffset);
            canvas.ShowText(pageNumber.ToString());
            canvas.EndText();
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

        Configuration config = new Configuration(
            options["input"],
            options["output"],
            options.ContainsKey("creategroups") && options["creategroups"].ToLower() == "y",
            options.ContainsKey("createsingle") && options["createsingle"].ToLower() == "y",
            options.ContainsKey("createreport") && options["createreport"].ToLower() == "y",
            options.ContainsKey("recursive") && options["recursive"].ToLower() == "y",
            options.ContainsKey("addpagenumber") && options["addpagenumber"].ToLower() == "y",
            options.ContainsKey("margin") ? float.Parse(options["margin"]) : 8f,
            options.ContainsKey("numberoffset") ? float.Parse(options["numberoffset"]) : 4f,
            options.ContainsKey("stretch") && options["stretch"].ToLower() == "y",
            options.ContainsKey("autosizetolerance") ? float.Parse(options["autosizetolerance"]) : 5f,
            options.ContainsKey("pagesize") ? options["pagesize"] : null,
            options.ContainsKey("pagewidth") ? (float?)float.Parse(options["pagewidth"]) : null,
            FONT_PATH,
            5f);

        //if (!config.CreateGroups && !config.CreateSingle) return 0;

        PageAnalyzer pa = new PageAnalyzer(config);

        pa.AnalyzePages();

        //foreach(var i in pa.A2Pages) Console.WriteLine(i);



        PdfCreator pdfCreator = new PdfCreator(pa);
        if (config.CreateGroups) pdfCreator.CreateGroupBySizePdfFromFiles(config.AddPageNumber);
        if (config.CreateSingle) pdfCreator.CreateOnePdfFromFiles(config.AddPageNumber);

        if (config.CreateReport) pa.WriteReportTextFile(false);

        return 0;
    }



    static string GetSizeLabel(float widthPts, float heightPts)
    {
        float wMm = PtsToMm(widthPts);
        float hMm = PtsToMm(heightPts);

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
                //sizeLabel = $"{kv.Key} ({paperShort}x{paperLong} mm)";
                sizeLabel = $"{kv.Key}";
                break;
            }
            else sizeLabel = $"{Math.Round(shortSide)}x{Math.Round(longSide)} mm";
        }
        return sizeLabel;
    }
    static bool MatchStandardPageSize(float origWPts, float origHPts, float paperWPts, float paperHPts, float percentThreshold)
    {
        float imgShort = Math.Min(origWPts, origHPts);
        float imgLong = Math.Max(origWPts, origHPts);

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










