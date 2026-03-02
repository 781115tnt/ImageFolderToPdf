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
    static readonly string[] ImageExtensions =
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", PDF_EXTENSION
    };

    static readonly Dictionary<string, (float WidthMm, float HeightMm)> PaperSizes =
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

    static string folder = "";
    static string outputPdf = "";
    static float marginMm = 8f;
    static float numberOffsetMm = 4f;
    static bool stretch = true;
    static float percentThreshold = 5f;
    static string pageSizeOption = "A4";
    static float? pageWidthMm = 210f;
    static string fontPath = @"C:\Windows\Fonts\arial.ttf";
    static float fontSize = 5f;

    static bool usePageWidth => pageWidthMm.HasValue;
    static bool isOneToOne => pageSizeOption?.Equals("1_1", StringComparison.OrdinalIgnoreCase) == true;
    static bool isAuto => pageSizeOption?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;


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

        folder = options["input"];
        outputPdf = options["output"];

        marginMm = options.ContainsKey("margin") ? float.Parse(options["margin"]) : 8f;
        numberOffsetMm = options.ContainsKey("numberoffset") ? float.Parse(options["numberoffset"]) : 4f;
        stretch = options.ContainsKey("stretch") && options["stretch"].ToLower() == "y";
        percentThreshold = options.ContainsKey("percentthreshold") ? float.Parse(options["percentthreshold"]) : 5f;
        pageSizeOption = options.ContainsKey("pagesize") ? options["pagesize"] : "A4";
        pageWidthMm = options.ContainsKey("pagewidth") ? float.Parse(options["pagewidth"]) : null;

        fontPath = @"C:\Windows\Fonts\arial.ttf";
        fontSize = 5f;

        CreatePdfFromFiles(
            folder,
            outputPdf,
            marginMm,
            fontPath,
            fontSize,
            numberOffsetMm,
            stretch,
            percentThreshold,
            pageSizeOption,
            pageWidthMm,
            true
        );

        CreatePdfFromFilesNoFormat(folder, outputPdf + "_NO_" + ".pdf", false);

        Console.WriteLine("Done.");
        return 0;
    }

    static void CalculatePageSize(ref float imageWidthPts, ref float imageHeightPts, out float pageWidth, out float pageHeight)
    {
        float margin = MmToPoints(marginMm);
        float numberOffset = MmToPoints(numberOffsetMm);

        if (usePageWidth)
        {
            float targetWidthPts = MmToPoints(pageWidthMm ?? 210f);

            float imgShort = Math.Min(imageWidthPts, imageHeightPts);
            float imgLong = Math.Max(imageWidthPts, imageHeightPts);
            float ratio = imgLong / imgShort;

            pageWidth = targetWidthPts;
            pageHeight = targetWidthPts * ratio;
        }
        else if (isOneToOne)
        {
            float imgShort = Math.Min(imageWidthPts, imageHeightPts);
            float imgLong = Math.Max(imageWidthPts, imageHeightPts);

            pageWidth = imgShort + (margin * 2);
            pageHeight = imgLong + (margin * 2);
        }
        else if (!string.IsNullOrEmpty(pageSizeOption) && !isAuto)
        {
            if (!PaperSizes.ContainsKey(pageSizeOption))
                throw new Exception($"Unsupported page size: {pageSizeOption}");

            var paper = PaperSizes[pageSizeOption];

            pageWidth = MmToPoints(Math.Min(paper.WidthMm, paper.HeightMm));
            pageHeight = MmToPoints(Math.Max(paper.WidthMm, paper.HeightMm));
        }
        else
        {
            pageWidth = Math.Min(imageWidthPts, imageHeightPts);
            pageHeight = Math.Max(imageWidthPts, imageHeightPts);

            if (isAuto)
            {
                foreach (var paper in PaperSizes)
                {
                    float paperW = MmToPoints(paper.Value.WidthMm);
                    float paperH = MmToPoints(paper.Value.HeightMm);

                    if (IsPaperMatch(pageWidth, pageHeight, paperW, paperH, percentThreshold))
                    {
                        pageWidth = Math.Min(paperW, paperH);
                        pageHeight = Math.Max(paperW, paperH);
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

            float margin = MmToPoints(marginMm);
            float numberOffset = MmToPoints(numberOffsetMm);

            float pageWidth;
            float pageHeight;

            CalculatePageSize(ref imageWidthPts, ref imageHeightPts, out pageWidth, out pageHeight);

            var newPage = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
            pageNumber++;

            var canvas = new PdfCanvas(newPage);

            float availableWidth = pageWidth - (margin * 2);
            float availableHeight = pageHeight - (margin * 2);

            bool allowStretch = stretch;
            if (isOneToOne)
                allowStretch = false;

            float drawWidth;
            float drawHeight;

            if (!allowStretch)
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

            // Page number
            if (addPageNumber)
            {
                canvas.BeginText();
                canvas.SetFontAndSize(font, fontSize);
                canvas.SetFillColor(ColorConstants.GRAY);
                canvas.MoveText(numberOffset, numberOffset);
                canvas.ShowText(pageNumber.ToString());
                canvas.EndText();
            }
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

        float margin = MmToPoints(marginMm);
        float numberOffset = MmToPoints(numberOffsetMm);

        float pageWidth;
        float pageHeight;
        CalculatePageSize(ref imageWidthPts, ref imageHeightPts, out pageWidth, out pageHeight);


        var page = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
        pageNumber++;

        var canvas = new PdfCanvas(page);

        bool isLandscape = imageWidthPts > imageHeightPts;

        float availableWidth = pageWidth - (margin * 2);
        float availableHeight = pageHeight - (margin * 2);

        bool allowStretch = stretch;
        if (isOneToOne)
            allowStretch = false;

        float drawWidth;
        float drawHeight;

        if (!isLandscape)
        {
            if (!allowStretch)
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

            if (!allowStretch)
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
        if (addPageNumber)
        {
            canvas.BeginText();
            canvas.SetFontAndSize(font, fontSize);
            canvas.SetFillColor(ColorConstants.GRAY);
            canvas.MoveText(numberOffset, numberOffset);
            canvas.ShowText(pageNumber.ToString());
            canvas.EndText();
        }
    }



    static void CreatePdfFromFiles(
         string folder,
         string outputPdf,
         float marginMm,
         string fontPath,
         float fontSize,
         float numberOffsetMm,
         bool stretch,
         float percentThreshold,
         string pageSizeOption,
         float? pageWidth,
         bool addPageNumber = true)
    {
        var supportedFiles = Directory
            .EnumerateFiles(folder)
            .Where(f => ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
            .ToList();

        using var writer = new PdfWriter(outputPdf);
        using var pdf = new PdfDocument(writer);

        PdfFont font = PdfFontFactory.CreateFont(
            fontPath,
            PdfEncodings.WINANSI,
            PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED
        );

        int pageNumber = 0;

        foreach (var file in supportedFiles)
        {
            if (file.EndsWith(PDF_EXTENSION) == false) AddPageFromImage(file, ref pageNumber, writer, pdf, font, addPageNumber);
            else AddPagesFromPdf(file, ref pageNumber, writer, pdf, font, addPageNumber);
        }
    }

    static void CreatePdfFromFilesNoFormat(
        string folder,
        string outputPdf,
        bool autoPortrait = false)
    {
        var supportedFiles = Directory
            .EnumerateFiles(folder)
            .Where(f => ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => IOPath.GetFileName(f), new NaturalSortComparer())
            .ToList();

        using var writer = new PdfWriter(outputPdf);
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