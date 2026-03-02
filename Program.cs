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

class Program
{
    static readonly string[] ImageExtensions =
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"
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

    static int Main(string[] args)
    {
        var options = ParseArguments(args);

        if (!options.ContainsKey("command") || options["command"] != "img2pdf")
        {
            Console.WriteLine("Usage: -command=img2pdf -input=folder -output=file.pdf");
            return 1;
        }

        if (!options.ContainsKey("input") || !options.ContainsKey("output"))
        {
            Console.WriteLine("Missing -input or -output");
            return 1;
        }

        string folder = options["input"];
        string outputPdf = options["output"];

        float marginMm = options.ContainsKey("margin") ? float.Parse(options["margin"]) : 8f;
        float numberOffsetMm = options.ContainsKey("numberoffset") ? float.Parse(options["numberoffset"]) : 4f;
        bool stretch = options.ContainsKey("stretch") && options["stretch"].ToLower() == "y";
        float percentThreshold = options.ContainsKey("percentthreshold") ? float.Parse(options["percentthreshold"]) : 5f;
        string pageSizeOption = options.ContainsKey("pagesize") ? options["pagesize"] : null;
        float? fitWidth = options.ContainsKey("fitwidth") ? float.Parse(options["fitwidth"]) : null;

        string fontPath = @"C:\Windows\Fonts\arial.ttf";
        float fontSize = 5f;

        CreatePdfFromImages(
            folder,
            outputPdf,
            marginMm,
            fontPath,
            fontSize,
            numberOffsetMm,
            stretch,
            percentThreshold,
            pageSizeOption,
            fitWidth
        );

        Console.WriteLine("Done.");
        return 0;
    }

    static void CreatePdfFromImages(
        string folder,
        string outputPdf,
        float marginMm,
        string fontPath,
        float fontSize,
        float numberOffsetMm,
        bool stretchImage,
        float percentThreshold,
        string pageSizeOption,
        float? fitWidth)
    {
        var images = Directory
            .EnumerateFiles(folder)
            .Where(f => ImageExtensions.Contains(IOPath.GetExtension(f).ToLower()))
            .OrderBy(f => f)
            .ToList();

        using var writer = new PdfWriter(outputPdf);
        using var pdf = new PdfDocument(writer);

        PdfFont font = PdfFontFactory.CreateFont(
            fontPath,
            PdfEncodings.WINANSI,
            PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED
        );

        int pageNumber = 0;

        foreach (var file in images)
        {
            var imageData = ImageDataFactory.Create(file);

            float pixelWidth = imageData.GetWidth();
            float pixelHeight = imageData.GetHeight();

            float dpiX = imageData.GetDpiX() > 0 ? imageData.GetDpiX() : 72;
            float dpiY = imageData.GetDpiY() > 0 ? imageData.GetDpiY() : 72;

            float imageWidthPts = pixelWidth * 72f / dpiX;
            float imageHeightPts = pixelHeight * 72f / dpiY;

            float pageWidth;
            float pageHeight;

            float margin = MmToPoints(marginMm);
            float numberOffset = MmToPoints(numberOffsetMm);

            bool isOneToOne = pageSizeOption?.Equals("1_1", StringComparison.OrdinalIgnoreCase) == true;
            bool isAuto = pageSizeOption?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;
            bool useFitWidth = fitWidth.HasValue;

            // -------- PAGE SIZE LOGIC --------

            if (useFitWidth)
            {
                float targetWidthPts = MmToPoints(fitWidth.Value);
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

            var page = pdf.AddNewPage(new PageSize(pageWidth, pageHeight));
            pageNumber++;

            var canvas = new PdfCanvas(page);

            bool isLandscape = imageWidthPts > imageHeightPts;

            float availableWidth = pageWidth - (margin * 2);
            float availableHeight = pageHeight - (margin * 2);

            bool allowStretch = stretchImage;
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

            canvas.BeginText();
            canvas.SetFontAndSize(font, fontSize);
            canvas.SetFillColor(ColorConstants.GRAY);
            canvas.MoveText(numberOffset, numberOffset);
            canvas.ShowText(pageNumber.ToString());
            canvas.EndText();
        }
    }

    static bool IsPaperMatch(
        float imageW,
        float imageH,
        float paperW,
        float paperH,
        float percentThreshold)
    {
        float imageShort = Math.Min(imageW, imageH);
        float imageLong = Math.Max(imageW, imageH);

        float paperShort = Math.Min(paperW, paperH);
        float paperLong = Math.Max(paperW, paperH);

        float diffShort = Math.Abs(imageShort - paperShort) / paperShort * 100f;
        float diffLong = Math.Abs(imageLong - paperLong) / paperLong * 100f;

        return diffShort <= percentThreshold &&
               diffLong <= percentThreshold;
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