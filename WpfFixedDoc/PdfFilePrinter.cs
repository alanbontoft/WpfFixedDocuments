using System;
using System.Linq;
using System.Printing;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Documents;
using System.IO.Packaging;
using System.Windows.Xps.Packaging;

public static class PdfFilePrinter
{
    private const string PdfPrinterDriveName = "Microsoft Print To PDF";

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    private class DOCINFOA
    {
        [MarshalAs(UnmanagedType.LPStr)]
        public string pDocName;
        [MarshalAs(UnmanagedType.LPStr)]
        public string pOutputFile;
        [MarshalAs(UnmanagedType.LPStr)]
        public string pDataType;
    }

    [DllImport("winspool.drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

    [DllImport("winspool.drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern int StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

    [DllImport("winspool.drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

    
    /// <summary>
    /// Print an XPS document to a 
    /// </summary>
    /// <param name="bytes"></param>
    /// <param name="outputFilePath"></param>
    /// <param name="documentTitle"></param>
    /// <exception cref="Exception"></exception>
	public static void PrintXpsToPdf(FixedDocument doc, string outputFilePath, string documentTitle)
    {
        var bytes = FixedDocumentToBytes(doc);

        // Get Microsoft Print to PDF print queue
        var pdfPrintQueue = GetMicrosoftPdfPrintQueue();

        // Copy byte array to unmanaged pointer
        var ptrUnmanagedBytes = Marshal.AllocCoTaskMem(bytes.Length);
        Marshal.Copy(bytes, 0, ptrUnmanagedBytes, bytes.Length);

        // Prepare document info
        var di = new DOCINFOA
        {
            pDocName = documentTitle,
            pOutputFile = outputFilePath,
            pDataType = "RAW"
        };

        // Print to PDF
        var errorCode = SendBytesToPrinter(pdfPrintQueue.Name, ptrUnmanagedBytes, bytes.Length, di, out var jobId);

        // Free unmanaged memory
        Marshal.FreeCoTaskMem(ptrUnmanagedBytes);

        // Check if job in error state (for example not enough disk space)
        var jobFailed = false;
        try
        {
            var pdfPrintJob = pdfPrintQueue.GetJob(jobId);
            if (pdfPrintJob.IsInError)
            {
                jobFailed = true;
                pdfPrintJob.Cancel();
            }
        }
        catch
        {
            // If job succeeds, GetJob will throw an exception. Ignore it. 
        }
        finally
        {
            pdfPrintQueue.Dispose();
        }

        if (errorCode > 0 || jobFailed)
        {
            try
            {
                if (File.Exists(outputFilePath))
                {
                    File.Delete(outputFilePath);
                }
            }
            catch
            {
                // ignored
            }
        }

        if (errorCode > 0)
        {
            throw new Exception($"Printing to PDF failed. Error code: {errorCode}.");
        }

        if (jobFailed)
        {
            throw new Exception("PDF Print job failed.");
        }
    }

    private static int SendBytesToPrinter(string szPrinterName, IntPtr pBytes, int dwCount, DOCINFOA documentInfo, out int jobId)
    {
        jobId = 0;
        var dwWritten = 0;
        var success = false;

        if (OpenPrinter(szPrinterName.Normalize(), out var hPrinter, IntPtr.Zero))
        {
            jobId = StartDocPrinter(hPrinter, 1, documentInfo);
            if (jobId > 0)
            {
                if (StartPagePrinter(hPrinter))
                {
                    success = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                    EndPagePrinter(hPrinter);
                }

                EndDocPrinter(hPrinter);
            }

            ClosePrinter(hPrinter);
        }

        // TODO: The other methods such as OpenPrinter also have return values. Check those?

        if (success == false)
        {
            return Marshal.GetLastWin32Error();
        }

        return 0;
    }

    private static PrintQueue GetMicrosoftPdfPrintQueue()
    {
        PrintQueue pdfPrintQueue = null;

        try
        {
            using (var printServer = new PrintServer())
            {
                var flags = new[] { EnumeratedPrintQueueTypes.Local };
                // FirstOrDefault because it's possible for there to be multiple PDF printers with the same driver name (though unusual)
                // To get a specific printer, search by FullName property instead (note that in Windows, queue name can be changed)
                pdfPrintQueue = printServer.GetPrintQueues(flags).FirstOrDefault(lq => lq.QueueDriver.Name == PdfPrinterDriveName);
            }

            if (pdfPrintQueue == null)
            {
                throw new Exception($"Could not find printer with driver name: {PdfPrinterDriveName}");
            }

            if (!pdfPrintQueue.IsXpsDevice)
            {
                throw new Exception($"PrintQueue '{pdfPrintQueue.Name}' does not understand XPS page description language.");
            }

            return pdfPrintQueue;
        }
        catch
        {
            pdfPrintQueue?.Dispose();
            throw;
        }
    }

    private static byte[] FixedDocumentToBytes(FixedDocument fixeddoc)
    {
        // Convert FixedDocument to XPS file in memory
        var ms = new MemoryStream();
        var package = Package.Open(ms, FileMode.Create);
        var xpsdoc = new XpsDocument(package);
        var writer = XpsDocument.CreateXpsDocumentWriter(xpsdoc);
        writer.Write(fixeddoc.DocumentPaginator);
        xpsdoc.Close();
        package.Close();

        // Get XPS file bytes
        var bytes = ms.ToArray();
        ms.Dispose();

        return bytes;
    }
}
