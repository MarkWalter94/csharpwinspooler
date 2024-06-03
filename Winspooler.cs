using System.Text;
using System.Runtime.InteropServices;
using System.Security;
using System.ComponentModel;
using Microsoft.Extensions.Logging;

// ReSharper disable UnusedMember.Local

/// <summary>
/// Origin: http://blog.csdn.net/csui2008/article/details/5718461
/// Modified and a little tested by Huan-Lin Tsai. May-29-2013.
/// 
/// Re-modified by Marco Ortali on Jun-03-2024.
///     - Clean up
///     - Added delete pj methods
///     - DInjection
///     - Translated to English the cinese comments
/// </summary>
public interface IWinspooler
{
    /// <summary>
    /// Get the settings information of the specified printer。
    /// </summary>
    /// <param name="printerName">Printer name</param>
    /// <returns>Specify printer setup information</returns>
    Winspooler.DEVMODE GetPrinterDevMode(string? printerName);

    /// <summary>
    /// Determine whether the specific paper currently defaulted to the printer is equal to the incoming size。
    /// </summary>
    /// <param name="FormName">Paper name。</param>
    /// <param name="width">Width。Unit: 1/10 of a millimeter.</param>
    /// <param name="length">Height。Unit: 1/10 of a millimeter.</param>
    /// <returns>Returns true if the paper size in the default printer's DEVMODE structure is the same as the specified width and height, otherwise returns false.</returns>
    bool IsPaperSize(string FormName, int width, int length);

    /// <summary>
    /// Change printer settings。
    /// </summary>
    /// <param name="printerName">The name of the printer. If it is empty, the name of the default printer will be automatically obtained.</param>
    /// <param name="prnSettings">Information to change</param>
    /// <returns>Whether the change was successful</returns>
    void ModifyPrinterSettings(string printerName, ref Winspooler.PrinterSettingsInfo prnSettings);

    /// <summary>
    /// Another version of changing printer settings.
    /// During testing, the application terminated abnormally without any error message.
    /// Please use ModifyPrinterSettings.
    /// </summary>
    /// <param name="printerName">Printer name. Pass null or an empty string to use the default printer.</param>
    /// <param name="printerSetting">Info to change</param>
    /// <returns>Whether the change was successful</returns>
    bool ModifyPrinterSettings_V2(string printerName, ref Winspooler.PrinterSettingsInfo printerSetting);

    /// <summary>
    /// Get the name of the default printer
    /// </summary>
    /// <returns>Returns the name of the default printer</returns>
    string GetDefaultPrinterName();

    /// <summary>
    /// Get the kind of paper. If it is 0, it is an error.
    /// </summary>
    /// <param name="printerName">Printer name. Pass null or an empty string to use the default printer.</param>
    /// <param name="paperName">Paper name, must be filled in</param>
    /// <returns>Kind</returns>
    short GetOnePaper(string printerName, string paperName);

    /// <summary>
    /// Get all available papers and output paper specifications and names to the console.
    /// </summary>
    /// <param name="printerName">Printer name. Pass null or an empty string to use the default printer.</param>
    void ShowPapers(string printerName);

    /// <summary>
    /// Abort all the jobs for the printer.
    /// </summary>
    /// <param name="printerName"></param>
    void AbortPrinter(string printerName);
    
    /// <summary>
    /// Delete all the print jobs for a printer.
    /// </summary>
    /// <param name="printerName"></param>
    void DeleteAllJobs(string printerName);
}


public class Winspooler : IWinspooler
{
    private readonly ILogger<Winspooler> _logger;

    public Winspooler(ILogger<Winspooler> logger)
    {
        _logger = logger;
    }

    #region "Private Variables"
    private int nRet;
    private int intError;
    #endregion

    public const Int32 JOB_CONTROL_PAUSE = 0x1;
    public const Int32 JOB_CONTROL_RESUME = 0x2;
    public const Int32 JOB_CONTROL_RESTART = 0x4;
    public const Int32 JOB_CONTROL_CANCEL = 0x3;
    public const Int32 JOB_CONTROL_DELETE = 0x5;
    public const Int32 JOB_CONTROL_RETAIN = 0x8;
    public const Int32 JOB_CONTROL_RELEASE = 0x9;

    #region "API Define"
    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool SetDefaultPrinter(string printerName);


    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern int DocumentProperties(
        IntPtr hwnd,
        IntPtr hPrinter,
        [MarshalAs(UnmanagedType.LPStr)] string pDeviceName,
        IntPtr pDevModeOutput,
        IntPtr pDevModeInput,
        int fMode
    );

    [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true,
        CharSet = CharSet.Ansi, ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
    private static extern bool GetPrinter(IntPtr hPrinter, Int32 dwLevel,
        IntPtr pPrinter, Int32 dwBuf, out Int32 dwNeeded);

    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA",
        SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool
        OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter,
            out IntPtr hPrinter, IntPtr pDefault); //ref PRINTER_DEFAULTS pd);

    [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr
        pPrinter, int Command);

    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

    [DllImport("GDI32.dll", EntryPoint = "CreateDC", SetLastError = true,
         CharSet = CharSet.Unicode, ExactSpelling = false,
         CallingConvention = CallingConvention.StdCall),
     SuppressUnmanagedCodeSecurity()]
    private static extern IntPtr CreateDC([MarshalAs(UnmanagedType.LPTStr)] string pDrive,
        [MarshalAs(UnmanagedType.LPTStr)] string pName,
        [MarshalAs(UnmanagedType.LPTStr)] string pOutput,
        ref DEVMODE pDevMode);

    [DllImport("GDI32.dll", EntryPoint = "ResetDC", SetLastError = true,
         CharSet = CharSet.Unicode, ExactSpelling = false,
         CallingConvention = CallingConvention.StdCall),
     SuppressUnmanagedCodeSecurity()]
    private static extern IntPtr ResetDC(
        IntPtr hDC,
        ref DEVMODE
            pDevMode);

    [DllImport("GDI32.dll", EntryPoint = "DeleteDC", SetLastError = true,
         CharSet = CharSet.Unicode, ExactSpelling = false,
         CallingConvention = CallingConvention.StdCall),
     SuppressUnmanagedCodeSecurity()]
    private static extern bool DeleteDC(IntPtr hDC);

    [DllImport("winspool.drv", EntryPoint = "DeviceCapabilitiesA", SetLastError = true)]
    private static extern Int32 DeviceCapabilities(
        [MarshalAs(UnmanagedType.LPStr)] String device,
        [MarshalAs(UnmanagedType.LPStr)] string? port,
        Int16 capability,
        IntPtr outputBuffer,
        IntPtr deviceMode);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EnumPrintersW(Int32 flags,
        [MarshalAs(UnmanagedType.LPTStr)] string printerName,
        Int32 level, IntPtr buffer, Int32 bufferSize, out Int32
            requiredBufferSize, out Int32 numPrintersReturned);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool AbortPrinter(IntPtr hPrinter);

    [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false,
         ExactSpelling = true, CallingConvention = CallingConvention.StdCall),
     SuppressUnmanagedCodeSecurity()]
    private static extern Int32 GetLastError();

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr windowHandle,
        uint Msg,
        IntPtr wParam,
        IntPtr lParam,
        SendMessageTimeoutFlags flags,
        uint timeout,
        out IntPtr result
    );

    [DllImport("Winspool.drv", SetLastError = true, EntryPoint = "EnumJobsA")]
    private static extern bool EnumJobs(
        IntPtr hPrinter, // handle to printer object
        UInt32 FirstJob, // index of first job
        UInt32 NoJobs, // number of jobs to enumerate
        UInt32 Level, // information level
        IntPtr pJob, // job information buffer
        UInt32 cbBuf, // size of job information buffer
        out UInt32 pcbNeeded, // bytes received or required
        out UInt32 pcReturned // number of jobs received
    );

    [DllImport("winspool.drv", EntryPoint = "SetJobA")]
    private static extern bool SetJob(IntPtr hPrinter, int JobId, int Level, IntPtr pJob, int Command_Renamed);

    [StructLayout(LayoutKind.Sequential)]
    public struct JOB_INFO_1W
    {
        public uint JobId;
        [MarshalAs(UnmanagedType.LPWStr)] public string pPrinterName;
        [MarshalAs(UnmanagedType.LPWStr)] public string pMachineName;
        [MarshalAs(UnmanagedType.LPWStr)] public string pUserName;
        [MarshalAs(UnmanagedType.LPWStr)] public string pDocument;
        [MarshalAs(UnmanagedType.LPWStr)] public string pDatatype;
        [MarshalAs(UnmanagedType.LPWStr)] public string pStatus;
        public uint Status;
        public uint Priority;
        public uint Position;
        public uint TotalPages;
        public uint PagesPrinted;
        public SYSTEMTIME Submitted;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SYSTEMTIME
    {
        public ushort wYear;
        public ushort wMonth;
        public ushort wDayOfWeek;
        public ushort wDay;
        public ushort wHour;
        public ushort wMinute;
        public ushort wSecond;
        public ushort wMilliseconds;
    }
    #endregion

    #region "Data structure"
    /// <summary>
    /// Paper access rights and other information
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct PRINTER_DEFAULTS
    {
        public int pDatatype;
        public int pDevMode;
        public int DesiredAccess; //Access to printer
    }


    //Paper orientation
    public enum PageOrientation
    {
        DMORIENT_PORTRAIT = 1, //Straight
        DMORIENT_LANDSCAPE = 2, //Horizontal
    }

    /// <summary>
    /// Paper type
    /// </summary>
    public enum PaperSize
    {
        DMPAPER_LETTER = 1, // Letter 8 1/2 x 11 in
        DMPAPER_LETTERSMALL = 2, // Letter Small 8 1/2 x 11 in
        DMPAPER_TABLOID = 3, // Tabloid 11 x 17 in
        DMPAPER_LEDGER = 4, // Ledger 17 x 11 in
        DMPAPER_LEGAL = 5, // Legal 8 1/2 x 14 in
        DMPAPER_STATEMENT = 6, // Statement 5 1/2 x 8 1/2 in
        DMPAPER_EXECUTIVE = 7, // Executive 7 1/4 x 10 1/2 in
        DMPAPER_A3 = 8, // A3 297 x 420 mm
        DMPAPER_A4 = 9, // A4 210 x 297 mm
        DMPAPER_A4SMALL = 10, // A4 Small 210 x 297 mm
        DMPAPER_A5 = 11, // A5 148 x 210 mm
        DMPAPER_B4 = 12, // B4 250 x 354
        DMPAPER_B5 = 13, // B5 182 x 257 mm
        DMPAPER_FOLIO = 14, // Folio 8 1/2 x 13 in
        DMPAPER_QUARTO = 15, // Quarto 215 x 275 mm
        DMPAPER_10X14 = 16, // 10x14 in
        DMPAPER_11X17 = 17, // 11x17 in
        DMPAPER_NOTE = 18, // Note 8 1/2 x 11 in
        DMPAPER_ENV_9 = 19, // Envelope #9 3 7/8 x 8 7/8
        DMPAPER_ENV_10 = 20, // Envelope #10 4 1/8 x 9 1/2
        DMPAPER_ENV_11 = 21, // Envelope #11 4 1/2 x 10 3/8
        DMPAPER_ENV_12 = 22, // Envelope #12 4 /276 x 11
        DMPAPER_ENV_14 = 23, // Envelope #14 5 x 11 1/2
        DMPAPER_CSHEET = 24, // C size sheet
        DMPAPER_DSHEET = 25, // D size sheet
        DMPAPER_ESHEET = 26, // E size sheet
        DMPAPER_ENV_DL = 27, // Envelope DL 110 x 220mm
        DMPAPER_ENV_C5 = 28, // Envelope C5 162 x 229 mm
        DMPAPER_ENV_C3 = 29, // Envelope C3 324 x 458 mm
        DMPAPER_ENV_C4 = 30, // Envelope C4 229 x 324 mm
        DMPAPER_ENV_C6 = 31, // Envelope C6 114 x 162 mm
        DMPAPER_ENV_C65 = 32, // Envelope C65 114 x 229 mm
        DMPAPER_ENV_B4 = 33, // Envelope B4 250 x 353 mm
        DMPAPER_ENV_B5 = 34, // Envelope B5 176 x 250 mm
        DMPAPER_ENV_B6 = 35, // Envelope B6 176 x 125 mm
        DMPAPER_ENV_ITALY = 36, // Envelope 110 x 230 mm
        DMPAPER_ENV_MONARCH = 37, // Envelope Monarch 3.875 x 7.5 in
        DMPAPER_ENV_PERSONAL = 38, // 6 3/4 Envelope 3 5/8 x 6 1/2 in
        DMPAPER_FANFOLD_US = 39, // US Std Fanfold 14 7/8 x 11 in
        DMPAPER_FANFOLD_STD_GERMAN = 40, // German Std Fanfold 8 1/2 x 12 in
        DMPAPER_FANFOLD_LGL_GERMAN = 41, // German Legal Fanfold 8 1/2 x 13 in
        DMPAPER_USER = 256, // user defined
        DMPAPER_FIRST = DMPAPER_LETTER,
        DMPAPER_LAST = DMPAPER_USER,
    }


    /// <summary>
    /// Paper source
    /// </summary>
    public enum PaperSource
    {
        DMBIN_UPPER = 1,
        DMBIN_LOWER = 2,
        DMBIN_MIDDLE = 3,
        DMBIN_MANUAL = 4,
        DMBIN_ENVELOPE = 5,
        DMBIN_ENVMANUAL = 6,
        DMBIN_AUTO = 7,
        DMBIN_TRACTOR = 8,
        DMBIN_SMALLFMT = 9,
        DMBIN_LARGEFMT = 10,
        DMBIN_LARGECAPACITY = 11,
        DMBIN_CASSETTE = 14,
        DMBIN_FORMSOURCE = 15,
        DMRES_DRAFT = -1,
        DMRES_LOW = -2,
        DMRES_MEDIUM = -3,
        DMRES_HIGH = -4
    }


    /// <summary>
    /// Whether to print on both sides and other information
    /// </summary>
    public enum PageDuplex
    {
        DMDUP_HORIZONTAL = 3,
        DMDUP_SIMPLEX = 1,
        DMDUP_VERTICAL = 2
    }


    /// <summary>
    /// Printer settings
    /// </summary>
    public struct PrinterSettingsInfo
    {
        public PageOrientation Orientation; //Printing direction
        public PaperSize Size; //Print paper type
        public PaperSource source; //paper source
        public PageDuplex Duplex; //Whether to print on both sides and other information
        public int pLength; //Height of paper
        public int pWidth; //Paper width
        public int pmFields; //information that needs to be changed "|" sum after operation
        public string pFormName; //Paper name
    }

    //PRINTER_INFO_2 - The printer information structure contains 1..9 levels，Please refer to API for details
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct PRINTER_INFO_2
    {
        [MarshalAs(UnmanagedType.LPStr)] public string pServerName;
        [MarshalAs(UnmanagedType.LPStr)] public string pPrinterName;
        [MarshalAs(UnmanagedType.LPStr)] public string pShareName;
        [MarshalAs(UnmanagedType.LPStr)] public string pPortName;
        [MarshalAs(UnmanagedType.LPStr)] public string pDriverName;
        [MarshalAs(UnmanagedType.LPStr)] public string pComment;
        [MarshalAs(UnmanagedType.LPStr)] public string pLocation;
        public IntPtr pDevMode;
        [MarshalAs(UnmanagedType.LPStr)] public string pSepFile;
        [MarshalAs(UnmanagedType.LPStr)] public string pPrintProcessor;
        [MarshalAs(UnmanagedType.LPStr)] public string pDatatype;
        [MarshalAs(UnmanagedType.LPStr)] public string pParameters;
        public IntPtr pSecurityDescriptor;
        public Int32 Attributes;
        public Int32 Priority;
        public Int32 DefaultPriority;
        public Int32 StartTime;
        public Int32 UntilTime;
        public Int32 Status;
        public Int32 cJobs;
        public Int32 AveragePPM;
    }


    //PRINTER_INFO_5 - The printer information structure contains 1..9 levels，Please refer to API for details
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct PRINTER_INFO_5
    {
        [MarshalAs(UnmanagedType.LPTStr)] public String PrinterName;
        [MarshalAs(UnmanagedType.LPTStr)] public string? PortName;
        [MarshalAs(UnmanagedType.U4)] public Int32 Attributes;
        [MarshalAs(UnmanagedType.U4)] public Int32 DeviceNotSelectedTimeout;
        [MarshalAs(UnmanagedType.U4)] public Int32 TransmissionRetryTimeout;
    }


    //PRINTER_INFO_9
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    internal struct PRINTER_INFO_9
    {
        public IntPtr pDevMode;
    }

    /// <summary>
    /// The DEVMODE data structure contains information about the initialization and environment of a printer or a display device
    ///The DEVMODE structure contains the printer（or display settings)Initialization and current status information,Please refer to API for details
    /// </summary>
    private const short CCDEVICENAME = 32;

    private const short CCFORMNAME = 32;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public struct DEVMODE
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCDEVICENAME)]
        public string dmDeviceName;

        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public short dmOrientation;
        public short dmPaperSize;
        public short dmPaperLength;
        public short dmPaperWidth;
        public short dmScale;
        public short dmCopies;
        public short dmDefaultSource;
        public short dmPrintQuality;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCFORMNAME)]
        public string dmFormName;

        public short dmUnusedPadding;
        public short dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
    }

    //SendMessageTimeout Flags
    [Flags]
    public enum SendMessageTimeoutFlags : uint
    {
        SMTO_NORMAL = 0x0000,
        SMTO_BLOCK = 0x0001,
        SMTO_ABORTIFHUNG = 0x0002,
        SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
    }
    #endregion

    #region "const Variables"
    //DEVMODE.dmFields
    const int DM_FORMNAME = 0x10000; //When changing the paper name, you need to set this constant in dmFields
    const int DM_PAPERSIZE = 0x0002; //When changing the paper type, you need to set this constant in dmFields
    const int DM_PAPERLENGTH = 0x0004; //When changing the paper length, you need to set this constant in dmFields
    const int DM_PAPERWIDTH = 0x0008; //When changing the paper width, you need to set this constant in dmFields
    const int DM_DUPLEX = 0x1000; //When changing whether the paper is printed on both sides, you need to set this constant in dmFields
    const int DM_ORIENTATION = 0x0001; //When changing the paper orientation, you need to set this constant in dmFields

    //Used to change the parameters of DocumentProperties. Please refer to the API for details.
    const int DM_IN_BUFFER = 8;
    const int DM_OUT_BUFFER = 2;

    //Used to set access permissions to the printer
    const int PRINTER_ACCESS_ADMINISTER = 0x4;
    const int PRINTER_ACCESS_USE = 0x8;
    const int STANDARD_RIGHTS_REQUIRED = 0xF0000;
    const int PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE);

    //Get all papers specified for printing
    const int PRINTER_ENUM_LOCAL = 2;
    const int PRINTER_ENUM_CONNECTIONS = 4;
    const int DC_PAPERNAMES = 16;
    const int DC_PAPERS = 2;
    const int DC_PAPERSIZE = 3;

    //sendMessageTimeOut
    const int WM_SETTINGCHANGE = 0x001A;
    const int HWND_BROADCAST = 0xffff;
    #endregion

    #region printer method
    public bool OpenPrinterEx(string szPrinter, out IntPtr hPrinter, ref PRINTER_DEFAULTS pd)
    {
        bool bRet = OpenPrinter(szPrinter, out hPrinter, IntPtr.Zero);
        return bRet;
    }

    public DEVMODE GetPrinterDevMode(string? printerName)
    {
        if (string.IsNullOrEmpty(printerName))
        {
            printerName = GetDefaultPrinterName();
        }

        var pd = new PRINTER_DEFAULTS
        {
            pDatatype = 0,
            pDevMode = 0,
            DesiredAccess = PRINTER_ALL_ACCESS
        };
        // Michael: some printers (e.g. network printer) do not allow PRINTER_ALL_ACCESS and will cause Access Is Denied error.
        // When this happen, try PRINTER_ACCESS_USE.

        if (!OpenPrinterEx(printerName, out var hPrinter, ref pd))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out var nBytesNeeded);
        if (nBytesNeeded <= 0)
        {
            throw new Exception("Unable to allocate memory");
        }


        // Allocate enough space for PRINTER_INFO_2... {ptrPrinterIn fo = Marshal.AllocCoTaskMem(nBytesNeeded)};
        IntPtr ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded);

        // The second GetPrinter fills in all the current settings, so all you 
        // need to do is modify what you're interested in...
        nRet = Convert.ToInt32(GetPrinter(hPrinter, 2, ptrPrinterInfo, nBytesNeeded, out _));
        if (nRet == 0)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        var pinfo = (PRINTER_INFO_2)Marshal.PtrToStructure(ptrPrinterInfo, typeof(PRINTER_INFO_2))!;
        IntPtr Temp = new IntPtr();
        if (pinfo.pDevMode == IntPtr.Zero)
        {
            // If GetPrinter didn't fill in the DEVMODE, try to get it by calling
            // DocumentProperties...
            IntPtr ptrZero = IntPtr.Zero;
            //get the size of the devmode structure
            int sizeOfDevMode = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);

            IntPtr ptrDM = Marshal.AllocCoTaskMem(sizeOfDevMode);
            var i = DocumentProperties(IntPtr.Zero, hPrinter, printerName, ptrDM, ptrZero, DM_OUT_BUFFER);
            if ((i < 0) || (ptrDM == IntPtr.Zero))
            {
                //Cannot get the DEVMODE structure.
                throw new Exception("Cannot get DEVMODE data");
            }

            pinfo.pDevMode = ptrDM;
        }

        intError = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, Temp, 0);

        IntPtr yDevModeData = Marshal.AllocHGlobal(intError);
        intError = DocumentProperties(IntPtr.Zero, hPrinter, printerName, yDevModeData, Temp, 2);
        var dm = (DEVMODE)Marshal.PtrToStructure(yDevModeData, typeof(DEVMODE))!; //Retrieve printer device information from memory space
        if ((nRet == 0) || (hPrinter == IntPtr.Zero))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        ClosePrinter(hPrinter);
        return dm;
    }
    
    public bool IsPaperSize(string FormName, int width, int length)
    {
        DEVMODE dm = GetPrinterDevMode(null);
        if (FormName == dm.dmFormName && width == dm.dmPaperWidth && length == dm.dmPaperLength)
            return true;
        else
            return false;
    }
    
    public void ModifyPrinterSettings(string printerName, ref PrinterSettingsInfo prnSettings)
    {
        PRINTER_INFO_9 printerInfo;
        printerInfo.pDevMode = IntPtr.Zero;
        if (String.IsNullOrEmpty(printerName))
        {
            printerName = GetDefaultPrinterName();
        }

        var prnDefaults = new PRINTER_DEFAULTS
        {
            pDatatype = 0,
            pDevMode = 0,
            DesiredAccess = PRINTER_ALL_ACCESS
        };

        if (!OpenPrinterEx(printerName, out var hPrinter, ref prnDefaults))
        {
            return;
        }

        IntPtr ptrPrinterInfo = IntPtr.Zero;
        try
        {
            //Get the size of the structure DEVMODE
            int iDevModeSize = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
            if (iDevModeSize < 0)
                throw new ApplicationException("Cannot get the size of the DEVMODE structure.");

            //Allocate memory space buffer pointing to structure DEVMODE
            IntPtr hDevMode = Marshal.AllocCoTaskMem(iDevModeSize + 100);

            //Get a pointer to the DEVMODE structure
            nRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, hDevMode, IntPtr.Zero, DM_OUT_BUFFER);
            if (nRet < 0)
                throw new ApplicationException("Cannot get the size of the DEVMODE structure.");
            //Assign value to dm
            DEVMODE dm = (DEVMODE)Marshal.PtrToStructure(hDevMode, typeof(DEVMODE))!;

            if ((((int)prnSettings.Duplex < 0) || ((int)prnSettings.Duplex > 3)))
            {
                throw new ArgumentOutOfRangeException("prnSettings.Duplex", "nDuplexSetting is incorrect.");
            }
            else
            {
                // Change printer settings
                if ((int)prnSettings.Size != 0) //Whether to change paper type
                {
                    dm.dmPaperSize = (short)prnSettings.Size;
                    dm.dmFields |= DM_PAPERSIZE;
                }

                if (prnSettings.pWidth != 0) //Whether to change the paper width
                {
                    dm.dmPaperWidth = (short)prnSettings.pWidth;
                    dm.dmFields |= DM_PAPERWIDTH;
                }

                if (prnSettings.pLength != 0) //Whether to change the paper height
                {
                    dm.dmPaperLength = (short)prnSettings.pLength;
                    dm.dmFields |= DM_PAPERLENGTH;
                }

                if (!String.IsNullOrEmpty(prnSettings.pFormName)) //Whether to change the paper name
                {
                    dm.dmFormName = prnSettings.pFormName;
                    dm.dmFields |= DM_FORMNAME;
                }

                if ((int)prnSettings.Orientation != 0) //Whether to change the paper orientation
                {
                    dm.dmOrientation = (short)prnSettings.Orientation;
                    dm.dmFields |= DM_ORIENTATION;
                }

                Marshal.StructureToPtr(dm, hDevMode, true);

                //Get the size of printer info
                nRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, printerInfo.pDevMode, printerInfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);
                if (nRet < 0)
                {
                    throw new ApplicationException("Unable to set the PrintSetting for this printer");
                }

                GetPrinter(hPrinter, 9, IntPtr.Zero, 0, out var nBytesNeeded);
                if (nBytesNeeded == 0)
                    throw new ApplicationException("GetPrinter failed.Couldn't get the nBytesNeeded for shared PRINTER_INFO_9 structure");

                //Configure memory block
                ptrPrinterInfo = Marshal.AllocCoTaskMem(nBytesNeeded);
                bool bSuccess = GetPrinter(hPrinter, 9, ptrPrinterInfo, nBytesNeeded, out _);
                if (!bSuccess)
                    throw new ApplicationException("GetPrinter failed.Couldn't get the nBytesNeeded for shared PRINTER_INFO_9 structure");
                //Assign value to printerInfo
                printerInfo = (PRINTER_INFO_9)Marshal.PtrToStructure(ptrPrinterInfo, printerInfo.GetType())!;
                printerInfo.pDevMode = hDevMode;

                //Gets a pointer to a PRINTER_INFO_9 structure
                Marshal.StructureToPtr(printerInfo, ptrPrinterInfo, true);

                //Set up the printer
                bSuccess = SetPrinter(hPrinter, 9, ptrPrinterInfo, 0);
                if (!bSuccess)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "SetPrinter() failed.Couldn't set the printer settings");

                // Set the printer to notify other apps that printer settings have been changed -- Do NOT use because it causes app halt serveral seconds!!
                /*
                PrinterHelper.SendMessageTimeout(
                    new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, IntPtr.Zero,
                    PrinterHelper.SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out hDummy);
                 */
            }
        }
        finally
        {
            ClosePrinter(hPrinter);

            //Free memory
            if (ptrPrinterInfo == IntPtr.Zero)
                Marshal.FreeHGlobal(ptrPrinterInfo);
            if (hPrinter == IntPtr.Zero)
                Marshal.FreeHGlobal(hPrinter);
        }
    }

    
    public bool ModifyPrinterSettings_V2(string printerName, ref PrinterSettingsInfo printerSetting)
    {
        PRINTER_DEFAULTS pd = new PRINTER_DEFAULTS
        {
            pDatatype = 0,
            pDevMode = 0,
            DesiredAccess = PRINTER_ALL_ACCESS
        };
        if (String.IsNullOrEmpty(printerName))
        {
            printerName = GetDefaultPrinterName();
        }

        if (!OpenPrinterEx(printerName, out var hPrinter, ref pd))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        //Call GetPrinter to get the number of bytes of PRINTER_INFO_2 in the memory space
        GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out var nBytesNeeded);
        if (nBytesNeeded <= 0)
        {
            ClosePrinter(hPrinter);
            return false;
        }

        //Allocate enough memory space for PRINTER_INFO_2
        IntPtr ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded);
        if (ptrPrinterInfo == IntPtr.Zero)
        {
            ClosePrinter(hPrinter);
            return false;
        }

        //Call GetPrinter to populate the current settings, which is the information you want to change (ptrPrinterInfo)
        if (!GetPrinter(hPrinter, 2, ptrPrinterInfo, nBytesNeeded, out nBytesNeeded))
        {
            Marshal.FreeHGlobal(ptrPrinterInfo);
            ClosePrinter(hPrinter);
            return false;
        }

        //Convert the pointer pointing to PRINTER_INFO_2 in the memory block into the PRINTER_INFO_2 structure
        //If GetPrinter does not get the DEVMODE structure, it will try to get the DEVMODE structure through DocumentProperties
        var pinfo = (PRINTER_INFO_2)Marshal.PtrToStructure(ptrPrinterInfo, typeof(PRINTER_INFO_2))!;
        if (pinfo.pDevMode == IntPtr.Zero)
        {
            // If GetPrinter didn't fill in the DEVMODE, try to get it by calling
            // DocumentProperties...
            IntPtr ptrZero = IntPtr.Zero;
            //get the size of the devmode structure
            nBytesNeeded = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
            if (nBytesNeeded <= 0)
            {
                Marshal.FreeHGlobal(ptrPrinterInfo);
                ClosePrinter(hPrinter);
                return false;
            }

            IntPtr ptrDM = Marshal.AllocCoTaskMem(nBytesNeeded);
            var i = DocumentProperties(IntPtr.Zero, hPrinter, printerName, ptrDM, ptrZero, DM_OUT_BUFFER);
            if ((i < 0) || (ptrDM == IntPtr.Zero))
            {
                //Cannot get the DEVMODE structure.
                Marshal.FreeHGlobal(ptrDM);
                ClosePrinter(ptrPrinterInfo);
                return false;
            }

            pinfo.pDevMode = ptrDM;
        }

        DEVMODE dm = (DEVMODE)Marshal.PtrToStructure(pinfo.pDevMode, typeof(DEVMODE))!;

        //Modify printer settings information
        if ((((int)printerSetting.Duplex < 0) || ((int)printerSetting.Duplex > 3)))
        {
            throw new ArgumentOutOfRangeException("printerSetting.Duplex", "nDuplexSetting is incorrect.");
        }
        else
        {
            if (String.IsNullOrEmpty(printerName))
            {
                printerName = GetDefaultPrinterName();
            }

            if ((int)printerSetting.Size != 0) //Whether to change paper type
            {
                dm.dmPaperSize = (short)printerSetting.Size;
                dm.dmFields |= DM_PAPERSIZE;
            }

            if (printerSetting.pWidth != 0) //Whether to change the paper width
            {
                dm.dmPaperWidth = (short)printerSetting.pWidth;
                dm.dmFields |= DM_PAPERWIDTH;
            }

            if (printerSetting.pLength != 0) //Whether to change the paper height
            {
                dm.dmPaperLength = (short)printerSetting.pLength;
                dm.dmFields |= DM_PAPERLENGTH;
            }

            if (!String.IsNullOrEmpty(printerSetting.pFormName)) //Whether to change the paper name
            {
                dm.dmFormName = printerSetting.pFormName;
                dm.dmFields |= DM_FORMNAME;
            }

            if ((int)printerSetting.Orientation != 0) //Whether to change the paper orientation
            {
                dm.dmOrientation = (short)printerSetting.Orientation;
                dm.dmFields |= DM_ORIENTATION;
            }

            Marshal.StructureToPtr(dm, pinfo.pDevMode, true);
            Marshal.StructureToPtr(pinfo, ptrPrinterInfo, true);
            pinfo.pSecurityDescriptor = IntPtr.Zero;
            //Make sure the driver_Dependent part of devmode is updated...
            nRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, pinfo.pDevMode, pinfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);
            if (nRet <= 0)
            {
                Marshal.FreeHGlobal(ptrPrinterInfo);
                ClosePrinter(hPrinter);
                return false;
            }

            //Update printer information
            if (!SetPrinter(hPrinter, 2, ptrPrinterInfo, 0))
            {
                Marshal.FreeHGlobal(ptrPrinterInfo);
                ClosePrinter(hPrinter);
                return false;
            }

            //Notify other applications that printer information has changed
            Winspooler.SendMessageTimeout(
                new IntPtr(HWND_BROADCAST), WM_SETTINGCHANGE, IntPtr.Zero, IntPtr.Zero,
                SendMessageTimeoutFlags.SMTO_NORMAL, 1000, out _);

            //Free memory
            if (ptrPrinterInfo == IntPtr.Zero)
                Marshal.FreeHGlobal(ptrPrinterInfo);
            if (hPrinter == IntPtr.Zero)
                Marshal.FreeHGlobal(hPrinter);

            return true;
        }
    }
    
    public string GetDefaultPrinterName()
    {
        StringBuilder dp = new StringBuilder(256);
        int size = dp.Capacity;
        if (GetDefaultPrinter(dp, ref size))
        {
            return dp.ToString();
        }
        else
        {
            return string.Empty;
        }
    }
    
    public short GetOnePaper(string printerName, string paperName)
    {
        short kind = 0;
        if (String.IsNullOrEmpty(printerName))
            printerName = GetDefaultPrinterName();
        EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
            string.Empty, 5, IntPtr.Zero, 0, out var requiredSize, out var numPrinters);

        int info5Size = requiredSize;
        IntPtr info5Ptr = Marshal.AllocHGlobal(info5Size);
        try
        {
            EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                string.Empty, 5, info5Ptr, info5Size, out requiredSize, out numPrinters);

            string? port = null;
            for (int i = 0; i < numPrinters; i++)
            {
                var info5 = (PRINTER_INFO_5)Marshal.PtrToStructure(
                    (i * Marshal.SizeOf(typeof(PRINTER_INFO_5))) + (int)info5Ptr,
                    typeof(PRINTER_INFO_5))!;
                if (info5.PrinterName == printerName)
                {
                    port = info5.PortName;
                }
            }

            int numNames = DeviceCapabilities(printerName, port!, DC_PAPERNAMES, IntPtr.Zero, IntPtr.Zero);
            if (numNames < 0)
            {
                int errorCode = GetLastError();
                Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                return 0;
            }

            var buffer = Marshal.AllocHGlobal(numNames * 64);
            numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, buffer, IntPtr.Zero);
            if (numNames < 0)
            {
                int errorCode = GetLastError();
                Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                return 0;
            }

            string[] names = new string[numNames];
            for (int i = 0; i < numNames; i++)
            {
                names[i] = Marshal.PtrToStringAnsi((i * 64) + (int)buffer)!;
            }

            Marshal.FreeHGlobal(buffer);

            int numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, IntPtr.Zero, IntPtr.Zero);
            if (numPapers < 0)
            {
                Console.WriteLine("No papers");
                return 0;
            }

            buffer = Marshal.AllocHGlobal(numPapers * 2);
            numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, buffer, IntPtr.Zero);
            if (numPapers < 0)
            {
                Console.WriteLine("No papers");
                return 0;
            }

            short[] kinds = new short[numPapers];
            for (int i = 0; i < numPapers; i++)
            {
                kinds[i] = Marshal.ReadInt16(buffer, i * 2);
            }

            for (int i = 0; i < numPapers; i++)
            {
                if (names[i] == paperName)
                {
                    kind = kinds[i];
                    break;
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(info5Ptr);
        }

        return kind;
    }
    
    public void ShowPapers(string printerName)
    {
        if (String.IsNullOrEmpty(printerName))
        {
            printerName = GetDefaultPrinterName();
        }

        EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
            string.Empty, 5, IntPtr.Zero, 0, out var requiredSize, out var numPrinters);

        int info5Size = requiredSize;
        IntPtr info5Ptr = Marshal.AllocHGlobal(info5Size);
        try
        {
            EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                string.Empty, 5, info5Ptr, info5Size, out requiredSize, out numPrinters);

            string? port = null;
            for (int i = 0; i < numPrinters; i++)
            {
                var info5 = (PRINTER_INFO_5)Marshal.PtrToStructure(
                    (i * Marshal.SizeOf(typeof(PRINTER_INFO_5))) + (int)info5Ptr,
                    typeof(PRINTER_INFO_5))!;
                if (info5.PrinterName == printerName)
                {
                    port = info5.PortName;
                }
            }

            int numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, IntPtr.Zero, IntPtr.Zero);
            if (numNames < 0)
            {
                int errorCode = GetLastError();
                Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                return;
            }

            var buffer = Marshal.AllocHGlobal(numNames * 64);
            numNames = DeviceCapabilities(printerName, port, DC_PAPERNAMES, buffer, IntPtr.Zero);
            if (numNames < 0)
            {
                int errorCode = GetLastError();
                Console.WriteLine("Number of names = {1}: {0}", errorCode, numNames);
                return;
            }

            string[] names = new string[numNames];
            for (int i = 0; i < numNames; i++)
            {
                names[i] = Marshal.PtrToStringAnsi((i * 64) + (int)buffer)!;
            }

            Marshal.FreeHGlobal(buffer);

            int numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, IntPtr.Zero, IntPtr.Zero);
            if (numPapers < 0)
            {
                Console.WriteLine("No papers");
                return;
            }

            buffer = Marshal.AllocHGlobal(numPapers * 2);
            numPapers = DeviceCapabilities(printerName, port, DC_PAPERS, buffer, IntPtr.Zero);
            if (numPapers < 0)
            {
                Console.WriteLine("No papers");
                return;
            }

            short[] kinds = new short[numPapers];
            for (int i = 0; i < numPapers; i++)
            {
                kinds[i] = Marshal.ReadInt16(buffer, i * 2);
            }

            for (int i = 0; i < numPapers; i++)
            {
                Console.WriteLine("Paper {0} : {1}", kinds[i], names[i]);
            }
        }
        finally
        {
            Marshal.FreeHGlobal(info5Ptr);
        }
    }

    public void AbortPrinter(string printerName)
    {
        PRINTER_DEFAULTS pd = new PRINTER_DEFAULTS
        {
            pDatatype = 0,
            pDevMode = 0,
            DesiredAccess = PRINTER_ALL_ACCESS
        };
        if (String.IsNullOrEmpty(printerName))
        {
            printerName = GetDefaultPrinterName();
        }

        if (!OpenPrinterEx(printerName, out var hPrinter, ref pd))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        if (!AbortPrinter(hPrinter))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        ClosePrinter(hPrinter);
        if (hPrinter == IntPtr.Zero)
            Marshal.FreeHGlobal(hPrinter);
    }

    public void DeleteAllJobs(string printerName)
    {
        PRINTER_DEFAULTS pd = new PRINTER_DEFAULTS
        {
            pDatatype = 0,
            pDevMode = 0,
            DesiredAccess = PRINTER_ALL_ACCESS
        };
        if (String.IsNullOrEmpty(printerName))
        {
            printerName = GetDefaultPrinterName();
        }

        if (!OpenPrinterEx(printerName, out var hPrinter, ref pd))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        /* Query for 99 jobs */
        const uint firstJob = 0u;
        const uint noJobs = 99u;
        const uint level = 1u;

        // Get byte size required for the function
        //Ignore errors, this call is just to get the size of the buffer needed
        EnumJobs(hPrinter, firstJob, noJobs, level, IntPtr.Zero, 0, out var needed, out _);

        // Populate the structs
        IntPtr pJob = Marshal.AllocHGlobal((int)needed);
        if (!EnumJobs(hPrinter, firstJob, noJobs, level, pJob, needed, out _, out var structsCopied))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        var jobInfos = new JOB_INFO_1W[structsCopied];
        int sizeOf = Marshal.SizeOf(typeof(JOB_INFO_1W));
        IntPtr pStruct = pJob;
        for (int i = 0; i < structsCopied; i++)
        {
            var jobInfo_1W = (JOB_INFO_1W)Marshal.PtrToStructure(pStruct, typeof(JOB_INFO_1W))!;
            jobInfos[i] = jobInfo_1W;
            pStruct += sizeOf;
        }

        Marshal.FreeHGlobal(pJob);
        _logger.LogDebug("Found {0} jobs", jobInfos.Length);

        foreach (var jobInfo in jobInfos)
        {
            _logger.LogDebug("Deleting job {0}", jobInfo.JobId);
            SetJob(hPrinter, (int)jobInfo.JobId, 0, IntPtr.Zero, JOB_CONTROL_DELETE);
        }

        ClosePrinter(hPrinter);
        if (hPrinter == IntPtr.Zero)
            Marshal.FreeHGlobal(hPrinter);
    }
    #endregion
}
