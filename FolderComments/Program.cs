// <copyright file="MainForm.cs" company="PublicDomain.is">
//     CC0 1.0 Universal (CC0 1.0) - Public Domain Dedication
//     https://creativecommons.org/publicdomain/zero/1.0/legalcode
// </copyright>

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace FolderComments
{
    /// <summary>
    /// Class with program entry point.
    /// </summary>
    internal sealed class Program
    {
        /// <summary>
        /// The fcs read.
        /// </summary>
        private static UInt32 FCS_READ = 0x00000001;

        /// <summary>
        /// The fcs forcewrite.
        /// </summary>
        private static UInt32 FCS_FORCEWRITE = 0x00000002;

        /// <summary>
        /// The fcsm infotip.
        /// </summary>
        private static UInt32 FCSM_INFOTIP = 0x00000004;

        /// <summary>
        /// SHGs the et set folder custom settings.
        /// </summary>
        /// <returns>The et set folder custom settings.</returns>
        /// <param name="pfcs">Pfcs.</param>
        /// <param name="pszPath">Psz path.</param>
        /// <param name="dwReadWrite">Dw read write.</param>
        [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        private static extern UInt32 SHGetSetFolderCustomSettings(ref LPSHFOLDERCUSTOMSETTINGS pfcs, string pszPath, UInt32 dwReadWrite);

        /// <summary>
        /// Lpshfoldercustomsettings.
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        struct LPSHFOLDERCUSTOMSETTINGS
        {
            public UInt32 dwSize;
            public UInt32 dwMask;
            public IntPtr pvid;
            public string pszWebViewTemplate;
            public UInt32 cchWebViewTemplate;
            public string pszWebViewTemplateVersion;
            public string pszInfoTip;
            public UInt32 cchInfoTip;
            public IntPtr pclsid;
            public UInt32 dwFlags;
            public string pszIconFile;
            public UInt32 cchIconFile;
            public int iIconIndex;
            public string pszLogo;
            public UInt32 cchLogo;
        }

        /// <summary>
        /// Program entry point.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            // Check for arguments and valid directory
            if (args.Length > 0)
            {
                // Set directory path
                string directoryPath = args[0];

                // Check it's a valid directory path
                if (!Directory.Exists(directoryPath))
                {
                    // Halt flow
                    return;
                }

                // Error handling & logging
                try
                {
                    // Declare current comments
                    string currentComments = string.Empty;

                    /* Get current comments */

                    // Set buffer size
                    const int BUFFERSIZE = 4096;

                    // Get settings
                    LPSHFOLDERCUSTOMSETTINGS folderCustomSettings = new LPSHFOLDERCUSTOMSETTINGS
                    {
                        dwMask = FCSM_INFOTIP,
                        pszInfoTip = new String(' ', BUFFERSIZE),
                        cchInfoTip = BUFFERSIZE
                    };

                    // Set dwSize
                    folderCustomSettings.dwSize = (uint)Marshal.SizeOf(folderCustomSettings);

                    // Read folder custom settings
                    UInt32 HRESULT = SHGetSetFolderCustomSettings(ref folderCustomSettings, directoryPath, FCS_READ);

                    // Check if OK
                    if (HRESULT == 0)
                    {
                        // Get previous directory comments
                        currentComments = folderCustomSettings.pszInfoTip;
                    }

                    /* Set current comments */

                    // Let user edit the comments
                    currentComments = Interaction.InputBox("Set new folder comments", "Edit comments", currentComments);

                    // Set settings
                    LPSHFOLDERCUSTOMSETTINGS FolderCustomSettings = new LPSHFOLDERCUSTOMSETTINGS
                    {
                        dwMask = FCSM_INFOTIP,
                        pszInfoTip = currentComments.Length > 0 ? currentComments : null,
                        cchInfoTip = 0
                    };

                    // WRITE folder custom settings
                    HRESULT = SHGetSetFolderCustomSettings(ref FolderCustomSettings, directoryPath, FCS_FORCEWRITE);

                    // TODO Success, advise user [Perhaps have it set as an option in a future version]
                    // UNDONE MessageBox.Show($"Folder comments were successfully set ", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception exception)
                {
                    // Advise user
                    MessageBox.Show($"Setting folder comments failed.{Environment.NewLine}Check error log for detailed info.", "Folder comments error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    try
                    {
                        // Log error event
                        File.AppendAllText("FolderComments-ErrorLog.txt", $"Setting folder comments failed. Path: {directoryPath}{Environment.NewLine}Message: {exception.Message}{Environment.NewLine}{Environment.NewLine}");
                    }
                    catch (Exception fileAppendException)
                    {
                        // Advise user
                        MessageBox.Show($"Error when writing \"FolderComments-ErrorLog.tx\" file.{Environment.NewLine}{Environment.NewLine}Message:{Environment.NewLine}{fileAppendException.Message}", "File append error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
        }
    }

}
