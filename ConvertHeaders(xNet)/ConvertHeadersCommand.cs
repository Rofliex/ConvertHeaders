using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace ConvertHeaders_xNet_
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ConvertHeadersCommand
    {
        static string[] definedHeaders = new[] { "Accept", "AcceptCharset", "AcceptLanguage", "AcceptDatetime", "CacheControl", "ContentType", "Date", "Expect", "From", "IfMatch", "IfModifiedSince", "IfNoneMatch", "IfRange", "IfUnmodifiedSince", "MaxForwards", "Pragma", "Range", "Referer", "Origin", "Upgrade", "UpgradeInsecureRequests", "UserAgent", "Via", "Warning", "DNT", "AccessControlAllowOrigin", "AcceptRanges", "Age", "Allow", "ContentEncoding", "ContentLanguage", "ContentLength", "ContentLocation", "ContentMD5", "ContentDisposition", "ContentRange", "ETag", "Expires", "LastModified", "Link", "Location", "P3P", "Refresh", "RetryAfter", "Server", "TransferEncoding" };
        private static string[] unprocessedCookie = { "Cookie", "Content-Length", "Host" };

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a5e22065-9dad-43a9-b7eb-23288010c8ee");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertHeadersCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ConvertHeadersCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ConvertHeadersCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ConvertHeadersCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ConvertHeadersCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                EnvDTE.DTE dte = (EnvDTE.DTE)ServiceProvider.GetServiceAsync(typeof(EnvDTE.DTE)).ConfigureAwait(false).GetAwaiter().GetResult();
                var textSelection = (TextSelection)dte.ActiveDocument.Selection;
                var headers = Clipboard.GetText()
                    .Split(new String[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(x => x.Count(c => c == ':') == 1)
                    .Select(x => x.Split(':'))
                    .Where(x => !unprocessedCookie.Contains(x[0]));


                List<string> resultCode = new List<string>();

                foreach (var header in headers)
                {

                    if (definedHeaders.Contains(header[0].Replace("-", "")))
                    {
                        resultCode.Add($"httpRequest.AddHeader(HttpHeader.{header[0].Replace("-", "")},\"{header[1].TrimStart(' ')}\");");
                    }
                    else
                    {
                        resultCode.Add($"httpRequest.AddHeader(\"{header[0]}\",\"{header[1].TrimStart(' ')}\");");
                    }
                }

                textSelection.Insert(string.Join("\r\n", resultCode) + "\r\n");
                textSelection.SelectAll();
                textSelection.SmartFormat();
                textSelection.Cancel();
            }
            catch { }
        }
    }
}
