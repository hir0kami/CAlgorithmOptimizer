using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace CAlgorithmOptimizer
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class OptimizeCommand
    {
        public const int CommandId = 4129;
        public static readonly Guid CommandSet = new Guid("6802a089-07da-4c89-91f8-a717d1917ab7");
        private readonly AsyncPackage package;

        private OptimizeCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static OptimizeCommand Instance { get; private set; }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new OptimizeCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Получение активного документа через DTE
            var dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
            var activeDoc = dte?.ActiveDocument;

            // Проверка, открыт ли активный документ
            if (activeDoc != null)
            {
                // Получение текста текущего документа
                var textDoc = activeDoc.Object("TextDocument") as EnvDTE.TextDocument;
                var editPoint = textDoc?.CreateEditPoint();
                string text = editPoint?.GetText(textDoc.EndPoint) ?? string.Empty;

                // Регулярное выражение для поиска вложенных циклов
                var nestedLoopPattern = @"for\s*\(.*\)\s*{[^}]*for\s*\(.*\)";
                var match = System.Text.RegularExpressions.Regex.Match(text, nestedLoopPattern, System.Text.RegularExpressions.RegexOptions.Singleline);

                if (match.Success)
                {
                    // Показываем сообщение и вставляем шаблон
                    var template = GetTemplate("OpenMP");
                    InsertTemplate(template);

                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        "Обнаружены вложенные циклы. Шаблон OpenMP вставлен в документ.",
                        "C-Algorithm Optimizer",
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
                else
                {
                    // Если вложенные циклы не найдены
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        "Вложенные циклы не найдены. Оптимизация не требуется.",
                        "C-Algorithm Optimizer",
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
            else
            {
                // Сообщение об ошибке, если активный документ не найден
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "Не удалось открыть активный документ. Убедитесь, что файл открыт.",
                    "Ошибка",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private void InsertTemplate(string template)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
            var activeDoc = dte?.ActiveDocument;

            if (activeDoc != null)
            {
                var textDoc = activeDoc.Object("TextDocument") as EnvDTE.TextDocument;
                var editPoint = textDoc?.CreateEditPoint();
                editPoint?.Insert(template);
            }
            else
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "Не удалось открыть активный документ для вставки шаблона.",
                    "Ошибка",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private string GetTemplate(string type)
        {
            if (type == "OpenMP")
            {
                return @"
#pragma omp parallel for
for (int i = 0; i < n; ++i) {
    // Parallelized code here
}";
            }
            return "";
        }
    }
}
