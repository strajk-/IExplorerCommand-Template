using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
using Windows.Management.Deployment;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Xml.Linq;
using Microsoft.Win32;

namespace CustomShellManager
{
    public class StartUp
    {
        const string PackageName = "CustomShell";

        [STAThread]
        public static void Main (string[] cmdArgs)
        {
            // Only run application logic if it isn't ran with Identity (not as an MSIX package)
            if (!ExecutionMode.IsRunningWithIdentity()) {
                // Define supported commands
                HashSet<string> validCommands = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                    "install", "i",
                    "uninstall", "u",
                    "installcert", "ic",
                    "uninstallcert", "uc",
                    "help", "?"
                };

                // Build parameter dictionary
                Dictionary<string, string> argsDict = new Dictionary<string, string>();
                string param = string.Empty;
                foreach (string args in cmdArgs) {
                    string argsParsed = args.Trim();
                    if (string.IsNullOrEmpty(argsParsed)) continue;
                    if (argsParsed.StartsWith("-") || argsParsed.StartsWith("/")) {
                        if (!string.IsNullOrEmpty(param)) {
                            if (argsDict.Count > 0) {
                                string lastKey = argsDict.Keys.Last();
                                argsDict[lastKey] = param;
                            }
                            param = string.Empty;
                        }
                        // Store args without the / or -
                        string newArg = argsParsed.Substring(1).Trim();
                        if (newArg.Length > 0) {
                            if (!argsDict.ContainsKey(newArg)) {
                                argsDict.Add(newArg, string.Empty);
                            }
                        }
                    } else param += args;
                }
                if (!string.IsNullOrEmpty(param)) {
                    if (argsDict.Count > 0) {
                        string lastKey = argsDict.Keys.Last();
                        argsDict[lastKey] = param;
                    }
                    param = string.Empty;
                }

                // Check if any commands were used that aren't supported
                bool hasInvalid = argsDict.Keys.Any(k => !validCommands.Contains(k));
                if (hasInvalid) {
                    argsDict.Clear();
                }

                // Check if Dictionary is empty
                if (argsDict.Count == 0) {
                    if (cmdArgs.Count() > 0) {
                        // Parameters were given but not recognized
                        argsDict.Add("help", string.Empty);
                    } else {
                        // Default if no parameter was given
                        argsDict.Add("install", string.Empty);
                    }
                }

                // Check if help was used
                if (argsDict.ContainsKey("?") || argsDict.ContainsKey("help")) {
                    string exeName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);
                    Console.WriteLine("Usage:");
                    Console.WriteLine($"  {exeName} [-command]");
                    Console.WriteLine();
                    var commands = new (string[] keys, string description)[]
                    {
                        (new[] {"-install", "-i"}, "Register the MSIX package"),
                        (new[] {"-uninstall", "-u"}, "Unregister the MSIX package"),
                        (new[] {"-installcert", "-u"}, "Installs Root Cert of this application (if one is present) on the LocalMachine"),
                        (new[] {"-uninstallcert", "-u"}, "Unistalls Root Cert of this application (if one is present) on the LocalMachine"),
                        (new[] {"-help", "-?"}, "Display this help message")
                    };

                    // Determine width of the keys column
                    int keyWidth = commands.Max(c => string.Join(", ", c.keys).Length);

                    Console.WriteLine("Commands:");
                    foreach (var cmd in commands) {
                        string keyText = string.Join(", ", cmd.keys);
                        Console.WriteLine($"  {keyText.PadRight(keyWidth)}  {cmd.description}");
                    }
                    return;
                }

                // Iterate through dictionary and run application based on parameters
                foreach (KeyValuePair<string, string> args in argsDict) {
                    switch (args.Key) {
                        case "install":
                        case "i": {
                                Console.WriteLine("Install parameter detected.");
                                string exePath = AppDomain.CurrentDomain.BaseDirectory;
                                string externalLocation = Path.Combine(exePath, @"");
                                string sparsePkgPath = Path.Combine(exePath, $"{PackageName}.msix");

                                // Install cert first if .exe has one, RegisterSparsePackage will not work if the Root CA is not trusted
                                InstallCert();

                                //Attempt registration
                                if (!RegisterSparsePackage(externalLocation, sparsePkgPath)) {
                                    Console.WriteLine("Package Registation failed while running WITHOUT Identity (not as MSIX)");
                                }
                            }
                            break;
                        case "uninstall":
                        case "u": {
                                Console.WriteLine("Uninstall parameter detected.");
                                RemoveContextMenuPackage();
                            }
                            break;

                        case "installcert":
                        case "ic": {
                                Console.WriteLine("Install Cert parameter detected.");
                                InstallCert();
                            }
                            break;
                        case "uninstallcert":
                        case "uc": {
                                Console.WriteLine("Uninstall Cert parameter detected.");
                                UninstallCert();
                            }
                            break;
                    }
                }
            } else {
                Console.WriteLine("Application is running off a package (MSIX) or Windows Version is older than 8, exiting...");
            }
        }

        private static void InstallCert()
        {
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            RootCertificateHandler rootCertificateHandler = new RootCertificateHandler(exePath);
            if (rootCertificateHandler.IsInit) {
                if (rootCertificateHandler.IsInstalled()) {
                    // Uninstall if present for reinstall
                    rootCertificateHandler.Uninstall();
                }
                rootCertificateHandler.Install();
            }
        }

        private static void UninstallCert()
        {
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            RootCertificateHandler rootCertificateHandler = new RootCertificateHandler(exePath);
            if (rootCertificateHandler.IsInit) {
                if (rootCertificateHandler.IsInstalled()) {
                    rootCertificateHandler.Uninstall();
                } else Console.WriteLine("Root CA not found on LocalMachine");
            }
        }

        [DllImport("Shell32.dll")]
        public static extern void SHChangeNotify (uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        private const uint SHCNE_ASSOCCHANGED = 0x8000000;
        private const uint SHCNF_IDLIST = 0x0;

        private static bool RegisterSparsePackage (string externalLocation, string sparsePkgPath)
        {
            bool registration = false;
            try {
                Uri externalUri = new Uri(externalLocation);
                Uri packageUri = new Uri(sparsePkgPath);

                Console.WriteLine("exe Location {0}", externalLocation);
                Console.WriteLine();
                Console.WriteLine("msix Address {0}", sparsePkgPath);
                Console.WriteLine();

                Console.WriteLine("  exe Uri {0}", externalUri);
                Console.WriteLine();
                Console.WriteLine("  msix Uri {0}", packageUri);
                Console.WriteLine();

                PackageManager packageManager = new PackageManager();

                //Declare use of an external location
                var options = new AddPackageOptions();
                options.ExternalLocationUri = externalUri;

                Windows.Foundation.IAsyncOperationWithProgress<DeploymentResult, DeploymentProgress> deploymentOperation = packageManager.AddPackageByUriAsync(packageUri, options);

                ManualResetEvent opCompletedEvent = new ManualResetEvent(false); // this event will be signaled when the deployment operation has completed.

                deploymentOperation.Completed = (depProgress, status) => { opCompletedEvent.Set(); };

                Console.WriteLine("Installing package {0}", sparsePkgPath);
                Console.WriteLine();

                Console.WriteLine("Waiting for package registration to complete...");
                Console.WriteLine();

                opCompletedEvent.WaitOne();

                if (deploymentOperation.Status == Windows.Foundation.AsyncStatus.Error) {
                    Windows.Management.Deployment.DeploymentResult deploymentResult = deploymentOperation.GetResults();
                    Console.WriteLine("Installation Error: {0}", deploymentOperation.ErrorCode);
                    Console.WriteLine();
                    Console.WriteLine("Detailed Error Text: {0}", deploymentResult.ErrorText);
                    Console.WriteLine();

                } else if (deploymentOperation.Status == Windows.Foundation.AsyncStatus.Canceled) {
                    Console.WriteLine("Package Registration Canceled");
                    Console.WriteLine();
                } else if (deploymentOperation.Status == Windows.Foundation.AsyncStatus.Completed) {
                    registration = true;
                    Console.WriteLine("Package Registration succeeded!");
                    Console.WriteLine();
                } else {
                    Console.WriteLine("Installation status unknown");
                    Console.WriteLine();
                }

                // Ensure registry keys for file extensions exist
                using (Package package = Package.Open(sparsePkgPath, FileMode.Open, FileAccess.Read)) {
                    Uri manifestUri = new Uri("/AppxManifest.xml", UriKind.Relative);
                    if (package.PartExists(manifestUri)) {
                        PackagePart manifestPart = package.GetPart(manifestUri);
                        XDocument doc;
                        using (var stream = manifestPart.GetStream()) {
                            doc = XDocument.Load(stream);
                        }

                        XNamespace desktop5 = "http://schemas.microsoft.com/appx/manifest/desktop/windows10/5";

                        foreach (var itemType in doc.Descendants(desktop5 + "ItemType")) {
                            string ext = itemType.Attribute("Type")?.Value;
                            if (!string.IsNullOrWhiteSpace(ext) && ext != "*") {
                                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(ext)) {
                                    if (key == null) {
                                        Registry.ClassesRoot.CreateSubKey(ext);
                                        Console.WriteLine($"Created missing registry key: {ext}");
                                    }
                                }
                            }
                        }
                    } else {
                        Console.WriteLine("AppxManifest.xml not found inside MSIX.");
                    }
                }

                if (registration) {
                    // Notify the shell about the change
                    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
                }
            }
            catch (Exception ex) {
                Console.WriteLine("AddPackageSample failed, error message: {0}", ex.Message);
                Console.WriteLine();
                Console.WriteLine("Full Stacktrace: {0}", ex.ToString());
                Console.WriteLine();

                return registration;
            }

            return registration;
        }

        public static void RemoveContextMenuPackage ()
        {
            try {
                PackageManager packageManager = new PackageManager();

                // Get all installed packages
                var packages = packageManager.FindPackages().Where(p => p.Id.Name.Equals(PackageName, StringComparison.OrdinalIgnoreCase));

                foreach (var package in packages) {
                    string manifestPath = Path.Combine(package.InstalledLocation.Path, "AppxManifest.xml");

                    if (!File.Exists(manifestPath)) continue;

                    // Load manifest XML
                    XmlDocument manifestXml = new XmlDocument();
                    manifestXml.Load(manifestPath);

                    // Check if this package actually uses fileExplorerContextMenus
                    if (manifestXml.OuterXml.Contains("windows.fileExplorerContextMenus")) {
                        Console.WriteLine($"Uninstalling package: {package.Id.FullName}");

                        var removeOperation = packageManager.RemovePackageAsync(package.Id.FullName);
                        ManualResetEvent opCompletedEvent = new ManualResetEvent(false);

                        removeOperation.Completed = (depProgress, status) => { opCompletedEvent.Set(); };
                        opCompletedEvent.WaitOne();

                        if (removeOperation.Status == Windows.Foundation.AsyncStatus.Completed) {
                            Console.WriteLine($"Successfully uninstalled: {package.Id.FullName}");
                        } else if (removeOperation.Status == Windows.Foundation.AsyncStatus.Error) {
                            var result = removeOperation.GetResults();
                            Console.WriteLine($"Error uninstalling {package.Id.FullName}: {removeOperation.ErrorCode}");
                            Console.WriteLine($"Detail: {result.ErrorText}");
                        }
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Failed to uninstall {PackageName} context menu package: {ex.Message}");
            }
        }
    }
}
