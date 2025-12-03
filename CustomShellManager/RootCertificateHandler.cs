using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace CustomShellManager
{
    public class RootCertificateHandler
    {
        public bool IsInit => _rootCert != null;
        private readonly X509Certificate2 _rootCert;

        public RootCertificateHandler (string signedFilePath)
        {
            try {
                // Extract signer certificate
                X509Certificate2 signer = new X509Certificate2(X509Certificate.CreateFromSignedFile(signedFilePath));

                // Build chain to get all certificates, but ignore trust
                X509Chain chain = new X509Chain();
                chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                chain.ChainPolicy.VerificationFlags =
                    X509VerificationFlags.IgnoreWrongUsage |
                    X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown |
                    X509VerificationFlags.IgnoreEndRevocationUnknown |
                    X509VerificationFlags.IgnoreRootRevocationUnknown |
                    X509VerificationFlags.IgnoreNotTimeValid;

                chain.Build(signer);

                // Take the **last certificate in the chain**, which should be the root
                _rootCert = chain.ChainElements[chain.ChainElements.Count - 1].Certificate;
            }
            catch {
                Console.WriteLine("Failed to initialize RootCertificateHandler.");
                _rootCert = null;
            }
        }

        public bool IsInstalled ()
        {
            if (!IsInit) {
                Console.WriteLine("Can't check install state, RootCertificateHandler failed to Init");
                return false;
            }

            try {
                X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                store.Open(OpenFlags.ReadOnly);

                bool exists = store.Certificates
                                   .OfType<X509Certificate2>()
                                   .Any(c => c.Thumbprint == _rootCert.Thumbprint);

                store.Close();
                return exists;
            }
            catch (Exception ex) {
                return false;
            }
        }

        public void Install ()
        {
            if (!IsInit) {
                Console.WriteLine("Can't install, RootCertificateHandler failed to Init");
                return;
            }

            if (IsInstalled()) {
                Console.WriteLine("Root CA already installed.");
                return;
            }

            try {
                X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                store.Open(OpenFlags.ReadWrite);
                store.Add(_rootCert);
                store.Close();

                Console.WriteLine("Root CA installed successfully.");
            }
            catch (Exception ex) {

            }
        }

        public void Uninstall ()
        {
            if (!IsInit) {
                Console.WriteLine("Can't uninstall, RootCertificateHandler failed to Init");
                return;
            }

            try {
                X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                store.Open(OpenFlags.ReadWrite);

                var toRemove = store.Certificates
                                    .OfType<X509Certificate2>()
                                    .Where(c => c.Thumbprint == _rootCert.Thumbprint)
                                    .ToList();

                foreach (var cert in toRemove) {
                    store.Remove(cert);
                }

                store.Close();
                Console.WriteLine("Root CA uninstalled successfully.");
            }
            catch (Exception ex) {

            }
        }

        public static bool FileHasSignature (string filePath, bool validateChain = false)
        {
            try {
                X509Certificate2 signer = new X509Certificate2(X509Certificate.CreateFromSignedFile(filePath));
                if (!validateChain && signer != null) return true;

                X509Chain chain = new X509Chain();
                chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                chain.ChainPolicy.VerificationFlags = X509VerificationFlags.IgnoreWrongUsage;

                return chain.Build(signer);
            }
            catch (Exception ex) {
                return false;
            }
        }

    }
}
