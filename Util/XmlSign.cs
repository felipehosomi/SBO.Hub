using System;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using System.IO;
using System.Deployment.Internal.CodeSigning;

namespace SBO.Hub.Util
{
    public class XmlSign
    {
        public static byte[] SignFile(X509Certificate2 cert, byte[] data)
        {
            try
            {
                ContentInfo content = new ContentInfo(data);
                SignedCms signedCms = new SignedCms(content, false);
                if (VerifySign(data))
                {
                    signedCms.Decode(data);
                }

                CmsSigner signer = new CmsSigner(cert);
                signer.IncludeOption = X509IncludeOption.WholeChain;
                signedCms.ComputeSignature(signer);

                return signedCms.Encode();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao assinar arquivo. A mensagem retornada foi: " + ex.Message);
            }
        }

        public static bool VerifySign(byte[] data)
        {
            try
            {
                SignedCms signed = new SignedCms();
                signed.Decode(data);
            }
            catch
            {
                return false; // Arquivo não assinado
            }
            return true;
        }

        public static byte[] SignFile(string CertFile, string CertPass, byte[] data)
        {
            FileStream fs = new FileStream(CertFile, FileMode.Open);
            byte[] buffer = new byte[fs.Length];
            fs.Read(buffer, 0, buffer.Length);
            X509Certificate2 cert = new X509Certificate2(buffer, CertPass);
            fs.Close();
            fs.Dispose();
            return SignFile(cert, data);
        }

        private static X509Certificate2 FindCertOnStore(int idx)
        {
            X509Store st = new X509Store(StoreLocation.CurrentUser);
            st.Open(OpenFlags.ReadOnly);
            X509Certificate2 ret = st.Certificates[idx];
            st.Close();
            return ret;
        }

        public static void SignXml(XmlDocument Doc, RSA Key)
        {
            // Check arguments.
            if (Doc == null)
                throw new ArgumentException("Doc");
            if (Key == null)
                throw new ArgumentException("Key");

            // Create a SignedXml object.
            SignedXml signedXml = new SignedXml(Doc);

            // Add the key to the SignedXml document.
            signedXml.SigningKey = Key;

            // Create a reference to be signed.
            Reference reference = new Reference();
            reference.Uri = "";

            // Add an enveloped transformation to the reference.
            XmlDsigEnvelopedSignatureTransform env = new XmlDsigEnvelopedSignatureTransform();
            reference.AddTransform(env);

            // Add the reference to the SignedXml object.
            signedXml.AddReference(reference);

            // Compute the signature.
            signedXml.ComputeSignature();

            // Get the XML representation of the signature and save
            // it to an XmlElement object.
            XmlElement xmlDigitalSignature = signedXml.GetXml();

            // Append the element to the XML document.
            Doc.DocumentElement.AppendChild(Doc.ImportNode(xmlDigitalSignature, true));
        }

        public XmlDocument SignXml(string xmlString, string signTag, X509Certificate2 certificate)
        {
            try
            {
                CryptoConfig.AddAlgorithm(typeof(RSAPKCS1SHA256SignatureDescription), "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256");

                // checking if there is a certified used on xml sign
                string _xnome = certificate.Subject.ToString();

                string x = certificate.GetKeyAlgorithm().ToString();

                // Create a new XML document.
                XmlDocument doc = new XmlDocument();

                // Format the document to ignore white spaces.
                doc.PreserveWhitespace = false;

                // Load the passed XML file using it’s name.
                try
                {
                    doc.LoadXml(xmlString);

                    var exportedKeyMaterial = certificate.PrivateKey.ToXmlString(true);
                    var key = new RSACryptoServiceProvider(new CspParameters(24 /* PROV_RSA_AES */));
                    key.PersistKeyInCsp = false;
                    key.FromXmlString(exportedKeyMaterial);

                    // cheching the element will be sign
                    int tagQuantity = doc.GetElementsByTagName(signTag).Count;

                    if (tagQuantity == 0)
                    {
                        throw new Exception("A tag de assinatura " + signTag.Trim() + " não existe");
                    }
                    else
                    {
                        if (tagQuantity > 1)
                        {
                            throw new Exception("A tag de assinatura " + signTag.Trim() + " não é unica");
                        }
                        else
                        {
                            try
                            {
                                // Create a SignedXml object.
                                SignedXml signedXml = new SignedXml(doc);

                                // Add the key to the SignedXml document
                                signedXml.SigningKey = key;
                                signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";

                                // Create a reference to be signed
                                Reference reference = new Reference();

                                XmlAttributeCollection tag = doc.GetElementsByTagName(signTag).Item(0).Attributes;
                                foreach (XmlAttribute xmlAttr in tag)
                                {
                                    if (xmlAttr.Name.ToLower() == "id")
                                        reference.Uri = "#" + xmlAttr.InnerText;
                                }

                                if (reference.Uri == null)
                                {
                                    tag = doc.GetElementsByTagName(signTag).Item(0).ChildNodes[0].Attributes;
                                    foreach (XmlAttribute xmlAttr in tag)
                                    {
                                        if (xmlAttr.Name.ToLower() == "id")
                                            reference.Uri = "#" + xmlAttr.InnerText;
                                    }
                                }

                                // Felipe Hosomi - se reference.Uri == null, dá erro na assinatura
                                if (reference.Uri == null)
                                {
                                    reference.Uri = String.Empty;
                                }

                                // Add an enveloped transformation to the reference.
                                reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());
                                reference.AddTransform(new XmlDsigC14NTransform());
                                reference.DigestMethod = "http://www.w3.org/2001/04/xmlenc#sha256";

                                // Add the reference to the SignedXml object.
                                signedXml.AddReference(reference);

                                // Create a new KeyInfo object
                                KeyInfo keyInfo = new KeyInfo();

                                // load the certificate into a keyinfox509data object
                                // and add it to the keyinfo object.
                                keyInfo.AddClause(new KeyInfoX509Data(certificate));

                                // add the keyinfo object to the signedxml object.
                                signedXml.KeyInfo = keyInfo;
                                signedXml.ComputeSignature();

                                // Get the XML representation of the signature and save
                                // it to an XmlElement object.
                                XmlElement xmlDigitalSignature = signedXml.GetXml();

                                // save element on XML
                                doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, true));

                                // XML document already signed
                                return doc;
                            }
                            catch (Exception e)
                            {
                                throw new Exception("Erro ao assinar documento - " + e.Message);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("Erro ao assinar documento - XML mal formado - " + e.Message);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Problema ao acessar o certificado digital" + e.Message);
            }
        }
    }
}
