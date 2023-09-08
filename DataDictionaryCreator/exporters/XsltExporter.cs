using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace DataDictionaryCreator
{
	/// <summary>
	///     Description of XsltExporter.
	/// </summary>
	public class XsltExporter : XmlExporter
    {
        private readonly FileInfo transformFileInfo;

        public XsltExporter(string transformFileName)
        {
            transformFileInfo = new FileInfo(Application.StartupPath + "/xsl/" + transformFileName);
            if (!transformFileInfo.Exists)
                throw new FileNotFoundException("Unable to find file converter: " + transformFileInfo.FullName);
        }

        protected override void SaveTo(MemoryStream stream, string fileName)
        {
            var xslTran = new XslCompiledTransform();
            xslTran.Load(transformFileInfo.FullName);
            var writer = XmlWriter.Create(fileName, xslTran.OutputSettings);
            xslTran.Transform(new XPathDocument(stream), null, writer);
            writer.Close();
        }
    }
}