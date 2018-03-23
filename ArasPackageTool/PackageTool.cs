using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using PLMIntegration;
using Aras.IOM;
using System.IO;
using System.Diagnostics;

namespace ArasPackageTool
{
    public partial class PackageTool : Form
    {
        public Innovator CoInnovator=null;
        public string CoErrorMessage="";
        public string CoLanguage = "zt";
        List<string> files = new List<string>();
        public PackageTool()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //using (Process exeProcess = Process.Start("Export\\export.exe"))
            //{
            //    exeProcess.WaitForExit();
            //}
            CoLanguage = txtLanguage.Text;
            files.Clear();
            try
            {
                string dir = txtDir.Text;
                GetFilesRecursive(dir);
                ReGetLabelLanguage();
                MessageBox.Show("結束完成");
            }
            catch(Exception ex)
            {
                MessageBox.Show("錯誤"+ex.ToString());
            }
           
            
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {

            Login login = new Login();

            PLMIntegration.Login ArasLogin = new PLMIntegration.Login();
            ArasLogin.CoLanguage = "zh-tw";
            bool resultLogin = ArasLogin.ShowLogin();
            CoErrorMessage = "";
            if (resultLogin == true)
            {
                if (ArasLogin.getInnovator != null)
                {
                    CoInnovator = ArasLogin.getInnovator;
                }
                else if (ArasLogin.getBtnIsLogin == true)
                {
                    CoErrorMessage = "登入錯誤" + ArasLogin.getErrorMessage;
                }
            }
            else
            {
                if (ArasLogin.getErrorMessage != "")
                    CoErrorMessage = "登入錯誤" + ArasLogin.getErrorMessage;
            }
            if (CoErrorMessage != "")
            {
                MessageBox.Show(CoErrorMessage);
            }
        }
        #region"=== Private ==="
        private void ReGetLabelLanguage()
        {
            foreach (string fullpath in files)
            {
                XDocument _xDoc = XDocument.Load(fullpath);
                Item itm = CoInnovator.newItem();
                itm.loadAML(_xDoc.ToString());

                SearchProperties(_xDoc);
                SearchField(_xDoc);
                SearchItemtype(_xDoc);
                SearchRelationshipType(_xDoc);
                SearchWorkflowMapActivity(_xDoc);
                SearchWorkflowMap(_xDoc);
                SearchWorkflowMapPath(_xDoc);
                SearchLifeCycleMap(_xDoc);
                SearchLifeCycleState(_xDoc);
                string prettyXML = PrintXML(_xDoc.FirstNode.ToString());
                File.WriteAllText(fullpath, prettyXML);
                //_xDoc.Save(fullpath);
            }
        }
        
        private void SearchProperties(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Property"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Property", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                       //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                        }
                    }
                }
            }
        }
        private void SearchField(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Field"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Field", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute("xml:lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                        }
                    }
                }
            }
        }
        private void SearchItemtype(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "ItemType"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("ItemType", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage+ ",label_plural_"+CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                            resLabel = (from _xml in resElem.Descendants(i18n + "label_plural")
                                            //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                        select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                        }
                    }
                }
            }
        }
        private void SearchRelationshipType(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "RelationshipType"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("RelationshipType", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage );
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                           
                        }
                    }
                }
            }
        }
        private void SearchWorkflowMapActivity(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Activity Template"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Activity Template", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }

                        }
                    }
                }
            }
        }
        private void SearchWorkflowMapPath(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Workflow Map Path"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Workflow Map Path", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }

                        }
                    }
                }
            }
        }
        private void SearchWorkflowMap(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Workflow Map"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Workflow Map", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }

                        }
                    }
                }
            }
        }
        private void SearchLifeCycleMap(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Life Cycle Map"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Life Cycle Map", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage+ ",description_"+CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                            resLabel = (from _xml in resElem.Descendants(i18n + "description")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }
                        }
                    }
                }
            }
        }
        private void SearchLifeCycleState(XDocument itm)
        {
            List<XElement> xElemProperty = (from _xml in itm.Descendants("Item")
                                            where (string)_xml.Attribute("type") == "Life Cycle State"
                                            select _xml).ToList<XElement>();
            if (xElemProperty.Count() > 0)
            {
                for (int i = 0; i < xElemProperty.Count(); i++)
                {
                    string xElemId = (string)xElemProperty[i].Attribute("id");
                    if (!string.IsNullOrEmpty(xElemId))
                    {
                        Item xElemItem = CoInnovator.newItem("Life Cycle State", "get");
                        xElemItem.setAttribute("where", "id='" + xElemId + "'");
                        xElemItem.setAttribute("select", "label,label_" + CoLanguage);
                        xElemItem = xElemItem.apply();
                        if (!xElemItem.isError())
                        {
                            XDocument resElem = XDocument.Parse(xElemItem.node.OuterXml);
                            XNamespace i18n = @"http://www.aras.com/I18N";
                            List<XElement> resLabel = (from _xml in resElem.Descendants(i18n + "label")
                                                           //where (string)_xml.Attribute(i18n + "lang") == CoLanguage
                                                       select _xml).ToList<XElement>();
                            for (int j = 0; j < resLabel.Count(); j++)
                            {
                                XElement findLabel = xElemProperty.Find(e => e.Value == resLabel[j].Value);
                                if (findLabel == null)
                                {
                                    xElemProperty[i].Add(resLabel[j]);
                                }
                            }

                        }
                    }
                }
            }
        }

        private void GetFilesRecursive(string sDir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    GetFilesRecursive(d);
                }
                foreach (var file in Directory.GetFiles(sDir))
                {
                    DoAction(file);
                }
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        private void DoAction(string filepath)
        {
            if (File.Exists(filepath))
            {
                if (Path.GetExtension(filepath) == ".xml")
                    files.Add(filepath);
            }

        }
        public static String PrintXML(String XML)
        {
            String Result = "";

            MemoryStream mStream = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(mStream, Encoding.Unicode);
            XmlDocument document = new XmlDocument();

            try
            {
                // Load the XmlDocument with the XML.
                document.LoadXml(XML);

                writer.Formatting = Formatting.Indented;

                // Write the XML into a formatting XmlTextWriter
                document.WriteContentTo(writer);
                writer.Flush();
                mStream.Flush();

                // Have to rewind the MemoryStream in order to read
                // its contents.
                mStream.Position = 0;

                // Read MemoryStream contents into a StreamReader.
                StreamReader sReader = new StreamReader(mStream);

                // Extract the text from the StreamReader.
                String FormattedXML = sReader.ReadToEnd();

                Result = FormattedXML;
            }
            catch (XmlException)
            {
            }

            mStream.Close();
            writer.Close();

            return Result;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

            folderBrowserDialog1.ShowDialog();

            txtDir.Text = folderBrowserDialog1.SelectedPath;
        }
    }
}
