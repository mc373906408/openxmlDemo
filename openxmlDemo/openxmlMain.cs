using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using System.Linq;
using System.Collections.Generic;

namespace openxmlDemo
{
    class PptDocumentAsMathML
    {

        public enum XSLMODE { OMML2MML, MML2TEX };

        public List<string> findMathParagraph(OpenXmlElementList childs, List<string> mathOuterXml)
        {
            List<string> copyMathOuterXml = mathOuterXml;
            foreach (OpenXmlElement child in childs)
            {
                if(child.LocalName=="m")
                {
                    copyMathOuterXml.Add(child.OuterXml);
                }

                if (child.HasChildren)
                {
                        
                    findMathParagraph(child.ChildElements, copyMathOuterXml);
                }
            }
            return copyMathOuterXml;
        }

        public string xslTransform(string xml, XSLMODE mode)
        {
            string officeML = "";
            /*加载XSL*/
            XslCompiledTransform xsl = new XslCompiledTransform();
            if (mode== XSLMODE.OMML2MML)
            {
                xsl.Load(@"C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL");
            }
            else if(mode== XSLMODE.MML2TEX)
            {
                xsl.Load(@"D:\QtTest\openxmlDemo\openxmlDemo\openxmlDemo\res\xsl\mmltex.xsl");
            }

            /*解析XML*/
            using (TextReader tr = new StringReader(xml))
            {
                using (XmlReader reader = XmlReader.Create(tr))
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        XmlWriterSettings settings = xsl.OutputSettings.Clone();

                        settings.ConformanceLevel = ConformanceLevel.Fragment;
                        settings.OmitXmlDeclaration = true;

                        XmlWriter xw = XmlWriter.Create(ms, settings);

                        xsl.Transform(reader, xw);
                        ms.Seek(0, SeekOrigin.Begin);

                        using (StreamReader sr = new StreamReader(ms, Encoding.UTF8))
                        {
                            officeML = sr.ReadToEnd();
                        }
                    }
                }
            }
            return officeML;
        }

        public void getPptDocumentAsMathML()
        {
            /*这里打开后会自动保存兼容性配置，不支持的会自动转为图片*/
            //OpenSettings settings = new OpenSettings();
            //settings.AutoSave = true;
            //MarkupCompatibilityProcessSettings markupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2019);
            //settings.MarkupCompatibilityProcessSettings = markupCompatibilityProcessSettings;
            /*打开PPT*/
            //using (PresentationDocument presentationDocument = PresentationDocument.Open(@"C:\Users\admin\Downloads\666.pptx", true, settings))
            using (PresentationDocument presentationDocument = PresentationDocument.Open(@"C:\Users\admin\Downloads\666.pptx", false))
            {
                
                /*获取演示部分*/
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                /*PPT页数*/
                int slidesCount = presentationPart.SlideParts.Count();

                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
                
                for (int index = 0; index < slidesCount; index++)
                {
                    /*PPT每页的ID*/
                    string relId = (slideIds[index] as SlideId).RelationshipId;
                    
                    SlidePart slide = (SlidePart)presentationPart.GetPartById(relId);

                    List<string> mathOuterXmlList = new List<string>();
                    /*找到数学公式的xml列表*/
                    mathOuterXmlList = findMathParagraph(slide.Slide.ChildElements, mathOuterXmlList);

                    foreach(string pptXml in mathOuterXmlList)
                    {
                        string mmlXml = xslTransform(pptXml, XSLMODE.OMML2MML);
                        if (mmlXml!="")
                        {
                            string lexXml = xslTransform(mmlXml, XSLMODE.MML2TEX);
                            //Console.Out.WriteLine(lexXml);
                        }
                    }
                }

                /*关闭PPT*/
                presentationDocument.Close();
            }
        }
    }

    class openxmlMain
    {
        static void Main()
        {
            PptDocumentAsMathML m_PptDocumentAsMathML = new PptDocumentAsMathML();
            m_PptDocumentAsMathML.getPptDocumentAsMathML();
            Console.WriteLine("end!!!!!!!!!!!!!");
            Console.ReadKey();
        }
    }
}
