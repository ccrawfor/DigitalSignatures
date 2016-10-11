using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Packaging;
using Microsoft.Office.Interop.Word;
using System.IO;


using System.Diagnostics;

namespace DigSigOOXml
{
    public class DigSigWeb
    {

        //should be delegated to the DigSigUtility 
        private string convertInchToPoints(string val)
        {
            string ret = null;
            decimal calc = 0;
            char[] totrim = {'i', 'n' };
            string interim = val.TrimEnd(totrim);
            calc = Convert.ToDecimal(interim) * 72;
            ret = Convert.ToString(Convert.ToInt32(calc));
            return ret;


        }

        public List<string> GetSignatureCommands(string doc) {

            //string filename = null;
            List<string> cmds = new List<string>();

            Package wordPackage = null;
            byte[] filebytes = Convert.FromBase64String(doc);
            using (MemoryStream ms = new MemoryStream(filebytes)) {

                wordPackage = Package.Open(ms);
            
            List<DigSigUtility> commands = new List<DigSigUtility> ();
            string docHeight = null;
            string docWidth = null;
            using (var document = WordprocessingDocument.Open(wordPackage))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                Body body = docPart.Document.Body;
                IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> result = body.Descendants();

                //get document height/width
                IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> pgSize = body.Descendants().Where(e => e.LocalName == "pgSz");
                foreach (DocumentFormat.OpenXml.OpenXmlElement e in pgSize)
                {
                    
                    IEnumerable<DocumentFormat.OpenXml.OpenXmlAttribute> k = e.GetAttributes().Where(a => a.LocalName == "w");
                    foreach (DocumentFormat.OpenXml.OpenXmlAttribute lr in k)
                    {
                        docWidth = lr.Value;
                    }
                    IEnumerable<DocumentFormat.OpenXml.OpenXmlAttribute> l = e.GetAttributes().Where(a => a.LocalName == "h");
                    foreach (DocumentFormat.OpenXml.OpenXmlAttribute lr in l)
                    {
                        docHeight = lr.Value;
                    }

                }

                //look for signature fields
                IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> r = body.Descendants().Where(e => e.LocalName == "shape" && e.OuterXml.Contains("SIG"));

                foreach (DocumentFormat.OpenXml.OpenXmlElement e in r)
                {
                    string name = null;
                    string label = null;
                    string page = null;
                   //pull the parent's styles
                    IEnumerable<DocumentFormat.OpenXml.OpenXmlAttribute> al = e.GetAttributes().Where(a => a.LocalName == "alt");
                    foreach (DocumentFormat.OpenXml.OpenXmlAttribute lr in al)
                    {
                        //add checks
                        string[] s = lr.Value.Split(':');
                        if (s.Length == 4)
                        {
                            name = s[1];
                            label = s[2];
                            page = s[3];

                        }


                    }
                    IEnumerable<DocumentFormat.OpenXml.OpenXmlAttribute> l = e.GetAttributes().Where( a => a.LocalName == "style");
                    //should be only 1 entry per signature field
                    foreach (DocumentFormat.OpenXml.OpenXmlAttribute lr in l)
                   {
                    
                        string[] s = lr.Value.Split(';');
                        char[] totrim = { 'p', 't', 'i', 'n' };
                        string left = null;
                        string top = null;
                        string height = null;
                        string width = null;
                        int iswitch = 0;

                        List<string> lstyles = s.ToList();
                        foreach (string v in lstyles)
                        {
                            bool calc = false;
                            string c = v.Split(':').ElementAt(1);
                            if (c.Substring((c.Length - 2)).Contains("in"))
                            {
                                calc = true;
                            }
                           
                            if (v.Split(':').ElementAt(0).Contains("left")) {
                                if (calc)
                                {
                                    left = convertInchToPoints(c);
                                }
                                else
                                {
                                    //calculation points
                                    left = v.Split(':').ElementAt(1).TrimEnd(totrim);
                                }
                                
                                iswitch++;
                            }
                            if (v.Split(':').ElementAt(0).Contains("width"))
                            {
                                if (calc)
                                {
                                    width = convertInchToPoints(c);
                                }
                                else
                                {
                                    //calculation points
                                    width = v.Split(':').ElementAt(1).TrimEnd(totrim);
                                }

                                iswitch++;

                            }
                            if (v.Split(':').ElementAt(0).Contains("height"))
                            {
                                if (calc)
                                {
                                    height = convertInchToPoints(c);
                                }
                                else
                                {
                                    //calculation points
                                    height = v.Split(':').ElementAt(1).TrimEnd(totrim);
                                }
                               
                                iswitch++;
                            }
                            if (v.Split(':').ElementAt(0).Contains("top"))
                            {
                                if (calc)
                                {
                                    top = convertInchToPoints(c);
                                }
                                else
                                {
                                    //calculation points
                                    top = v.Split(':').ElementAt(1).TrimEnd(totrim);
                                }
                                
                                iswitch++;
                            }

                         

                            
                        }

                        if (iswitch == 4)
                        {
                            DigSigUtility j = new DigSigUtility();
                            j._left = left;
                            j._top = top;
                            j._height = height;
                            j._width = width;
                            j.setDocHeight(docHeight);
                            j.setDocWidth(docWidth);
                            j.name = name;
                            j.label = label;
                            j.page = page;

                            commands.Add(j);
                        }


                   }

                     
                }

                foreach (DigSigUtility sdw in commands)
                {
                    cmds.Add(sdw.command());
          

                }
       
                return cmds;


            }

            }
      
    }

          public string ConvertToPDF(string Document, string DocName) {


              int length = 32768;
              //To-Do: place checks for extension to process only docx or dotx
              string path = "c:\\temp\\" + DocName;
              File.Delete(@path);
              Byte[] bytes = Convert.FromBase64String(Document);
              File.WriteAllBytes(path, bytes);

        
              string c = Path.ChangeExtension(path, ".pdf");
              string conv = "c:\\tmp\\" + Path.GetFileName(c);
              


              File.Delete(@conv);
              Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
              word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
              Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(@path);
              doc.ExportAsFixedFormat(@conv, WdExportFormat.wdExportFormatPDF);
              var docClose = (Microsoft.Office.Interop.Word._Document)doc;
              docClose.Close();
              var wordClose = (Microsoft.Office.Interop.Word._Application)word;
              wordClose.Quit();


              FileStream fs = new FileStream(@conv, FileMode.Open, FileAccess.Read);

              byte[] buffer = new byte[length];
              int read = 0;


              int chunk;
              while ((chunk = fs.Read(buffer, read, buffer.Length - read)) > 0)
              {
                  read += chunk;

                  // If we've reached the end of our buffer, check to see if there's
                  // any more information
                  if (read == buffer.Length)
                  {
                      int nextByte = fs.ReadByte();

                      // End of stream? If so, we're done
                      if (nextByte == -1)
                      {
                          return Convert.ToBase64String(buffer);
                      }

                      // Nope. Resize the buffer, put in the byte we've just
                      // read, and continue
                      byte[] newBuffer = new byte[buffer.Length * 2];
                      Array.Copy(buffer, newBuffer, buffer.Length);
                      newBuffer[read] = (byte)nextByte;
                      buffer = newBuffer;
                      read++;
                  }
              }
              // Buffer is now too big. Shrink it.
              byte[] ret = new byte[read];
              Array.Copy(buffer, ret, read);
              fs.Close();
              return Convert.ToBase64String(ret);

          }

        }

}
