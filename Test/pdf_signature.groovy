import wslite.soap.*
import org.apache.commons.codec.binary.*
import java.io.FileOutputStream
import java.io.ByteArrayOutputStream
import java.io.IOException
import java.io.FileInputStream

 
import com.lowagie.text.*
import com.lowagie.text.pdf.PdfAcroForm
import com.lowagie.text.pdf.PdfBorderArray
import com.lowagie.text.pdf.PdfBorderDictionary
import com.lowagie.text.pdf.PdfReader
import com.lowagie.text.pdf.PdfWriter
import com.lowagie.text.pdf.PdfStamper
import com.lowagie.text.pdf.PdfPCell
import com.lowagie.text.pdf.PdfFormField
import com.lowagie.text.pdf.PdfSignatureAppearance
import com.lowagie.text.pdf.PdfTemplate
import com.lowagie.text.pdf.PdfAnnotation

byte[] input = new File('C://temp/Applicant.docx').readBytes()
String b64 = Base64.encodeBase64String(input)
     

def client = new SOAPClient('http://localhost:59340/WCFService1/Service.svc?wsdl')
def response = client.send(SOAPAction: 'http://tempuri.org/IService/GetSignatureCommands')
  {
    body {
        GetSignatureCommands( xmlns: 'http://tempuri.org/') {
                     doc(b64)
                }
            }
        }

        
def resp = response.GetSignatureCommandsResponse.GetSignatureCommandsResult.text()

//get signature coordinates into map 
def map = [:]
resp.split("&").each {param ->
    def nv = param.split("=")
    map[nv[0]] = nv[1]
    }

//convert document to pdf


client = new SOAPClient('http://localhost:59340/WCFService1/Service.svc?wsdl')
response = client.send(SOAPAction: 'http://tempuri.org/IService/ConvertToPDF')
  {
    body {
        ConvertToPDF( xmlns: 'http://tempuri.org/') {
                     Document(b64)
                     DocName('Applicant.docx')
                }
            }
        }
 
resp = response.ConvertToPDFResponse.ConvertToPDFResult.text()

byte[] pdf = resp.decodeBase64()

def label = map['label']
def page = map['page'].toInteger()

def llx = map['left'].toInteger()
def lly = map['bottom'].toInteger()

def urx = map['width'].toInteger() + map['left'].toInteger()
def ury = lly + map['height'].toInteger()

def coord = [llx, lly, urx, ury]

ByteArrayOutputStream baos = new ByteArrayOutputStream()
PdfReader reader = new PdfReader(pdf);        

PdfStamper stamper = new PdfStamper(reader, baos);

PdfFormField pff = PdfFormField.createSignature(stamper.getWriter());
pff.setWidget(new Rectangle(coord[0], coord[1], coord[2],coord[3]), null);
pff.setFieldName(label);
pff.setFlags(PdfAnnotation.FLAGS_PRINT);
pff.setPage(page);
stamper.addAnnotation(pff, page); 

stamper.close();


new File('c:/tmp/Applicant.pdf').withOutputStream {
        it.write baos.toByteArray()
} 

