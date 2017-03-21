/**
 * 
 */
package word2pdf;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;



/**
 * @author Joaquin Gayoso-Cabada
 *
 */
public class Principal {

	/**
	 * @param args
	 */
	public static void main(String[] args) {

if (args.length>0)
	{
	System.out.println("Se procesara el archivo con path-> " + args[0]);
	try {
		List<String> Imagenes = generateImages(args[0]);
		System.out.println("Se ha procesado el archivo con path-> " + args[0]);
		try {
			System.out.println("Se insertaran las imagenes en el archivo word con path-> " + args[0]+".docx");
			DocConvert(Imagenes,args[0]);
			System.out.println("Se han insertado las imagenes en el archivo word con path-> " + args[0]+".docx");
		} catch (Exception e) {
			System.err.println("Error Procesando ->" + args[0]);
			e.printStackTrace();
		}
		
		try {
			for (String string : Imagenes) {
				File file=new File(string);
				String path = file.getCanonicalPath();
				File filePath = new File(path);
				if (!filePath.delete())
						System.err.println("Error con " + file.getAbsolutePath());
			}
			System.out.println("Limpiando Imagenes");
		} catch (Exception e) {
			System.err.println("Error Limpiando Imagenes ->" + args[0]);
			e.printStackTrace();
		}
		
	} catch (IOException e) {
		System.err.println("Error Procesando ->" + args[0]);
		e.printStackTrace();
	}
	}
else	
	System.err.println("No hay parametro: java -jar \"filepath\"");

	}

	 public static void DocConvert(List<String> Docs, String orifile) throws InvalidFormatException, IOException {

		 XWPFDocument docx = new XWPFDocument();  
		 XWPFParagraph par = docx.createParagraph();
		 XWPFRun run = par.createRun();
		 for (int i = 0; i < Docs.size(); i++) {
			 String archivos=Docs.get(i);
			 File Farchivo=new File(archivos);
			 
//			 BufferedImage bimg = ImageIO.read(Farchivo);
//			 int width          = bimg.getWidth();
//			 int height         = bimg.getHeight();
			 
			 InputStream pic = new FileInputStream(archivos);
			 run.addPicture(pic, XWPFDocument.PICTURE_TYPE_PNG, Farchivo.getAbsolutePath(),Units.toEMU(425), Units.toEMU(601));
			// par.addPicture(pic, XWPFDocument.PICTURE_TYPE_PNG, null,300,300);
			 System.out.println(i+"/"+Docs.size()+"->"+archivos);
			 pic.close();
		}
		
		 FileOutputStream fos = new FileOutputStream(orifile+".docx");
	        docx.write(fos);
	        fos.flush();
	        fos.close();
	        docx.close();
		 
		 
		 		 
		 /*
		 XWPFDocument docx = new XWPFDocument();  
//		 XWPFParagraph par = docx.createParagraph();  
//		 XWPFRun run = par.createRun();

	        for (int i = 0; i < Docs.size(); i++) {
				 String archivos=Docs.get(i);
				 InputStream pic = new FileInputStream(archivos);
				 byte [] picbytes = IOUtils.toByteArray(pic);
				 docx.addPictureData(picbytes, Document.PICTURE_TYPE_JPEG);
				 System.out.println(i+"/"+Docs.size()+"->"+archivos);
			}
	        
	        FileOutputStream fos = new FileOutputStream(orifile+".docx");
	        docx.write(fos);
	        fos.flush();
	        fos.close();
	        docx.close();
	        */
		 
  }
	
	public static List<String> generateImages(String pdfFilename) throws IOException  {

		List<String> Salida=new ArrayList<String>();
		PDDocument document = PDDocument.load(new File(pdfFilename));
		PDFRenderer pdfRenderer = new PDFRenderer(document);
		for (int page = 0; page < document.getNumberOfPages(); ++page)
		{ 
		    BufferedImage bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.RGB);

		    String actImg=pdfFilename + "-" + (page+1) + ".png";
		    Salida.add(actImg);
		    // suffix in filename will be used as the file format
		    ImageIOUtil.writeImage(bim, actImg, 300);
		    System.out.println(page+"/"+document.getNumberOfPages()+"->"+actImg);
		}
		document.close();   
	    
		return Salida;
	}
	
	

}





