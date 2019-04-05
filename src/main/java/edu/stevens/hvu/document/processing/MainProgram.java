package edu.stevens.hvu.document.processing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import javax.xml.bind.JAXBElement;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.ProtectDocument;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTFFCheckBox;
import org.docx4j.wml.CTFFData;
import org.docx4j.wml.CTFFName;
import org.docx4j.wml.R;
import org.docx4j.wml.STDocProtect;
import org.docx4j.wml.SdtContent;
import org.docx4j.wml.SdtPr;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Text;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;

/**
 * Hello world!
 *
 */
public class MainProgram {
	public static void main(String[] args) {
		ExecutorService exec = Executors.newFixedThreadPool(2);

		for (int i = 1; i < 6; i++) {
			exec.submit(new DocProcessor(i));
			//System.out.println(Thread.currentThread().getName());
			//exec.execute(new DocProcessor(i));
		}
		exec.shutdown();

		System.out.println("DONE");
	}

	public static void readDocxFileUsingApachePOI() {
		try {
			String cwd = System.getProperty("user.dir");
			FileInputStream fis = new FileInputStream(cwd + "\\src\\my-resources\\Sample form.docx");

			@SuppressWarnings("resource")
			XWPFDocument document = new XWPFDocument(fis);

			// read normal text
			List<XWPFParagraph> paragraphs = document.getParagraphs();

			for (XWPFParagraph para : paragraphs) {
				System.out.println(para.getText());
			}

			// read table
			List<XWPFTable> table = document.getTables();

			for (XWPFTable xwpfTable : table) {
				List<XWPFTableRow> row = xwpfTable.getRows();
				for (XWPFTableRow xwpfTableRow : row) {
					List<XWPFTableCell> cell = xwpfTableRow.getTableCells();
					for (XWPFTableCell xwpfTableCell : cell) {
						if (xwpfTableCell != null) {
							System.out.println(xwpfTableCell.getText());
						}
					}
				}
			}

			fis.close();
		} catch (

		Exception e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("rawtypes")
	public static void readDocxFileUsingDocx4j() {
		String cwd = System.getProperty("user.dir");
		File doc = new File(cwd + "\\src\\my-resources\\Sample form.docx");

		try {
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
			MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

			// Name
			List<Object> nameNodes = mainDocumentPart.getJAXBNodesViaXPath("//w:sdt[w:sdtPr[w:tag[@w:val='name']]]",
					false);
			for (Object n : nameNodes) {
				if (n instanceof JAXBElement && ((JAXBElement) n).getValue() instanceof SdtRun) {
					SdtRun run = (SdtRun) ((JAXBElement) n).getValue();
					SdtContent content = run.getSdtContent();
					for (Object c : content.getContent()) {
						if (c instanceof R) {
							for (Object t : ((R) c).getContent()) {
								if (t instanceof JAXBElement) {
									Text name = (Text) ((JAXBElement) t).getValue();
									name.setValue("VU HA DUNG");

								}
							}
						}
					}
				}
			}

			// Title
			List<Object> titleNodes = mainDocumentPart.getJAXBNodesViaXPath("//w:sdt[w:sdtPr[w:tag[@w:val='title']]]",
					false);
			for (Object n : titleNodes) {
				if (n instanceof JAXBElement && ((JAXBElement) n).getValue() instanceof SdtRun) {
					SdtRun run = (SdtRun) ((JAXBElement) n).getValue();
					SdtContent content = run.getSdtContent();
					for (Object c : content.getContent()) {
						if (c instanceof R) {
							for (Object t : ((R) c).getContent()) {
								if (t instanceof JAXBElement) {
									Text name = (Text) ((JAXBElement) t).getValue();
									name.setValue("Chuyen vien cao cap");

								}
							}
						}
					}
				}
			}

			List<Object> checkBoxes = mainDocumentPart.getJAXBNodesViaXPath("//w:ffData[w:name[@w:val = 'Check3']]",
					false);
			for (Object b : checkBoxes) {
				for (Object o : ((CTFFData) b).getNameOrEnabledOrCalcOnExit()) {
					if (o instanceof JAXBElement && ((JAXBElement) o).getValue() instanceof CTFFCheckBox) {
						CTFFCheckBox cb = (CTFFCheckBox) ((JAXBElement) o).getValue();
						BooleanDefaultTrue value = new BooleanDefaultTrue();
						value.setVal(true);
						cb.setChecked(value);
					}
				}
			} // End of for loops CTFFCheckBox

			/*
			 * // Other option ProtectDocument protection = new
			 * ProtectDocument(wordMLPackage);
			 * protection.restrictEditing(STDocProtect.READ_ONLY, "qwerty");
			 * FileOutputStream stream = new FileOutputStream(cwd +
			 * "\\target\\Sample output.pdf"); Docx4J.toPDF(wordMLPackage, stream);
			 * Docx4J.save(wordMLPackage, new java.io.File(cwd +
			 * "\\target\\Sample output.docx"), Docx4J.FLAG_SAVE_ENCRYPTED_AGILE, "1234");
			 * 
			 * // Protect the PDF File file = new File(cwd + "\\target\\Sample output.pdf");
			 * PDDocument document = PDDocument.load(file); // Creating access permission
			 * object AccessPermission ap = new AccessPermission(); // Creating
			 * StandardProtectionPolicy object StandardProtectionPolicy spp = new
			 * StandardProtectionPolicy("1234", "1234", ap); // Setting the length of the
			 * encryption key spp.setEncryptionKeyLength(128); // Setting the access
			 * permissions spp.setPermissions(ap); // Protecting the document
			 * document.protect(spp); // Saving the document document.save(cwd +
			 * "\\target\\Sample output.pdf"); // Closing the document document.close();
			 */

			// documents4j - Require MS Office Word to be installed in the server!!!
			File exportDocFile = new File(cwd + "\\target\\Sample output.docx");
			wordMLPackage.save(exportDocFile);

			File exportPdfFile = new File(cwd + "\\target\\Sample output.pdf");

			FileInputStream docxInputStream = new FileInputStream(exportDocFile);
			FileOutputStream pdfOutputStream = new FileOutputStream(exportPdfFile);
			IConverter converter = LocalConverter.builder().build();
			converter.convert(docxInputStream).as(DocumentType.DOCX).to(pdfOutputStream).as(DocumentType.PDF).execute();
			pdfOutputStream.close();

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}