package edu.stevens.hvu.document.processing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTFFCheckBox;
import org.docx4j.wml.CTFFData;
import org.docx4j.wml.R;
import org.docx4j.wml.SdtContent;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Text;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;

public class DocProcessor implements Runnable {

	private int id;
	private Thread t;

	public DocProcessor(int id) {
		this.id = id;
		this.t = new Thread(this);
		System.out.println("New thread: " + t.getName());
		t.start();
	}

	public void readDocxFileUsingDocx4j(int i) {
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

			String outputName = "Sample output_" + i;
			// documents4j - Require MS Office Word to be installed in the server!!!
			File exportDocFile = new File(cwd + "\\target\\" + outputName + ".docx");
			wordMLPackage.save(exportDocFile);

			File exportPdfFile = new File(cwd + "\\target\\" + outputName + ".pdf");

			FileInputStream docxInputStream = new FileInputStream(exportDocFile);
			FileOutputStream pdfOutputStream = new FileOutputStream(exportPdfFile);
			IConverter converter = LocalConverter.builder().build();
			converter.convert(docxInputStream).as(DocumentType.DOCX).to(pdfOutputStream).as(DocumentType.PDF).execute();
			pdfOutputStream.close();

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public void run() {
		this.readDocxFileUsingDocx4j(id);
	}

}
