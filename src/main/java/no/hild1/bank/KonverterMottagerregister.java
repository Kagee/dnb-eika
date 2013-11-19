package no.hild1.bank;
// no.hild1.bank.KonverterMottagerregister
import java.awt.Container;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileFilter;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.PlainDocument;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class KonverterMottagerregister {
	private Workbook workbook = null;
	private DataFormatter formatter = null;
	private FormulaEvaluator evaluator = null;
	//private String avtalenummer = null;
	private Element rootElement = null;
	private Document doc = null;
	final int INNLAND_KONTONR  		= 0;
	final int INNLAND_LEVNR    		= 1;
	final int INNLAND_NAVN     		= 2;
	final int INNLAND_ADDR1    		= 3;
	final int INNLAND_ADDR2    		= 4;
	final int INNLAND_POSTNUMMER    = 5;
	final int INNLAND_POSTSTED		= 6;
	final int INNLAND_REF			= 8;
	final int INNLAND_MLD			= 9;

	public boolean convertExcelToXML(File inputFile, File outputFile)
			throws Exception {
		if(!inputFile.exists()) {
			throw new IllegalArgumentException("Klarte ikke finne " + inputFile);
		}

		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

		// root elements
		doc = docBuilder.newDocument();
		rootElement = doc.createElement("creditorlist");
		doc.appendChild(rootElement);

		this.openWorkbook(inputFile);
		this.convertToXML();

		// write the content into xml file
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		DOMSource sourceDoc = new DOMSource(doc);
		outputFile.delete();
		StreamResult result = new StreamResult(outputFile);

		transformer.transform(sourceDoc, result);

		System.out.println("Konvertering til TERRA/SDC XML fullført uten feil, KonvertertFraVekselbanken.xml lagret");
		return true;
	}
	private boolean convertToXML() throws InvalidFormatException {
		Sheet sheet = null;

		System.out.println("Starter konvertering til SDC XML");
		int numSheets = this.workbook.getNumberOfSheets();
		boolean foundAtleastOneValid = false;

		for(int i = 0; i < numSheets; i++) {
			sheet = this.workbook.getSheetAt(i);
			String name = sheet.getSheetName();
			if (name.contains("innland")) {
				System.out.println("Fant innland-regneark, prosesserer");
				foundAtleastOneValid = true;
				this.processInnland(sheet);
			} else if (name.contains("utland")) {
				System.out.println("Fant utland-regneark, vet ikke hvordan dette prosesseres");
			} else {
				System.out.println("Fant ukjent regneark: " + name + ", ignorerer");
			}
		}
		if (!foundAtleastOneValid) {
			throw new InvalidFormatException("Fant ikke innland-regneark, ingen output laget. (utland-konvertering støttes ikke atm.)");
		}
		return true;
	}
	private void processInnland(Sheet sheet) throws InvalidFormatException {
		if(sheet.getPhysicalNumberOfRows() > 0) {
			int lastRowNum = 0;
			Row row = null;
			lastRowNum = sheet.getLastRowNum();
			System.out.println("Innland har " + lastRowNum + " rader" );
			row = sheet.getRow(0);
			int lastCellNum = row.getLastCellNum();
			System.out.println("Innland rad 0 har " + lastCellNum + " celler" );
			System.out.println("Kjører tilregnelighetssjekk");

			String KO = text(row, INNLAND_KONTONR);
			String LE = text(row, INNLAND_LEVNR);
			String NA = text(row, INNLAND_NAVN);
			String A1 = text(row, INNLAND_ADDR1);
			String A2 = text(row, INNLAND_ADDR2);
			String NR = text(row, INNLAND_POSTNUMMER);
			String ST = text(row, INNLAND_POSTSTED);
			if (KO.equals("Kontonr") && LE.equals("Lev.nr") && NA.equals("Navn") && A1.equals("Adresse 1")
					&& A2.equals("Adresse 2") && NR.equals("Postnr.") && ST.equals("Poststed")) {
				System.out.println("Første rad ser OK ut, fortsetter");
				for(int j = 1; j <= lastRowNum; j++) {
					System.out.println("Prosesserer rad "  + j + " av " + lastRowNum);
					row = sheet.getRow(j);
					this.rowToXML(row, j);
				}
			} else {
				throw new InvalidFormatException("Kjenner ikke igjen første rad\n" +
						"Skulle vært (1), fant (2)\n(1) 'Kontonr' 'Lev.nr' 'Navn' 'Adresse 1' 'Adresse 2' 'Postnr.' 'Poststed'\n'" +
						"" + KO + "' '" + LE + "' '" + NA + "' '" + A1 + "' '" + A2 + "' '" + NR + "' '" + ST + "'");
			}
		}
	}
	private String text(Row row, int pos) {
		Cell cell = row.getCell(pos);
		if (cell != null) {
			String s = null;
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (pos == INNLAND_POSTNUMMER) {
					s = String.format("%04.0f", cell.getNumericCellValue());
				} else if (pos == INNLAND_KONTONR) {
					s = String.format("%011.0f", cell.getNumericCellValue());
				} else {
					s = String.format("%1.0f", cell.getNumericCellValue());
				}
				break;

			case Cell.CELL_TYPE_FORMULA:
				s = this.formatter.formatCellValue(cell, this.evaluator);
				break;
			default:
				s = cell.getStringCellValue();
			}
			return s.trim();
		} else {
			return "";
		}
	}
	private void rowToXML(Row row, int j) {
		String KO = text(row, INNLAND_KONTONR).trim();
		if (!KO.equals("")) {

			System.out.print("Fant mottager: ");

			String LE = text(row, INNLAND_LEVNR).trim();
			String NA = text(row, INNLAND_NAVN).trim();
			if(NA.equals("")) { NA = "Uten navn"; }

			System.out.println(NA + " (" + LE + "): " + KO);

			String A1 = text(row, INNLAND_ADDR1).trim();
			String A2 = text(row, INNLAND_ADDR2).trim();
			String NR = text(row, INNLAND_POSTNUMMER).trim();
			String ST = text(row, INNLAND_POSTSTED);
			String REF = text(row, INNLAND_REF);
			String MLD = text(row, INNLAND_MLD);
			Element creditor = doc.createElement("creditor");

			Element agreementid = emptyElement("agreementid");
			//doc.createElement("agreementid");
			//agreementid.appendChild();
			//agreementid.appendChild(doc.createTextNode(this.avtalenummer));
			creditor.appendChild(agreementid);

			Element name = doc.createElement("name");
			// Siden vi ikke vet hvorleverandørnummer skal lagres,
			// så lagrer vi det som en del av navnet
			if (!LE.equals("")) {
				NA = NA + " (" + LE + ")";
			}
			name.appendChild(doc.createTextNode(NA));
			creditor.appendChild(name);

			Element address = doc.createElement("address");
			addAddressLines(A1, A2, NR, ST, address);
			creditor.appendChild(address);

			Element phonenumberlist = doc.createElement("phonenumberlist");
			creditor.appendChild(phonenumberlist);

			Element emaillist = doc.createElement("emaillist");
			creditor.appendChild(emaillist);

			Element paymentlist = doc.createElement("paymentlist");

			addPaymentRecord(KO, NA, REF, MLD, paymentlist);

			creditor.appendChild(paymentlist);

			rootElement.appendChild(creditor);
		} else {
			System.out.println("WARNING: Fant intet kontonummer, ignorerer rad " + j);
		}
	}
	private Element textElement(String name, String content) {
		Element foo = doc.createElement(name);
		foo.appendChild(doc.createTextNode(content));
		return foo;
	}

	private Element emptyElement(String name) {
		return doc.createElement(name);
	}

	private void addPaymentRecord(String KO, String NA, String REF, String MLD, Element paymentlist) {
		Element payment = doc.createElement("payment");
		REF = REF.trim();
		MLD = MLD.trim();
		KO = KO.replace(".","");
		KO = KO.replace(" ","");
		if(REF.equals("")) { REF ="Ingen beskrivelse"; } 
		payment.appendChild(textElement("description", REF));
		payment.appendChild(textElement("creditorpaymentportlettype", "KID"));
		payment.appendChild(textElement("toresourceid", KO));
		payment.appendChild(textElement("toresourcetype", "KS-KONTO"));
		payment.appendChild(textElement("currencytype", "NOK"));
		payment.appendChild(textElement("receipttype", "0"));
		String tekstPaaEgenKontoutskrift = KO + " - " + NA;
		// Maxlen 25
		if(tekstPaaEgenKontoutskrift.length() > 25) {
			tekstPaaEgenKontoutskrift = tekstPaaEgenKontoutskrift.substring(0,25);
		}
		payment.appendChild(textElement("debitoradvistext", tekstPaaEgenKontoutskrift));
		payment.appendChild(emptyElement("creditoradvistext"));

		Element senderaddress = doc.createElement("senderaddress");
		for(int i = 1; i <= 5; i++) {
			senderaddress.appendChild(emptyElement("addressline")); 
		}
		payment.appendChild(senderaddress);

		Element paymentlongadvis = doc.createElement("paymentlongadvis");
		paymentlongadvis.appendChild(textElement("rowlength", "35"));
		paymentlongadvis.appendChild(textElement("numberofrows", "12"));
		if(!MLD.equals("")) {
			for(String line : MLD.split("\\r?\\n")) {	
				paymentlongadvis.appendChild(textElement("advis", line));
			}
		}
		payment.appendChild(paymentlongadvis);

		paymentlist.appendChild(payment);
	}

	private void addAddressLines(String addr1, String addr2, String postnr, String poststed, Element address) {
		int adressLines = 0;
		if(!addr1.equals("")) {
			adressLines++;
			address.appendChild(textElement("addressline", addr1));
		}    
		if(!addr2.equals("")) {
			adressLines++;
			address.appendChild(textElement("addressline", addr2));
		}
		if(!(postnr.equals("") && poststed.equals(""))) {
			adressLines++;
			address.appendChild(textElement("addressline", postnr + " " + poststed));
		}
		while (adressLines <= 4) {
			adressLines++;
			address.appendChild(emptyElement("addressline"));
		}
	}

	private void openWorkbook(File file) throws FileNotFoundException, IOException, InvalidFormatException {
		FileInputStream fis = null;
		try {
			System.out.println("Åpner arbeidsbok [" + file.getName() + "]");
			fis = new FileInputStream(file);
			this.workbook = WorkbookFactory.create(fis);
			this.evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
			this.formatter = new DataFormatter(true);
		}
		finally {
			if(fis != null) {
				fis.close();
			}
		}
	}
	JFileChooser fc;
	Container panel;
	JLabel inputFileLabel;
	JButton runButton;
//	JTextField avtalenummerField;
	boolean file = false;
	boolean nummer = true;
	JFrame guiFrame;
	TeePrintStream outStream;
	public KonverterMottagerregister()
	{
		guiFrame = new JFrame();
		panel = guiFrame.getContentPane();
		//make sure the program exits when the frame closes
		guiFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		guiFrame.setTitle("Konverter Mottagerregister (Vekselbanken / Terra)");
		guiFrame.setSize(450,250);

		//This will center the JFrame in the middle of the screen
		guiFrame.setLocationRelativeTo(null);
		panel.setLayout(new GridBagLayout());
		GridBagConstraints c = new GridBagConstraints();

		//JLabel avtalenummerLabel = new JLabel("Avtalenummer (15 tegn):");

		fc = new JFileChooser(System.getProperty("user.dir"));
		fc.addChoosableFileFilter(new XSLXFilter());
		fc.setFileFilter(new XSLXFilter());
		fc.setAcceptAllFileFilterUsed(false);

		//avtalenummerField = new JTextField(15);

		runButton = new JButton( "Konverter");
		c.insets = new Insets(3,3,3,3);
		c.fill = GridBagConstraints.HORIZONTAL;
		c.gridwidth = 2;
		c.gridx = 0;
		c.gridy = 0;
	//	panel.add(avtalenummerLabel, c);

		c.gridx = 2; // one right
	//	panel.add(avtalenummerField, c);
		//JTextFieldLimit limit = new JTextFieldLimit(15);
		//limit.setDocumentFilter(new DocumentInputFilter());

		/*avtalenummerField.setDocument(limit);
		avtalenummerField.addKeyListener(new KeyListener() {
			@Override
			public void keyPressed(KeyEvent arg0) {

			}

			@Override
			public void keyReleased(KeyEvent e) {
				nummer = (avtalenummerField.getText().length() == 15);
				if(nummer) {
					avtalenummer = avtalenummerField.getText();
				}
				//System.out.println("Nummer er " + nummer);
				runButton.setEnabled((nummer && file));
				//System.out.println("Runbutton is " + runButton.isEnabled());
			}

			@Override
			public void keyTyped(KeyEvent e) {

			}
		});
		
		*/
		JButton velgFilButton = new JButton( "Velg fil");
		velgFilButton.setMinimumSize(new Dimension(250,10));
		inputFileLabel = new JLabel("Ingen fil valgt");

		c.gridx = 0; // back left
		c.gridwidth = 4;
		c.gridy++; // one down
		c.gridy++; // one down
		panel.add(inputFileLabel, c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(velgFilButton, c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		panel.add(new JLabel(), c);
		c.gridy++; // one down
		runButton.setEnabled(false);
		panel.add(runButton, c);
		c.gridwidth = 4;
		c.gridy++; // one down
		panel.add(runButton, c);
		
		JMenuBar menuBar = new JMenuBar();
		
		JMenu  hjelpMenu = new JMenu ("Hjelp");
		JMenuItem  hjelpMenuItem = new JMenuItem ("Hjelp");
		hjelpMenuItem.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				JOptionPane.showMessageDialog(null, "1. Velg Excel-filen som skal importeres\n" +
						"2. Klikk Konverter\n(Konverter vil først aktiveres når trinn 1 er gjort)" , "Hjelp", JOptionPane.INFORMATION_MESSAGE);	
			}
		});
		
		JMenuItem  omMenuItem = new JMenuItem ("Om");
		omMenuItem.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(null, "En kopi av Apache License 2.0-lisensen er vedlagt dette programmet.\n\n" +
						"Dette programmet ble skrevet av Anders Einar Hilden (c) 2012 for Terra Servicesenter AS\n\n" +
						"Apache POI\nCopyright 2009 The Apache Software Foundation\n\n" +
						"This product includes software developed by\nThe Apache Software Foundation (http://www.apache.org/).\n\n" +
						"This product contains the DOM4J library (http://www.dom4j.org).\nCopyright 2001-2005 (C) MetaStuff, Ltd. All Rights Reserved.\n\n" +
						"This product contains parts that were originally based on software from BEA.\nCopyright (c) 2000-2003, BEA Systems, <http://www.bea.com/>.\n\n" +
						"This product contains W3C XML Schema documents. Copyright 2001-2003 (c)\nWorld Wide Web Consortium (Massachusetts Institute of Technology, European\nResearch Consortium for Informatics and Mathematics, Keio University)\n\n" +
						"This product contains the Piccolo XML Parser for Java\n(http://piccolo.sourceforge.net/). Copyright 2002 Yuval Oren.\n\n" +
						"This product contains the chunks_parse_cmds.tbl file from the vsdump program.\nCopyright (C) 2006-2007 Valek Filippov (frob@df.ru)" , "Om", JOptionPane.INFORMATION_MESSAGE);
			}
		});
		hjelpMenu.add(hjelpMenuItem);
		hjelpMenu.add(omMenuItem);
		menuBar.add(Box.createHorizontalGlue());
		menuBar.add(hjelpMenu);
		guiFrame.setJMenuBar(menuBar);
		
		
		runButton.addActionListener(new ActionListener()
		{
			@Override
			public void actionPerformed(ActionEvent event)
			{
				File inputFile = new File((fc.getSelectedFile()).getAbsolutePath());
				File logFile  = new File(inputFile.getParent() + File.separator + "run.log");
				File outputFile = new File(inputFile.getParent() + File.separator + "KonvertertFraVekselbanken.xml");
				logFile.delete();
				PrintStream printStream;
				try {
					printStream = new PrintStream(new FileOutputStream(logFile));
					outStream = new TeePrintStream(System.out, printStream);
					System.setOut(outStream);
					if(convertExcelToXML(inputFile, outputFile)) {
						JOptionPane.showMessageDialog(null, "Konverteringen er fullført.\nFilen som skal brukes i nettbanken finner du her:\n\n" + outputFile.getAbsolutePath() + "\n\nDette vinduet vil lukke seg når du klikker OK" , "Ferdig", JOptionPane.INFORMATION_MESSAGE);
						guiFrame.dispose();
					}

				} catch (Exception e) {
					System.out.print(e);
					e.printStackTrace(System.out);
					JOptionPane.showMessageDialog(null, "Noe feil skjedde. Se " + logFile.getAbsolutePath() + " for detaljer", "Ops!", JOptionPane.ERROR_MESSAGE);
					guiFrame.dispose();
				}

			}
		});

		velgFilButton.addActionListener(new ActionListener()
		{
			@Override
			public void actionPerformed(ActionEvent event)
			{
				int returnVal = fc.showDialog(panel, "Konverter denne filen");
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					File inputFile = fc.getSelectedFile();
					inputFileLabel.setText(inputFile.getAbsolutePath());
					file = true;
					runButton.setEnabled((nummer && file));
					// repack to resize
					guiFrame.pack();
				} else {
					file = false;
					runButton.setEnabled((nummer && file));
				}

			}
		});

		//make sure the JFrame is visible
		guiFrame.setMinimumSize(new Dimension(300, 200));
		//guiFrame.pack();
		guiFrame.setVisible(true);
	}

	public static void main(String[] args) throws Exception {

		try {
			// Set System L&F
			UIManager.setLookAndFeel(
					UIManager.getSystemLookAndFeelClassName());
		} 
		catch (UnsupportedLookAndFeelException e) {
			// handle exception
		}
		catch (ClassNotFoundException e) {
			// handle exception
		}
		catch (InstantiationException e) {
			// handle exception
		}
		catch (IllegalAccessException e) {
			// handle exception
		}

		try {
			new KonverterMottagerregister();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			throw new Exception(e);
		}
	}
}

/* ImageFilter.java is used by FileChooserDemo2.java. */
class XSLXFilter extends FileFilter {

	public static String getExtension(File f) {
		String ext = null;
		String s = f.getName();
		int i = s.lastIndexOf('.');

		if (i > 0 &&  i < s.length() - 1) {
			ext = s.substring(i+1).toLowerCase();
		}
		return ext;
	}

	//Accept all directories and all gif, jpg, tiff, or png files.
	public boolean accept(File f) {
		if (f.isDirectory()) {
			return true;
		}

		String extension = getExtension(f);
		if (extension != null) {
			if (extension.equals("xlsx")) {
				return true;
			} else {
				return false;
			}
		}
		return false;
	}

	//The description of this filter
	public String getDescription() {
		return "Microsoft XSLX-regneark";
	}
}

class JTextFieldLimit extends PlainDocument {
	/**
	 * 
	 */
	private static final long serialVersionUID = 7173959251764451072L;
	private int limit;

	JTextFieldLimit(int limit) {
		super();
		this.limit = limit;
	}

	public void insertString( int offset, String  str, AttributeSet attr ) throws BadLocationException {
		if (str == null) return;
		System.out.println(str);
		
		if ((getLength() + str.length()) > limit ) {
			JOptionPane.showMessageDialog(null, "Avtalenummer kan ikke være lenger enn 15 tegn"
					, "Avtalenummer: ", JOptionPane.ERROR_MESSAGE);
			return;
		}
		if(!str.matches("^[0-9]+$")) {
			JOptionPane.showMessageDialog(null, "Avtalenummer bare inneholde tall (0123456789)"
					, "Avtalenummer: ", JOptionPane.ERROR_MESSAGE);
			return;
		}

		super.insertString(offset, str, attr);
	}
}

class TeePrintStream extends PrintStream {
	private final PrintStream second;
	// http://stackoverflow.com/questions/1994255/how-to-write-console-output-to-a-txt-file
	public TeePrintStream(OutputStream main, PrintStream second) {
		super(main);
		this.second = second;
	}

	/**
	 * Closes the main stream.
	 * The second stream is just flushed but <b>not</b> closed.
	 * @see java.io.PrintStream#close()
	 */
	@Override
	public void close() {
		// just for documentation
		super.close();
	}
	@Override
	public void flush() {
		super.flush();
		second.flush();
	}
	@Override
	public void write(byte[] buf, int off, int len) {
		super.write(buf, off, len);
		second.write(buf, off, len);
	}
	@Override
	public void write(int b) {
		super.write(b);
		second.write(b);
	}
	@Override
	public void write(byte[] b) throws IOException {
		super.write(b);
		second.write(b);
	}
}

