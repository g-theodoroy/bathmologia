package bathmologia_java;

import java.awt.Color;
import java.awt.Cursor;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.Random;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JSeparator;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.record.RecordInputStream.LeftoverDataException;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.WorkbookUtil;
import org.javatuples.Pair;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;

public class bathmologia_java extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private static String mydir; // για την επιλογή των αρχείων - φακέλων
	private static String workdir; // ο τρέχων φάκελος εργασίας

	private JPanel contentPane;
	private JTextField textField;
	private JTextField textField_1;
	private JTextField textField_2;
	private JTextField textField_3;
	private JTextField textField_4;
	private JTextField textField_5;
	private JTextField textField_6;
	private JTextField textField_7;
	private JTextField textField_8;
	private JTextField textField_9;
	private JTextField textField_10;
	private JTextField textField_11;
	private ProgressWindow pw;
	private JTextField textField_12;
	private JLabel label_13;
	private JTextField textField_13;
	private JButton button_13;
	private JButton button_8;
	private JButton button_9;
	private JTextField textField_14;
	private JTextField textField_15;
	// ΓΙΑ ΤΟ ΔΙΑΒΑΣΜΑ ΤΩΝ ΚΕΦΑΛΙΔΩΝ ΣΤΗΛΩΝ ΑΠΟ XML
	private NodeList epikefalides1;
	private NodeList epikefalides2;
	private NodeList epikefalides3;
	private NodeList epikefalides4;
	private NodeList epikefalides5;
	private NodeList logFile;
	private NodeList configFile;
	private Document doc;
	private boolean esperino = false;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					bathmologia_java frame = new bathmologia_java();
					frame.setVisible(true);
					// μεταφέρω την εστίαση στο "Περί"
					frame.button_8.requestFocusInWindow();

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public bathmologia_java() {
		// φτιάχνω log file .bathmologia.log
		// και στέλνω εκεί την έξοδο των λαθών και του συστήματος
		try {
			InputStream in = getClass().getResourceAsStream("epikefalides.xml");
			DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
			doc = docBuilder.parse(in);

			// normalize text representation
			doc.getDocumentElement().normalize();
			
			logFile = doc.getElementsByTagName("logFile");

			// βρίσκω το workdir
			workdir = new File(".").getAbsolutePath().substring(0, new File(".").getAbsolutePath().length() - 2);
			// δημιουργία του bathmologia.log
			PrintStream pst = new PrintStream(workdir + File.separator + logFile.item(0).getTextContent());
			// στέλνω τα μυνήματα στο log
			System.setOut(pst);
			// στέλνω τα λάθη στο log
			System.setErr(pst);
			// τυπώνω την ημνία ώρα
			Date date = new Date();
			System.out.println("ΕΤΡΕΞΕ ΤΕΛΕΥΤΑΙΑ ΦΟΡΑ : " + date.toString());
			// τυπώνω το workdir
			System.out.println("ΦΑΚΕΛΟΣ ΕΡΓΑΣΙΑΣ ΓΙΑ CONF - LOG : " + workdir);
		} catch (FileNotFoundException e2) {
			e2.printStackTrace();
		} catch (ParserConfigurationException e1) {
			e1.printStackTrace();
		} catch (SAXException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		/**
		 * Create the frame.
		 */
		setType(Type.NORMAL);
		setTitle("ΕΡΓΑΛΕΙΟ ΚΑΤΑΧΩΡΙΣΗΣ ΒΑΘΜΟΛΟΓΙΩΝ");

		// ΤΡΟΠΟΠΟΙΩ ΤΗΝ ΕΞΟΔΟ ΜΕ ΤΟ ΚΟΥΜΠΙ Χ ΝΑ ΚΑΝΕΙ ΚΛΙΚ ΣΤΟ ΚΟΥΜΠΙ ΕΞΟΔΟΣ
		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent we) {
				button_9.doClick();
			}
		});

		setBounds(100, 100, 600, 590);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 10, 5, 10));
		setContentPane(contentPane);
		GridBagLayout gbl_contentPane = new GridBagLayout();
		gbl_contentPane.columnWidths = new int[] { 0, 0, 10, 0 };
		gbl_contentPane.rowHeights = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
				0, 0, 0 };
		gbl_contentPane.columnWeights = new double[] { 0.0, 1.0, 1.0, 0.0 };
		gbl_contentPane.rowWeights = new double[] { 0.0, 1.0, 1.0, 1.0, 1.0, 0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0,
				1.0, 1.0, 0.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 1.0, 1.0, 1.0, Double.MIN_VALUE, 0.0 };
		contentPane.setLayout(gbl_contentPane);

		// Τύπος σχολείου label combobox checkbox
		JLabel label_14 = new JLabel("Τύπος Σχολείου");
		GridBagConstraints gbc_label_14 = new GridBagConstraints();
		gbc_label_14.weightx = 5.0;
		gbc_label_14.anchor = GridBagConstraints.WEST;
		gbc_label_14.insets = new Insets(10, 0, 5, 5);
		gbc_label_14.gridx = 0;
		gbc_label_14.gridy = 0;
		contentPane.add(label_14, gbc_label_14);

		final JComboBox<String> comboBox_2 = new JComboBox<String>();
		comboBox_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// αν η επιλογή είναι ΗΜΕΡΗΣΙΟ 3 τάξεις κρύβω την Δ τάξη
				if (comboBox_2.getSelectedIndex() == 0) {
					label_13.setVisible(false);
					textField_13.setVisible(false);
					button_13.setVisible(false);
					esperino = false;
				}
				// αν η επιλογή είναι ΕΣΠΕΡΙΝΟ 4 τάξεις εμφανίζω την Δ τάξη
				if (comboBox_2.getSelectedIndex() == 1) {
					label_13.setVisible(true);
					textField_13.setVisible(true);
					button_13.setVisible(true);
					esperino = true;
				}
			}
		});
		comboBox_2.setModel(
				new DefaultComboBoxModel<String>(new String[] { "ΗΜΕΡΗΣΙΟ - 3 ΤΑΞΕΙΣ", "ΕΣΠΕΡΙΝΟ  - 4 ΤΑΞΕΙΣ" }));
		comboBox_2.setMaximumRowCount(2);
		GridBagConstraints gbc_comboBox_2 = new GridBagConstraints();
		gbc_comboBox_2.weightx = 95.0;
		gbc_comboBox_2.gridwidth = 2;
		gbc_comboBox_2.insets = new Insets(10, 0, 5, 5);
		gbc_comboBox_2.fill = GridBagConstraints.HORIZONTAL;
		gbc_comboBox_2.gridx = 1;
		gbc_comboBox_2.gridy = 0;
		contentPane.add(comboBox_2, gbc_comboBox_2);

		final JCheckBox chckbxNewCheckBox = new JCheckBox("");
		chckbxNewCheckBox.setToolTipText("Όχι καθυστέρηση");
		GridBagConstraints gbc_chckbxNewCheckBox = new GridBagConstraints();
		gbc_chckbxNewCheckBox.insets = new Insets(10, 0, 5, 0);
		gbc_chckbxNewCheckBox.gridx = 3;
		gbc_chckbxNewCheckBox.gridy = 0;
		contentPane.add(chckbxNewCheckBox, gbc_chckbxNewCheckBox);

		// Γραμμή
		JSeparator separator = new JSeparator();
		GridBagConstraints gbc_separator = new GridBagConstraints();
		gbc_separator.gridwidth = 4;
		gbc_separator.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator.insets = new Insets(2, 0, 5, 0);
		gbc_separator.gridx = 0;
		gbc_separator.gridy = 1;
		contentPane.add(separator, gbc_separator);

		// Αρχείο Α τάξης label textbox button
		JLabel label = new JLabel("Αρχείο Α Τάξης");
		GridBagConstraints gbc_label = new GridBagConstraints();
		gbc_label.weightx = 5.0;
		gbc_label.insets = new Insets(0, 0, 5, 5);
		gbc_label.anchor = GridBagConstraints.WEST;
		gbc_label.gridx = 0;
		gbc_label.gridy = 2;
		contentPane.add(label, gbc_label);

		textField = new JTextField();
		GridBagConstraints gbc_textField = new GridBagConstraints();
		gbc_textField.weightx = 95.0;
		gbc_textField.gridwidth = 2;
		gbc_textField.insets = new Insets(0, 0, 5, 5);
		gbc_textField.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField.gridx = 1;
		gbc_textField.gridy = 2;
		contentPane.add(textField, gbc_textField);
		textField.setColumns(10);

		JButton button = new JButton("...");
		button.setToolTipText("Επιλέξτε το xls της Α τάξης");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// επιλογή αρχείου xls
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία xls", "xls");
				openFile.setFileFilter(filter);
				// αν υπάρχει ο φάκελος εργασίας στέλνω εκεί το dialogbox
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					// ενημέρωση του textbox με το όνομα του αρχείου
					textField.setText(openFile.getSelectedFile().getAbsolutePath());
					// ενημέρωση mydir με το path του αρχείου
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}
			}
		});
		button.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button = new GridBagConstraints();
		gbc_button.anchor = GridBagConstraints.NORTH;
		gbc_button.insets = new Insets(0, 0, 5, 0);
		gbc_button.gridx = 3;
		gbc_button.gridy = 2;
		contentPane.add(button, gbc_button);

		// Αρχείο B τάξης label textbox button
		JLabel label_1 = new JLabel("Αρχείο Β Τάξης");
		GridBagConstraints gbc_label_1 = new GridBagConstraints();
		gbc_label_1.anchor = GridBagConstraints.WEST;
		gbc_label_1.insets = new Insets(0, 0, 5, 5);
		gbc_label_1.gridx = 0;
		gbc_label_1.gridy = 3;
		contentPane.add(label_1, gbc_label_1);

		textField_1 = new JTextField();
		textField_1.setColumns(10);
		GridBagConstraints gbc_textField_1 = new GridBagConstraints();
		gbc_textField_1.gridwidth = 2;
		gbc_textField_1.insets = new Insets(0, 0, 5, 5);
		gbc_textField_1.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_1.gridx = 1;
		gbc_textField_1.gridy = 3;
		contentPane.add(textField_1, gbc_textField_1);

		JButton button_1 = new JButton("...");
		button_1.setToolTipText("Επιλέξτε το xls της Β τάξης");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία xls", "xls");
				openFile.setFileFilter(filter);
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_1.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}

			}
		});
		button_1.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_1 = new GridBagConstraints();
		gbc_button_1.insets = new Insets(0, 0, 5, 0);
		gbc_button_1.gridx = 3;
		gbc_button_1.gridy = 3;
		contentPane.add(button_1, gbc_button_1);

		// Αρχείο Γ τάξης label textbox button
		JLabel label_2 = new JLabel("Αρχείο Γ Τάξης");
		GridBagConstraints gbc_label_2 = new GridBagConstraints();
		gbc_label_2.anchor = GridBagConstraints.WEST;
		gbc_label_2.insets = new Insets(0, 0, 5, 5);
		gbc_label_2.gridx = 0;
		gbc_label_2.gridy = 4;
		contentPane.add(label_2, gbc_label_2);

		textField_2 = new JTextField();
		textField_2.setColumns(10);
		GridBagConstraints gbc_textField_2 = new GridBagConstraints();
		gbc_textField_2.gridwidth = 2;
		gbc_textField_2.insets = new Insets(0, 0, 5, 5);
		gbc_textField_2.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_2.gridx = 1;
		gbc_textField_2.gridy = 4;
		contentPane.add(textField_2, gbc_textField_2);

		JButton button_2 = new JButton("...");
		button_2.setToolTipText("Επιλέξτε το xls της Γ τάξης");
		button_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία xls", "xls");
				openFile.setFileFilter(filter);
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_2.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}
			}
		});
		button_2.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_2 = new GridBagConstraints();
		gbc_button_2.insets = new Insets(0, 0, 5, 0);
		gbc_button_2.gridx = 3;
		gbc_button_2.gridy = 4;
		contentPane.add(button_2, gbc_button_2);

		// Αρχείο Δ τάξης label textbox button
		label_13 = new JLabel("Αρχείο Δ Τάξης");
		GridBagConstraints gbc_label_13 = new GridBagConstraints();
		gbc_label_13.anchor = GridBagConstraints.WEST;
		gbc_label_13.insets = new Insets(0, 0, 5, 5);
		gbc_label_13.gridx = 0;
		gbc_label_13.gridy = 5;
		contentPane.add(label_13, gbc_label_13);
		label_13.setVisible(false);

		textField_13 = new JTextField();
		textField_13.setColumns(10);
		GridBagConstraints gbc_textField_13 = new GridBagConstraints();
		gbc_textField_13.weighty = 1.0;
		gbc_textField_13.gridwidth = 2;
		gbc_textField_13.insets = new Insets(0, 0, 5, 5);
		gbc_textField_13.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_13.gridx = 1;
		gbc_textField_13.gridy = 5;
		contentPane.add(textField_13, gbc_textField_13);
		textField_13.setVisible(false);

		button_13 = new JButton("...");
		button_13.setToolTipText("Επιλέξτε το xls της Δ τάξης");
		button_13.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία xls", "xls");
				openFile.setFileFilter(filter);
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_13.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}
			}
		});
		button_13.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_13 = new GridBagConstraints();
		gbc_button_13.insets = new Insets(0, 0, 5, 0);
		gbc_button_13.gridx = 3;
		gbc_button_13.gridy = 5;
		contentPane.add(button_13, gbc_button_13);
		button_13.setVisible(false);

		// Γραμμή
		JSeparator separator_1 = new JSeparator();
		GridBagConstraints gbc_separator_1 = new GridBagConstraints();
		gbc_separator_1.gridwidth = 4;
		gbc_separator_1.insets = new Insets(2, 0, 5, 0);
		gbc_separator_1.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator_1.gridx = 0;
		gbc_separator_1.gridy = 6;
		contentPane.add(separator_1, gbc_separator_1);

		// Αναθέσεις καθηγητών label textbox button
		JLabel label_3 = new JLabel("Αναθέσεις καθηγητών");
		GridBagConstraints gbc_label_3 = new GridBagConstraints();
		gbc_label_3.anchor = GridBagConstraints.WEST;
		gbc_label_3.insets = new Insets(0, 0, 5, 5);
		gbc_label_3.gridx = 0;
		gbc_label_3.gridy = 7;
		contentPane.add(label_3, gbc_label_3);

		textField_3 = new JTextField();
		textField_3.setColumns(10);
		GridBagConstraints gbc_textField_3 = new GridBagConstraints();
		gbc_textField_3.gridwidth = 2;
		gbc_textField_3.insets = new Insets(0, 0, 5, 5);
		gbc_textField_3.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_3.gridx = 1;
		gbc_textField_3.gridy = 7;
		contentPane.add(textField_3, gbc_textField_3);

		JButton button_3 = new JButton("...");
		button_3.setToolTipText("Επιλέξτε το xls με τις Αναθέσεις καθηγητών");
		button_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία xls", "xls");
				openFile.setFileFilter(filter);
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_3.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}

			}
		});
		button_3.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_3 = new GridBagConstraints();
		gbc_button_3.insets = new Insets(0, 0, 5, 0);
		gbc_button_3.gridx = 3;
		gbc_button_3.gridy = 7;
		contentPane.add(button_3, gbc_button_3);

		// Τμήματα μαθητών label textbox button
		JLabel label_4 = new JLabel("Τμήματα μαθητών");
		GridBagConstraints gbc_label_4 = new GridBagConstraints();
		gbc_label_4.anchor = GridBagConstraints.WEST;
		gbc_label_4.insets = new Insets(0, 0, 5, 5);
		gbc_label_4.gridx = 0;
		gbc_label_4.gridy = 8;
		contentPane.add(label_4, gbc_label_4);

		textField_4 = new JTextField();
		textField_4.setColumns(10);
		GridBagConstraints gbc_textField_4 = new GridBagConstraints();
		gbc_textField_4.gridwidth = 2;
		gbc_textField_4.insets = new Insets(0, 0, 5, 5);
		gbc_textField_4.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_4.gridx = 1;
		gbc_textField_4.gridy = 8;
		contentPane.add(textField_4, gbc_textField_4);

		JButton button_4 = new JButton("...");
		button_4.setToolTipText("Επιλέξτε το xls με τα Τμήματα μαθητών");
		button_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία xls", "xls");
				openFile.setFileFilter(filter);
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_4.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}

			}
		});
		button_4.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_4 = new GridBagConstraints();
		gbc_button_4.insets = new Insets(0, 0, 5, 0);
		gbc_button_4.gridx = 3;
		gbc_button_4.gridy = 8;
		contentPane.add(button_4, gbc_button_4);

		// Γραμμή
		JSeparator separator_2 = new JSeparator();
		GridBagConstraints gbc_separator_2 = new GridBagConstraints();
		gbc_separator_2.gridwidth = 4;
		gbc_separator_2.insets = new Insets(2, 0, 5, 0);
		gbc_separator_2.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator_2.gridx = 0;
		gbc_separator_2.gridy = 9;
		contentPane.add(separator_2, gbc_separator_2);

		// Φάκελος Δημιουργίας Αρχείων label textbox button
		JLabel label_5 = new JLabel("Φάκελος Δημιουργίας Αρχείων");
		GridBagConstraints gbc_label_5 = new GridBagConstraints();
		gbc_label_5.anchor = GridBagConstraints.WEST;
		gbc_label_5.insets = new Insets(0, 0, 5, 5);
		gbc_label_5.gridx = 0;
		gbc_label_5.gridy = 10;
		contentPane.add(label_5, gbc_label_5);

		textField_5 = new JTextField();
		textField_5.setColumns(10);
		GridBagConstraints gbc_textField_5 = new GridBagConstraints();
		gbc_textField_5.gridwidth = 2;
		gbc_textField_5.insets = new Insets(0, 0, 5, 5);
		gbc_textField_5.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_5.gridx = 1;
		gbc_textField_5.gridy = 10;
		contentPane.add(textField_5, gbc_textField_5);

		JButton button_5 = new JButton("...");
		button_5.setToolTipText("Επιλέξτε το φάκελο δημιουργίας αρχείων");
		button_5.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή φακέλων");
				openFile.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_5.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}
			}
		});
		button_5.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_5 = new GridBagConstraints();
		gbc_button_5.insets = new Insets(0, 0, 5, 0);
		gbc_button_5.gridx = 3;
		gbc_button_5.gridy = 10;
		contentPane.add(button_5, gbc_button_5);

		// Φάκελος Συλλογής Αρχείων label textbox button
		JLabel label_6 = new JLabel("Φάκελος Συλλογής Αρχείων");
		GridBagConstraints gbc_label_6 = new GridBagConstraints();
		gbc_label_6.anchor = GridBagConstraints.WEST;
		gbc_label_6.insets = new Insets(0, 0, 5, 5);
		gbc_label_6.gridx = 0;
		gbc_label_6.gridy = 11;
		contentPane.add(label_6, gbc_label_6);

		textField_6 = new JTextField();
		textField_6.setColumns(10);
		GridBagConstraints gbc_textField_6 = new GridBagConstraints();
		gbc_textField_6.gridwidth = 2;
		gbc_textField_6.insets = new Insets(0, 0, 5, 5);
		gbc_textField_6.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_6.gridx = 1;
		gbc_textField_6.gridy = 11;
		contentPane.add(textField_6, gbc_textField_6);

		JButton button_6 = new JButton("...");
		button_6.setToolTipText("Επιλέξτε το φάκελο με τα συμπληρωμένα αρχεία");
		button_6.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή φακέλων");
				openFile.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_6.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}
			}
		});
		button_6.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_6 = new GridBagConstraints();
		gbc_button_6.insets = new Insets(0, 0, 5, 0);
		gbc_button_6.gridx = 3;
		gbc_button_6.gridy = 11;
		contentPane.add(button_6, gbc_button_6);

		JSeparator separator_3 = new JSeparator();
		GridBagConstraints gbc_separator_3 = new GridBagConstraints();
		gbc_separator_3.gridwidth = 4;
		gbc_separator_3.insets = new Insets(2, 0, 5, 0);
		gbc_separator_3.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator_3.gridx = 0;
		gbc_separator_3.gridy = 12;
		contentPane.add(separator_3, gbc_separator_3);

		JLabel label_7 = new JLabel("Όνομα σχολείου");
		label_7.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_label_7 = new GridBagConstraints();
		gbc_label_7.insets = new Insets(0, 0, 5, 5);
		gbc_label_7.anchor = GridBagConstraints.WEST;
		gbc_label_7.gridx = 0;
		gbc_label_7.gridy = 13;
		contentPane.add(label_7, gbc_label_7);

		textField_7 = new JTextField();
		textField_7.setToolTipText("Πληκτρολογείστε Όνομα σχολείου");
		textField_7.setColumns(10);
		GridBagConstraints gbc_textField_7 = new GridBagConstraints();
		gbc_textField_7.insets = new Insets(0, 0, 5, 5);
		gbc_textField_7.gridwidth = 2;
		gbc_textField_7.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_7.gridx = 1;
		gbc_textField_7.gridy = 13;
		contentPane.add(textField_7, gbc_textField_7);

		JLabel label_8 = new JLabel("Περίοδος     Σχ. έτος");
		label_8.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_label_8 = new GridBagConstraints();
		gbc_label_8.anchor = GridBagConstraints.WEST;
		gbc_label_8.insets = new Insets(0, 0, 5, 5);
		gbc_label_8.gridx = 0;
		gbc_label_8.gridy = 14;
		contentPane.add(label_8, gbc_label_8);

		textField_8 = new JTextField();
		textField_8.setToolTipText("Πληκτρολογείστε Βαθμολογική Περίοδο");
		textField_8.setColumns(10);
		GridBagConstraints gbc_textField_8 = new GridBagConstraints();
		gbc_textField_8.insets = new Insets(0, 0, 5, 5);
		gbc_textField_8.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_8.gridx = 1;
		gbc_textField_8.gridy = 14;
		contentPane.add(textField_8, gbc_textField_8);

		textField_14 = new JTextField();
		textField_14.setToolTipText("Πληκτρολογείστε Σχολικό έτος");
		textField_14.setColumns(10);
		GridBagConstraints gbc_textField_14 = new GridBagConstraints();
		gbc_textField_14.insets = new Insets(0, 0, 5, 5);
		gbc_textField_14.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_14.gridx = 2;
		gbc_textField_14.gridy = 14;
		contentPane.add(textField_14, gbc_textField_14);

		JLabel label_12 = new JLabel("Πόλη            Ημερομηνία");
		label_12.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_label_12 = new GridBagConstraints();
		gbc_label_12.weighty = 1.0;
		gbc_label_12.weightx = 1.0;
		gbc_label_12.anchor = GridBagConstraints.WEST;
		gbc_label_12.insets = new Insets(0, 0, 5, 5);
		gbc_label_12.gridx = 0;
		gbc_label_12.gridy = 15;
		contentPane.add(label_12, gbc_label_12);

		textField_12 = new JTextField();
		textField_12.setToolTipText("Πληκτρολογείστε Πόλη");
		textField_12.setColumns(10);
		GridBagConstraints gbc_textField_12 = new GridBagConstraints();
		gbc_textField_12.weightx = 1.0;
		gbc_textField_12.insets = new Insets(0, 0, 5, 5);
		gbc_textField_12.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_12.gridx = 1;
		gbc_textField_12.gridy = 15;
		contentPane.add(textField_12, gbc_textField_12);

		textField_15 = new JTextField();
		textField_15.setToolTipText("Πληκτρολογείστε Ημ/νια παράδοσης");
		textField_15.setColumns(10);
		GridBagConstraints gbc_textField_15 = new GridBagConstraints();
		gbc_textField_15.insets = new Insets(0, 0, 5, 5);
		gbc_textField_15.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_15.gridx = 2;
		gbc_textField_15.gridy = 15;
		contentPane.add(textField_15, gbc_textField_15);

		JSeparator separator_4 = new JSeparator();
		GridBagConstraints gbc_separator_4 = new GridBagConstraints();
		gbc_separator_4.gridwidth = 3;
		gbc_separator_4.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator_4.insets = new Insets(2, 0, 5, 5);
		gbc_separator_4.gridx = 0;
		gbc_separator_4.gridy = 16;
		contentPane.add(separator_4, gbc_separator_4);

		JLabel label_9 = new JLabel("Είδος Βαθμού");
		label_9.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_label_9 = new GridBagConstraints();
		gbc_label_9.anchor = GridBagConstraints.WEST;
		gbc_label_9.insets = new Insets(0, 0, 5, 5);
		gbc_label_9.gridx = 0;
		gbc_label_9.gridy = 17;
		contentPane.add(label_9, gbc_label_9);

		final JComboBox<String> comboBox = new JComboBox<String>();
		comboBox.setModel(new DefaultComboBoxModel<String>(new String[] { "ΠΡΟΦΟΡΙΚΟΣ", "ΓΡΑΠΤΟΣ" }));
		comboBox.setMaximumRowCount(2);
		GridBagConstraints gbc_comboBox = new GridBagConstraints();
		gbc_comboBox.fill = GridBagConstraints.HORIZONTAL;
		gbc_comboBox.gridwidth = 2;
		gbc_comboBox.insets = new Insets(0, 0, 5, 5);
		gbc_comboBox.gridx = 1;
		gbc_comboBox.gridy = 17;
		contentPane.add(comboBox, gbc_comboBox);

		JLabel label_10 = new JLabel("Διατήρηση των ήδη καταχωρημένων βαθμών");
		label_10.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_label_10 = new GridBagConstraints();
		gbc_label_10.gridwidth = 2;
		gbc_label_10.anchor = GridBagConstraints.WEST;
		gbc_label_10.insets = new Insets(0, 0, 5, 5);
		gbc_label_10.gridx = 0;
		gbc_label_10.gridy = 18;
		contentPane.add(label_10, gbc_label_10);

		final JComboBox<String> comboBox_1 = new JComboBox<String>();
		comboBox_1.setModel(new DefaultComboBoxModel<String>(new String[] { "ΝΑΙ", "ΟΧΙ" }));
		comboBox_1.setMaximumRowCount(2);
		GridBagConstraints gbc_comboBox_1 = new GridBagConstraints();
		gbc_comboBox_1.fill = GridBagConstraints.HORIZONTAL;
		gbc_comboBox_1.insets = new Insets(0, 0, 5, 5);
		gbc_comboBox_1.gridx = 2;
		gbc_comboBox_1.gridy = 18;
		contentPane.add(comboBox_1, gbc_comboBox_1);

		JButton btnXls = new JButton("Συλλογή Αρχείων  xls   από τους καθηγητές");
		btnXls.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				Thread worker = new Thread() {

					public void run() {

						// ΜΕΤΑΒΛΗΤΕΣ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΩ
						File dir;
						File[] files;
						InputStream inp;
						InputStream inpa;
						InputStream inpb;
						InputStream inpg;
						OutputStream otpa;
						OutputStream otpb;
						OutputStream otpg;
						Workbook wb = null;
						Workbook wba;
						Workbook wbb;
						Workbook wbg;

						InputStream inpd;
						OutputStream otpd;
						Workbook wbd = null;

						Sheet sheet;
						String mystr;
						int x;
						int j;
						int c;
						int r;
						String kodsentoni = null;
						String keepvalue = null;
						Map<Pair<String, Pair<Integer, Integer>>, String> data_map = new HashMap<Pair<String, Pair<Integer, Integer>>, String>();
						Row myrow;

						String name;

						JTextArea textArea = new JTextArea();
						textArea.setSize(400, Short.MAX_VALUE);
						textArea.setWrapStyleWord(true);
						textArea.setLineWrap(true);
						textArea.setBackground(new Color(0, 0, 0, 0));

						Cell mycell;
						int sheetcount = 0;
						int wbcount = 0;

						boolean nosleep = false;
						if (chckbxNewCheckBox.isSelected() == true) {
							nosleep = true;
						}

						String uniqueval = ""; // θα τη διαβάσω από το xls
						String uniquerror = ""; // αν υπάρχει ασυμφωνία θα βάλω
												// τα αρχεία

						// αν επέλεξαν αρχείο ελέγχου παίρνω τα περιεχόμενα στη
						// μεταβλητή uniquestr
						String uniquestr = "";
						File uniquefile = new File(textField_11.getText());
						if (uniquefile.exists()) {
							try {
								uniquestr = FileUtils.readFileToString(uniquefile, "UTF-8");
							} catch (IOException e) {
								e.printStackTrace();
							}
						} else {
							if (!"".equals(textField_11.getText().trim())) {
								mystr = "Δυστυχώς δεν μπορώ να ανοίξω το αρχείο ελέγχου \"" + textField_11.getText()
										+ "\"";
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Έλεγχος αρχείων!!!",
										JOptionPane.ERROR_MESSAGE);
								return;
							}
						}

						// έλεγχος αν υπάρχουν τα αρχεία xls των τάξεων και το
						// φάκελο συλλογής αρχείων
						Boolean chk = false;
						mystr = "Για τη συλλογή των βαθμών από τα αρχεία xls των καθηγητών είναι απαραίτητο να επιλέξετε τα αρχεία xls της Α, Β, και Γ Τάξης και τον φάκελο που έχετε αποθηκεύσει τα αρχεία των καθηγητών.\n\nΠαρουσιάστηκαν τα παρακάτω σφάλματα:";
						if (esperino == true)
							mystr = "Για τη συλλογή των βαθμών από τα αρχεία xls των καθηγητών είναι απαραίτητο να επιλέξετε τα αρχεία xls της Α, Β, Γ και Δ Τάξης και τον φάκελο που έχετε αποθηκεύσει τα αρχεία των καθηγητών.\n\nΠαρουσιάστηκαν τα παρακάτω σφάλματα:";

						File chkfile = new File(textField.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Α Τάξης: '" + textField.getText() + "'";
							chk = true;
						}
						chkfile = new File(textField_1.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Β Τάξης: '" + textField_1.getText()
									+ "'";
							chk = true;
						}
						chkfile = new File(textField_2.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Γ Τάξης: '" + textField_2.getText()
									+ "'";
							chk = true;
						}
						if (esperino == true) {
							chkfile = new File(textField_13.getText());
							if (!chkfile.exists()) {
								mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Δ Τάξης: '" + textField_13.getText()
										+ "'";
								chk = true;
							}
						}

						chkfile = new File(textField_6.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο φάκελο Συλλογής Αρχείων: '" + textField_6.getText()
									+ "'";
							chk = true;
						}

						if (chk == true) {
							textArea.setText(mystr);
							JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ελλιπή στοιχεία!!!",
									JOptionPane.ERROR_MESSAGE);
							return;
						}

						// Ενημέρωση του χρήστη για ΣΥΛΛΟΓΗ ΑΡΧΕΙΩΝ ΑΠΟ
						// ΚΑΘΗΓΗΤΕΣ
						mystr = "Θα εκτελεστούν οι ακόλουθες εργασίες:\n\n"
								+ "1. Διάβασμα των αρχείων Βαθμολογιών των καθηγητών από τον φάκελο συλλογής αρχείων.\n"
								+ "2. Καταχώριση των βαθμών στα αρχεία \"Κατάσταση Μαθημάτων ανά Τάξη\" για κάθε τάξη.\n\nΘέλετε να συνεχίστει η διαδικασία;";
						// textArea.setSize(400, Short.MAX_VALUE); // limit =
						// width
						// in
						// pixels, e.g.
						// 500
						System.out.println("\n\n\n" + mystr);
						textArea.setText(mystr);
						int n = JOptionPane.showConfirmDialog(contentPane, textArea, "GΘ@2020: ΕΝΑΡΞΗ ΔΙΑΔΙΚΑΣΙΑΣ",
								JOptionPane.YES_NO_OPTION);
						// αν πατήσει ΟΧΙ σταματάω
						if (n == JOptionPane.NO_OPTION) {
							return;
						}

						// Ανοίγω το progresswindow
						contentPane.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
						pw = new ProgressWindow();
						pw.lb.setText("Συλλογή βαθμολογικών καταστάσεων");

						// ΔΙΑΒΑΖΩ ΟΛΑ ΤΑ ΑΡΧΕΙΑ ΣΕ ΕΝΑ ΠΙΝΑΚΑ
						dir = new File(textField_6.getText());
						files = dir.listFiles(new FilenameFilter() {
							public boolean accept(File dir, String name) {
								return name.toLowerCase().endsWith(".xls");
							}
						});

						// αριθμός αρχείων xls
						wbcount = files.length;

						int numxls = 0;
						int counter = -1;
						// για κάθε αρχείο xls στον φάκελο
						for (File xls : files) {
							counter++;
							try {
								// αν δεν είναι checked το checkbox κοιμήσου
								// 15/1000 του sec
								if (nosleep == false) {
									Thread.sleep(15);
								}
								// ενημέρωση της τιμής της progressbar (20% του
								// συνολικού χρόνου η τρέχουσα εργασία)
								pw.pb.setValue(counter * 20 / wbcount);
							} catch (InterruptedException ex) {
							}
							try {
								// άνοιγμα αρχείου xls
								inp = new FileInputStream(xls);
								wb = WorkbookFactory.create(inp);

								// Αν υπάρχει uniquestring (έχει επιλεγεί αρχείο
								// ελέγχου και το διάβασα νωρίτερα)
								// Περνάω από όλα τα φύλλα και παίρνω από το
								// κελί (r:3,c:7) του xls το αποθηκευμένο
								// uniquestring και συγκρίνω
								// αν δεν υπάρχει το κελί ή οι τιμές δεν
								// συμφωνούν συνεχίζω στο επόμενο φύλλο
								// Αν βρώ έστω και σε ένα φύλλο το κλειδί να
								// συμφωνεί συνεχίζω αλλιώς προσθέτω το όνομα
								// του xls στο μήνυμα
								// λάθους για να ενημερώσω τον χρήστη και
								// continue στο επόμενο xls
								if (!uniquestr.equals("")) {
									boolean chkmatch = false;
									for (x = 0; x < wb.getNumberOfSheets(); x++) {
										Sheet thesheet = wb.getSheetAt(x);
										Row therow = thesheet.getRow(3);
										if (therow == null) {
											continue;
										}
										Cell thecell = therow.getCell(7);
										if (thecell == null) {
											continue;
										}
										if (thecell.getCellType() != Cell.CELL_TYPE_STRING) {
											continue;
										}
										uniqueval = thecell.getStringCellValue();
										if (uniquestr.equals(uniqueval)) {
											chkmatch = true;
										}
									}
									if (chkmatch == false) {
										uniquerror += "\n" + xls.getName();
										continue;
									}
								}
								// αυξάνω τους μετρητές για ενημέρωση του χρήστη
								// στο τέλος
								numxls++;
								sheetcount = sheetcount + wb.getNumberOfSheets();

								// Ανοίγω κάθε φύλλο του xls
								for (x = 0; x < wb.getNumberOfSheets(); x++) {
									sheet = wb.getSheetAt(x);

									// ελέγχω αν υπάρχει το ξεκίνημα το στο κελί
									// (r:3,c:7)
									// αν δεν υπάρχει ούτε η γραμμή ούτε η στήλη
									// ή αν υπάρχει το κελί και τα τέσσερα
									// πρώτα γράμματα δεν είναι "chk_" τότε το
									// Φύλλο αυτό δεν φτιάχτηκε από έμένα
									// αλλά προστέθηκε από το χρήστη και δεν
									// έχει στοιχεία και παραλείπεται
									myrow = sheet.getRow(3);
									if (myrow == null) {
										sheetcount--;
										continue;
									}
									mycell = myrow.getCell(7);
									if (mycell == null) {
										sheetcount--;
										continue;
									}
									if (mycell.getCellType() != Cell.CELL_TYPE_STRING) {
										sheetcount--;
										continue;
									}
									if (!mycell.getStringCellValue().substring(0, 4).equals("chk_")) {
										sheetcount--;
										continue;
									}

									// παίρνω τον κωδικό τάξης του τμήματος
									// (α,β,γ,δ) από το κελί (r:0,c:7)
									kodsentoni = sheet.getRow(0).getCell(7).getStringCellValue();
									// παίρνω τη στήλη που θα μπούν οι βαθμοί
									// στο 187.xls από το κελί (r:1,c:7)
									c = ((int) sheet.getRow(1).getCell(7).getNumericCellValue()) - 1;

									// διαβάζω ταυπόλοιπα δεδομένα και τα βάζω
									// όλα σε πίνακα
									for (j = 5; j < ((int) sheet.getLastRowNum()) - 1; j++) {
										myrow = sheet.getRow(j);
										if (myrow == null) {
											continue;
										}
										if (myrow.getCell(7) == null) {
											continue;
										}
										if (myrow.getCell(7).getCellType() == Cell.CELL_TYPE_NUMERIC) {
											// παίρνω τη γραμμή που θα μπεί ο
											// βαθμός στο 187.xls από το κελί
											// (r:myrow,c:7)
											r = ((int) myrow.getCell(7).getNumericCellValue()) - 1;
											// διαβάζω το βαθμό και ανάλογα με
											// την τιμή (null,numeric,string) αν
											// χρειάζεται
											// μετατρέπω σε string και εισάγω
											// μια γραμμή στον πινακα data_map
											// (kodsentoni,(r,c))=τιμή
											if (myrow.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK) {
												data_map.put(Pair.with(kodsentoni, Pair.with(r, c)), null);
											} else if (myrow.getCell(5).getCellType() == Cell.CELL_TYPE_NUMERIC) {
												data_map.put(Pair.with(kodsentoni, Pair.with(r, c)),
														Double.toString(myrow.getCell(5).getNumericCellValue()));
											} else if (myrow.getCell(5).getCellType() == Cell.CELL_TYPE_STRING) {
												data_map.put(Pair.with(kodsentoni, Pair.with(r, c)),
														myrow.getCell(5).getStringCellValue());
											}
										}

									}

								} // for (x = 0 ; x < wb.getNumberOfSheets()-1;
									// x++)

								wb.close();

							} catch (LeftoverDataException e) {
								// ΕΛΕΓΧΟΣ ΑΝΟΙΓΜΑΤΟΣ ΑΡΧΕΙΟΥ XLS ΤΟY MYSCHOOL
								e.printStackTrace();
								mystr = "Δυστυχώς δεν μπορώ να ανοίξω το αρχέιο xls: ''" + xls
										+ "'' !!! \nΑνοίξτε το με το Excell ή το Libreoffice"
										+ " και Αποθηκεύστε το ως ''Excell 97''." + "\n\nError: " + e.getMessage();
								textArea.setSize(400, Short.MAX_VALUE);
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;
							} catch (EncryptedDocumentException | InvalidFormatException | IOException e1) {
								// ΕΛΕΓΧΟΣ ΑΝΟΙΓΜΑΤΟΣ ΑΡΧΕΙΟΥ XLS ΤΟY MYSCHOOL
								e1.printStackTrace();
								mystr = "Δυστυχώς δεν μπορώ να ανοίξω το αρχέιο xls:\n''" + xls + "'' !!!"
										+ "\n\nError: " + e1.getMessage();
								textArea.setSize(400, Short.MAX_VALUE);
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;
							}

						} // For (File xls:files)

						// εισαγωγή δεδομένων στα 187.xls
						try {
							// ανοίγω τα αρχεία 187.xls
							inpa = new FileInputStream(textField.getText());
							inpb = new FileInputStream(textField_1.getText());
							inpg = new FileInputStream(textField_2.getText());
							wba = WorkbookFactory.create(inpa);
							wbb = WorkbookFactory.create(inpb);
							wbg = WorkbookFactory.create(inpg);
							// αν είναι εσπερινό και της Δ τάξης
							if (esperino == true) {
								inpd = new FileInputStream(textField_13.getText());
								wbd = WorkbookFactory.create(inpd);
							}

							boolean chkreplacevalue = true;

							keepvalue = comboBox_1.getSelectedItem().toString();

							counter = -1;

							// για κάθε γραμμή του πίνακα data_map
							for (Entry<Pair<String, Pair<Integer, Integer>>, String> entry : data_map.entrySet()) {

								counter++;
								try {
									if (nosleep == false) {
										Thread.sleep(1);
									}
									// ενημέρωση της progressbar (αυτή η εργασία
									// είναι το 80%) + 20% από πριν
									pw.pb.setValue(counter * 80 / data_map.size() + 20);
								} catch (InterruptedException ex) {
								}

								// παίρνω το αρχείο 187 (α,β,γ,δ)
								String xlsfile = entry.getKey().getValue0();
								chkreplacevalue = true;
								// γραμμή
								r = entry.getKey().getValue1().getValue0();
								// στήλη
								c = entry.getKey().getValue1().getValue1();

								// Αν είναι της Α τάξης
								if (xlsfile.equals("a")) {

									sheet = wba.getSheetAt(0);
									// επιλέγω τη γραμμή.
									myrow = sheet.getRow(r);
									// Αν δεν υπάρχει τη δημιουργώ
									if (myrow == null)
										myrow = sheet.createRow(r);
									// επιλέγω το κελί στη στήλη.
									mycell = myrow.getCell(c);
									// Αν δεν υπάρχει το δημιουργώ
									if (mycell == null)
										mycell = myrow.createCell(c);
									// Αν δεν είναι κενό ελέγχω αν θα πρέπει να
									// αντικαταστήσω την τιμή
									// Αν το combobox λέει ΝΑΙ κράτα τις τιμές
									// τότε ενημερώνω τη μεταβλητή
									// chkreplacevalue σε false = μην
									// αντικαθιστάς
									if (mycell.getCellType() != Cell.CELL_TYPE_BLANK) {
										if (keepvalue.equals("ΝΑΙ")) {
											chkreplacevalue = false;
										}
									}
									// Αν επιτρέπεται η αντικατάσταση
									if (chkreplacevalue == true) {
										// Αν η νέα τιμή είναι null σβήνω ότι
										// υπάρχει
										if (entry.getValue() == null) {
											mycell.setCellValue("");
										} else {
											// Αλλιώς μετατρέπω τη τιμή πρώτα σε
											// Double με ένα δεκαδικό ψηφίο
											Double mynum = Double.parseDouble(entry.getValue());
											// ΣΤΡΟΓΓΥΛΕΜΑ ΣΤΟ 1 ΔΕΚΑΔΙΚΟ
											mynum = Math.round(mynum * 10) / 10.0d;
											// αν τo δεκαδικό είναι .0 παίρνω
											// μόνο το ακέραιο μέρος
											if ((mynum == Math.floor(mynum)) && !Double.isInfinite(mynum)) {
												mycell.setCellValue(mynum.intValue());
											} else {
												mycell.setCellValue(mynum);
											}
										}

									} // if (chkreplacevalue == true)
								}

								// Αν είναι της Β τάξης
								if (xlsfile.equals("b")) {

									sheet = wbb.getSheetAt(0);
									myrow = sheet.getRow(r);
									if (myrow == null)
										myrow = sheet.createRow(r);
									mycell = myrow.getCell(c);
									if (mycell == null)
										mycell = myrow.createCell(c);
									if (mycell.getCellType() != Cell.CELL_TYPE_BLANK) {
										if (keepvalue.equals("ΝΑΙ")) {
											chkreplacevalue = false;
										}
									}
									if (chkreplacevalue == true) {
										if (entry.getValue() == null) {
											mycell.setCellValue("");
										} else {
											Double mynum = Double.parseDouble(entry.getValue());
											// ΣΤΡΟΓΓΥΛΕΜΑ ΣΤΟ 1 ΔΕΚΑΔΙΚΟ
											mynum = Math.round(mynum * 10) / 10.0d;
											if ((mynum == Math.floor(mynum)) && !Double.isInfinite(mynum)) {
												mycell.setCellValue(mynum.intValue());
											} else {
												mycell.setCellValue(mynum);
											}
										}

									} // if (chkreplacevalue == true)

								}

								// Αν είναι της Γ τάξης
								if (xlsfile.equals("g")) {

									sheet = wbg.getSheetAt(0);
									myrow = sheet.getRow(r);
									if (myrow == null)
										myrow = sheet.createRow(r);
									mycell = myrow.getCell(c);
									if (mycell == null)
										mycell = myrow.createCell(c);
									if (mycell.getCellType() != Cell.CELL_TYPE_BLANK) {
										if (keepvalue.equals("ΝΑΙ")) {
											chkreplacevalue = false;
										}
									}
									if (chkreplacevalue == true) {
										if (entry.getValue() == null) {
											mycell.setCellValue("");
										} else {
											Double mynum = Double.parseDouble(entry.getValue());
											// ΣΤΡΟΓΓΥΛΕΜΑ ΣΤΟ 1 ΔΕΚΑΔΙΚΟ
											mynum = Math.round(mynum * 10) / 10.0d;
											if ((mynum == Math.floor(mynum)) && !Double.isInfinite(mynum)) {
												mycell.setCellValue(mynum.intValue());
											} else {
												mycell.setCellValue(mynum);
											}
										}

									} // if (chkreplacevalue == true)
								}

								// Αν είναι της Δ τάξης (ΜΟΝΟ ΕΣΠΕΡΙΝΑ)
								if (esperino == true) {
									if (xlsfile.equals("d")) {

										sheet = wbd.getSheetAt(0);
										myrow = sheet.getRow(r);
										if (myrow == null)
											myrow = sheet.createRow(r);
										mycell = myrow.getCell(c);
										if (mycell == null)
											mycell = myrow.createCell(c);
										if (mycell.getCellType() != Cell.CELL_TYPE_BLANK) {
											if (keepvalue.equals("ΝΑΙ")) {
												chkreplacevalue = false;
											}
										}
										if (chkreplacevalue == true) {
											if (entry.getValue() == null) {
												mycell.setCellValue("");
											} else {
												Double mynum = Double.parseDouble(entry.getValue());
												// ΣΤΡΟΓΓΥΛΕΜΑ ΣΤΟ 1 ΔΕΚΑΔΙΚΟ
												mynum = Math.round(mynum * 10) / 10.0d;
												if ((mynum == Math.floor(mynum)) && !Double.isInfinite(mynum)) {
													mycell.setCellValue(mynum.intValue());
												} else {
													mycell.setCellValue(mynum);
												}
											}

										} // if (chkreplacevalue == true)
									}
								}

								// System.out.println("file : " + xlsfile + ",
								// row : " +
								// r + ", col : " + c + " Value : " +
								// entry.getValue());

							} // for (Entry<Pair<String, Pair<Integer,
								// Integer>>,
								// String> entry : data_map.entrySet())

							// Απόθήκευση των αρχείων
							// Α τάξη
							name = textField.getText();
							String newname = name.substring(0, name.length() - 4) + "_ΕΝΗΜΕΡΩΜΕΝΟ.xls";
							otpa = new FileOutputStream(newname);
							wba.write(otpa);

							// Β τάξη
							name = textField_1.getText();
							newname = name.substring(0, name.length() - 4) + "_ΕΝΗΜΕΡΩΜΕΝΟ.xls";
							otpb = new FileOutputStream(newname);
							wbb.write(otpb);

							// Γ τάξη
							name = textField_2.getText();
							newname = name.substring(0, name.length() - 4) + "_ΕΝΗΜΕΡΩΜΕΝΟ.xls";
							otpg = new FileOutputStream(newname);
							wbg.write(otpg);

							// Δ τάξη
							if (esperino == true) {
								name = textField_13.getText();
								newname = name.substring(0, name.length() - 4) + "_ΕΝΗΜΕΡΩΜΕΝΟ.xls";
								otpd = new FileOutputStream(newname);
								wbd.write(otpd);
							}

							// Τακτοποίηση
							wba.close();
							wbb.close();
							wbg.close();
							if (esperino == true)
								wbd.close();

							pw.frame.setVisible(false);
							pw = null;
							contentPane.setCursor(null);

							// Ενημέρωση του χρήστη
							mystr = "Εισήχθησαν " + sheetcount + " Βαθμολογικές Καταστάσεις από " + numxls
									+ " αρχεία καθηγητών";
							// αν έγινε έλεγχος με αρχείο ελέγχου και αν
							// παραλείφθκαν αρχεία και ποιά
							if (!uniquerror.equals("")) {
								mystr += "\n\n\nΠΡΟΒΛΗΜΑ ΤΑΥΤΟΠΟΙΗΣΗΣ!!!\n\nΤα παρακάτω αρχεία δεν μπόρεσαν να ταυτοποιηθούν με το κλειδί που επιλέξατε και δεν καταχωρήθηκαν:\n"
										+ uniquerror;
							}
							System.out.println("\n\n\n" + mystr);
							textArea.setText(mystr);
							JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: ΤΕΛΟΣ ΔΙΑΔΙΚΑΣΙΑΣ",
									JOptionPane.INFORMATION_MESSAGE, null);

						} catch (EncryptedDocumentException | InvalidFormatException | IOException e1) {
							// Auto-generated catch block
							e1.printStackTrace();
						}

					}
				};

				worker.start(); // So we don't hold up the dispatch thread.

			}// ΤΕΛΟΣ VOID
		});

		button_8 = new JButton("Περί ...");
		button_8.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String mystr = "Το συγκεκριμμένο \"πρόγραμμα\" δημιουργήθηκε τον Ιανουάριο του 2015 "
						+ "για τις ανάγκες καταχώρiσης βαθμολογίας στο 3ο Γενικό Λύκειο Πάτρας από τον "
						+ "Θεοδώρου Γεώργιο ΠΕ11.\n\nemail: g.theodoroy@gmail.com\n\n"
						+ "Δεν φέρω ευθύνη για τυχόν λάθη που μπορεί να συμβούν από τη χρήση "
						+ "του παρόντος. Κανένα \"πρόγραμμα\" δεν υποκάθιστά την σκέψη και την "
						+ "κρίση του χρήστη!\n\n" + "Rubish in, Rubish out! Καλή δύναμη...";
				JTextArea textArea = new JTextArea();
				textArea.setSize(400, Short.MAX_VALUE); // limit = width in
														// pixels, e.g. 500
				textArea.setWrapStyleWord(true);
				textArea.setLineWrap(true);
				textArea.setBackground(new Color(0, 0, 0, 0));
				textArea.setText(mystr);
				JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Πληροφορίες",
						JOptionPane.INFORMATION_MESSAGE, null);
			}
		});

		JButton btnNewButton = new JButton("Δημιουργία Αρχείωv xls για τους καθηγητές");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				Thread worker = new Thread() {

					public void run() {

						// ΜΕΤΑΒΛΗΤΕΣ ΠΟΥ ΘΑ ΧΡΕΙΑΣΤΩ
						// ------------------------------------------------------------------------
						ArrayList<Integer> stiles = new ArrayList<Integer>();
						int startrow = 0;
						int lessonsrow = 0;
						int firstcol = 0;
						int lastcol = 0;
						String myfile;
						String mystr;
						String myfile2open;
						List<String[]> students = new ArrayList<String[]>();
						String name;
						InputStream inp;
						OutputStream otp;
						Workbook wb;
						Sheet sheet;
						String sheetname;
						ArrayList<ArrayList<String[]>> anatheseis = new ArrayList<ArrayList<String[]>>();
						Map<Pair<String, String>, String> am_map = new HashMap<Pair<String, String>, String>();
						Map<Pair<String, String>, String> lesson_map = new HashMap<Pair<String, String>, String>();
						Map<Pair<String, Pair<Integer, Integer>>, String> data_map = new HashMap<Pair<String, Pair<Integer, Integer>>, String>();

						JTextArea textArea = new JTextArea();
						textArea.setSize(400, Short.MAX_VALUE);
						textArea.setWrapStyleWord(true);
						textArea.setLineWrap(true);
						textArea.setBackground(new Color(0, 0, 0, 0));

						String tmimata;
						String tmima;
						String mathima;
						String tmima_mathima;
						int row2print;
						Row myrow;
						int aa;
						String myref;
						String myformula;
						Cell mycell;
						String kodsentoni = null;
						int sheetcount = 0;
						int wbcount = 0;
						String mypassword = textField_9.getText().trim();
						Date dNow = new Date();
						SimpleDateFormat ft = new SimpleDateFormat("yyyyMMdd_HHmmss");

						String prefix = "";
						String dirprefix = "";
						prefix = textField_10.getText().trim();
						if (!prefix.equals("")) {
							dirprefix = prefix;
							prefix += "_";
						}
						String unique_indentifier = "chk_" + prefix + ft.format(dNow);
						String topos = "";
						topos = textField_12.getText().trim();
						String Hmnia = "";
						Hmnia = textField_15.getText().trim();
						if (Hmnia.equals(""))
							Hmnia = "....../...../........";

						boolean nosleep = false;
						if (chckbxNewCheckBox.isSelected() == true) {
							nosleep = true;
						}

						int is_blank = Cell.CELL_TYPE_BLANK;
						int is_numeric = Cell.CELL_TYPE_NUMERIC;
						int is_string = Cell.CELL_TYPE_STRING;

						// έλεγχος αν υπάρχουν τα αρχεία που έχουν καταχωρηθεί
						Boolean chk = false;
						mystr = "Για τη δημιουργία των αρχείων xls των καθηγητών είναι απαραίτητο να επιλέξετε τα αρχεία xls της Α, Β, και Γ Τάξης, τις Αναθέσεις καθηγητών, τα Τμήματα των Μαθητών και ένα φάκελο δημιουργίας Αρχείων για καθηγητές.\nΕπίσης να πληκτρολογήσετε το Όνομα του Σχολείου και την περίοδο & Σχ. έτος (πχ: Α Τριμηνο 2015-15).\n\nΠαρουσιάστηκαν τα παρακάτω σφάλματα:";
						if (esperino == true)
							mystr = "Για τη δημιουργία των αρχείων xls των καθηγητών είναι απαραίτητο να επιλέξετε τα αρχεία xls της Α, Β, Γ και Δ Τάξης, τις Αναθέσεις καθηγητών, τα Τμήματα των Μαθητών και ένα φάκελο δημιουργίας Αρχείων για καθηγητές.\nΕπίσης να πληκτρολογήσετε το Όνομα του Σχολείου και την περίοδο & Σχ. έτος (πχ: Α Τριμηνο 2015-15).\n\nΠαρουσιάστηκαν τα παρακάτω σφάλματα:";

						File chkfile = new File(textField.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Α Τάξης: '" + textField.getText() + "'";
							chk = true;
						}
						chkfile = new File(textField_1.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Β Τάξης: '" + textField_1.getText()
									+ "'";
							chk = true;
						}
						chkfile = new File(textField_2.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Γ Τάξης: '" + textField_2.getText()
									+ "'";
							chk = true;
						}
						if (esperino == true) {
							chkfile = new File(textField_13.getText());
							if (!chkfile.exists()) {
								mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο xls της Δ Τάξης: '" + textField_13.getText()
										+ "'";
								chk = true;
							}
						}

						chkfile = new File(textField_3.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο Αναθέσεων καθηγητών: '" + textField_3.getText()
									+ "'";
							chk = true;
						}
						chkfile = new File(textField_4.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο αρχείο Τμημάτων μαθητών: '" + textField_4.getText()
									+ "'";
							chk = true;
						}
						chkfile = new File(textField_5.getText());
						if (!chkfile.exists()) {
							mystr = mystr + "\n\n" + "Σφάλμα στο φάκελο Δημιουργίας Αρχείων: '" + textField_5.getText()
									+ "'";
							chk = true;
						}
						if (textField_7.getText().trim().equals("")) {
							mystr = mystr + "\n\n" + "Πληκτρολογείστε το Όνομα του Σχολείου'";
							chk = true;
						}
						if (textField_8.getText().trim().equals("")) {
							mystr = mystr + "\n\n" + "Πληκτρολογείστε τη Βαθμολογική περίοδο & Σχ. έτος.";
							chk = true;
						}
						if (textField_14.getText().trim().equals("")) {
							mystr = mystr + "\n\n" + "Πληκτρολογείστε το Σχoλικό έτος.";
							chk = true;
						}
						if (chk == true) {
							textArea.setText(mystr);
							JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ελλιπή στοιχεία!!!",
									JOptionPane.ERROR_MESSAGE);
							return;
						}

						// ΔΙΑΒΑΣΜΑ ΑΠΟ XML ΤΩΝ ΕΠΙΚΕΦΑΛΙΔΩΝ ΤΩΝ ΣΤΗΛΩΝ ΠΟΥ ΜΕ
						// ΕΝΣΙΑΦΕΡΟΥΝ
						// ΓΙΑ ΝΑ ΒΡΩ ΤΟΝ ΑΡΙΘΜΟ ΣΤΗΛΗΣ ΚΑΙ ΣΕΙΡΑΣ ΤΩΝ ΔΕΔΟΜΕΝΩΝ
						try {
							InputStream in = getClass().getResourceAsStream("epikefalides.xml");
							DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
							DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
							doc = docBuilder.parse(in);

							// normalize text representation
							doc.getDocumentElement().normalize();

							epikefalides1 = doc.getElementsByTagName("ep1");
							epikefalides2 = doc.getElementsByTagName("ep2");
							epikefalides3 = doc.getElementsByTagName("ep3");
							epikefalides4 = doc.getElementsByTagName("ep4");
							epikefalides5 = doc.getElementsByTagName("ep5");

							// ΕΚΤΥΠΩΣΗ ΣΤΟ LOG FILE ΤΩΝ ΚΕΦΑΛΙΔΩΝ ΣΤΗΛΩΝ ΓΙΑ
							// ΕΛΕΓΧΟ
							System.out.println(
									"\n\n\nKεφαλίδες από xml για ευρεση στηλών - γραμμών δεδομένων \nκαι αντίστοιχες στήλες - γραμμές (η αρίθμηση έχει βάση το (0) μηδέν)");

						} catch (SAXParseException err) {
							System.out.println("** Parsing error" + ", line " + err.getLineNumber() + ", uri "
									+ err.getSystemId());
							System.out.println(" " + err.getMessage());

						} catch (SAXException e) {
							Exception x = e.getException();
							((x == null) ? e : x).printStackTrace();

						} catch (Throwable t) {
							t.printStackTrace();
						}

						// ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΩΝ ΓΙΑ ΚΑΘΗΓΗΤΕΣ
						mystr = "Θα εκτελεστούν οι ακόλουθες εργασίες:\n\n"
								+ "1. Διάβασμα του αρχείου με τα στοιχεία των μαθητών.\n"
								+ "2. Διάβασμα του αρχείου με τις Αναθέσεις των καθηγητών.\n"
								+ "3. Διάβασμα των αρχείων \"Κατάσταση Μαθημάτων ανά Τάξη\" για κάθε τάξη.\n\n"
								+ "4. Δημιουργία των αρχείων για τους Καθηγητές.\n\nΘέλετε να συνεχίστει η διαδικασία;";
						textArea.setSize(400, Short.MAX_VALUE); // limit = width
																// in pixels,
																// e.g. 500
						textArea.setText(mystr);
						int n = JOptionPane.showConfirmDialog(contentPane, textArea, "GΘ@2020: ΕΝΑΡΞΗ ΔΙΑΔΙΚΑΣΙΑΣ",
								JOptionPane.OK_CANCEL_OPTION);
						if (n == JOptionPane.CANCEL_OPTION) {
							return;
						}

						// ΑΝΟΙΓΜΑ ΤΟΥ ΑΡΧΕΙΟΥ ΤΜΗΜΑΤΑ ΜΑΘΗΤΩΝ
						// -----------------------------------------------------------------
						myfile = textField_4.getText();
						myfile2open = myfile.replace("\\", "\\\\");
						try {
							inp = new FileInputStream(myfile2open);
							wb = WorkbookFactory.create(inp);
							sheet = wb.getSheetAt(0);

							// ΕΥΡΕΣΗ ΤΩΝ ΣΤΗΛΩΝ ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ
							for (int i = 0; i < sheet.getLastRowNum(); i++) {
								Row row = sheet.getRow(i);
								if (row == null) {
									continue;
								}
								for (int j = 0; j < row.getLastCellNum(); j++) {
									Cell cell = row.getCell(j);
									if (cell == null) {
										continue;
									}
									// Βρίσκω τις στήλες που με ενδιαφέρουν
									// ελέγχοντας τις επικεφαλίδες των στηλών
									// στο xls τμήματα
									// και τις αποθηκεύω σε ένα πίνακα
									if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
										if (cell.getStringCellValue().equals(epikefalides1.item(0).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides2.item(0).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides3.item(0).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides4.item(0).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides5.item(0).getTextContent())) {
											stiles.add(j);
										}
									}
									if (stiles.size() >= 5) {
										break;
									}
								}
								if (stiles.size() >= 5) {
									startrow = i + 1;
									break;
								}
							}
							// ΑΝ ΔΕΝ ΒΡΩ ΟΛΕΣ ΤΙΣ ΣΤΗΛΕΣ
							// ----------------------------------------------------------------
							if (stiles.size() < 5) {
								mystr = "Δυστυχώς υπάρχει πρόβλημα με τα στοιχεία του αρχείου των Τμημάτων του μαθητή. Δεν βρέθηκαν όλες οι απαραίτητες στήλες με τα δεδομένα.\n\nΒεβαιωθείτε ότι έχετε επιλέξει το σωστό αρχείο με τα τμήματα μαθητή.\n\nΑλλιώς επικοινωνείστε με τον προγραμματιστή...";
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;

							}

							if (stiles.size() == 5) {
								System.out.println("");
								System.out.println("arxeio : " + doc.getElementsByTagName("arxeio").item(0)
										.getAttributes().getNamedItem("name").getNodeValue());
								System.out.println("ep1 : " + epikefalides1.item(0).getTextContent() + "\t στήλη : "
										+ stiles.get(0).toString());
								System.out.println("ep2 : " + epikefalides2.item(0).getTextContent() + "\t στήλη : "
										+ stiles.get(1).toString());
								System.out.println("ep3 : " + epikefalides3.item(0).getTextContent() + "\t στήλη : "
										+ stiles.get(2).toString());
								System.out.println("ep4 : " + epikefalides4.item(0).getTextContent() + "\t στήλη : "
										+ stiles.get(3).toString());
								System.out.println("ep5 : " + epikefalides5.item(0).getTextContent() + "\t στήλη : "
										+ stiles.get(4).toString());
								System.out.println("αρχική σειρα δεδομένων : " + startrow);
							}
							// System.out.println(stiles);
							// System.out.println(startrow);
							// System.out.println( "rows =" +
							// sheet.getLastRowNum());

							// ΑΠΟΘΗΚΕΥΣΗ ΤΩΝ ΣΤΟΙΧΕΙΩΝ ΤΜΗΜΑΤΩΝ ΜΑΘΗΤΩΝ ΣΕ
							// ΠΙΝΑΚΑ
							while (startrow <= sheet.getLastRowNum()) {
								if (sheet.getRow(startrow).getCell((int) stiles.get(0)).getCellType() == is_numeric) {
									students.add(new String[] {
											Integer.toString((int) sheet.getRow(startrow).getCell((int) stiles.get(0))
													.getNumericCellValue()),
											sheet.getRow(startrow).getCell((int) stiles.get(1)).getStringCellValue(),
											sheet.getRow(startrow).getCell((int) stiles.get(2)).getStringCellValue(),
											sheet.getRow(startrow).getCell((int) stiles.get(3)).getStringCellValue(),
											sheet.getRow(startrow).getCell((int) stiles.get(4)).getStringCellValue()
													+ "," });
									// System.out.println(Arrays.toString(students.get(students.size()-1)));
								}
								startrow = startrow + 1;
							}

							// ΚΛΕΙΣΙΜΟ ΑΡΧΕΙΟΥ
							wb.close();

							// System.out.println(students.get(0)[1]);
							// System.out.println(students.get(1)[1]);
							// System.out.println(students.get(0)[4]);
							// System.out.println(students.get(1)[4]);

							// ----- ΑΝΟΙΓΜΑ ΤΟΥ ΑΡΧΕΙΟΥ ΑΝΑΘΕΣΕΩΝ
							// --------------------------------------------------------------------
							myfile = textField_3.getText();
							myfile2open = myfile.replace("\\", "\\\\");
							inp = new FileInputStream(myfile2open);
							wb = WorkbookFactory.create(inp);
							sheet = wb.getSheetAt(0);

							// ΕΥΡΕΣΗ ΤΩΝ ΣΤΗΛΩΝ ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ
							stiles.clear();
							startrow = 0;
							for (int i = 0; i < sheet.getLastRowNum(); i++) {
								Row row = sheet.getRow(i);
								if (row == null) {
									// do something with an empty row
									continue;
								}
								for (int j = 0; j < row.getLastCellNum(); j++) {
									Cell cell = row.getCell(j);
									if (cell == null) {
										// do something with an empty cell
										continue;
									}
									// Βρίσκω τις στήλες που με ενδιαφέρουν
									// ελέγχοντας τις επικεφαλίδες των στηλών
									// στο xls αναθέσεις
									// και τις αποθηκεύω σε ένα πίνακα
									if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
										if (cell.getStringCellValue().equals(epikefalides1.item(1).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides2.item(1).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides3.item(1).getTextContent())) {
											stiles.add(j);
										} else if (cell.getStringCellValue()
												.equals(epikefalides4.item(1).getTextContent())) {
											stiles.add(j);
										}
									}
									if (stiles.size() >= 4) {
										break;
									}
								}
								if (stiles.size() >= 4) {
									startrow = i + 1;
									break;
								}
							}
							// ΑΝ ΔΕΝ ΒΡΩ ΟΛΕΣ ΤΙΣ ΣΤΗΛΕΣ
							// ----------------------------------------------------------------
							if (stiles.size() < 4) {
								mystr = "Δυστυχώς υπάρχει πρόβλημα με τα στοιχεία του αρχείου των αναθέσεων. Δεν βρέθηκαν όλες οι απαραίτητες στήλες με τα δεδομένα.\n\nΒεβαιωθείτε ότι έχετε επιλέξει το σωστό αρχείο με τις αναθέσεις των καθηγητών.\n\nΑλλιώς επικοινωνείστε με τον προγραμματιστή...";
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;

							}

							if (stiles.size() == 4) {
								System.out.println("");
								System.out.println("arxeio : " + doc.getElementsByTagName("arxeio").item(1)
										.getAttributes().getNamedItem("name").getNodeValue());
								System.out.println("ep1 : " + epikefalides1.item(1).getTextContent() + "\t στήλη : "
										+ stiles.get(0).toString());
								System.out.println("ep2 : " + epikefalides2.item(1).getTextContent() + "\t στήλη : "
										+ stiles.get(1).toString());
								System.out.println("ep3 : " + epikefalides3.item(1).getTextContent() + "\t στήλη : "
										+ stiles.get(2).toString());
								System.out.println("ep4 : " + epikefalides4.item(1).getTextContent() + "\t στήλη : "
										+ stiles.get(3).toString());
								System.out.println("αρχική σειρα δεδομένων : " + startrow);
							}
							// System.out.println(stiles);
							// System.out.println(startrow);
							// System.out.println( "rows =" +
							// sheet.getLastRowNum());

							// ΑΠΟΘΗΚΕΥΣΗ ΤΩΝ ΑΝΑΘΕΣΕΩΝ ΣΕ ΠΙΝΑΚΑ
							// ----------------------------------------------------------------------------------------------------------

							name = sheet.getRow(startrow).getCell((int) stiles.get(0)).getStringCellValue() + " "
									+ sheet.getRow(startrow).getCell((int) stiles.get(1)).getStringCellValue();
							while (startrow <= sheet.getLastRowNum()) {
								if (sheet.getRow(startrow).getCell((int) stiles.get(2)).getCellType() != is_blank
										&& sheet.getRow(startrow).getCell((int) stiles.get(3))
												.getCellType() != is_blank) {
									if (sheet.getRow(startrow).getCell((int) stiles.get(0)).getCellType() != is_blank
											&& sheet.getRow(startrow).getCell((int) stiles.get(1))
													.getCellType() != is_blank) {
										anatheseis.add(new ArrayList<String[]>());
										name = sheet.getRow(startrow).getCell((int) stiles.get(0)).getStringCellValue()
												+ " " + sheet.getRow(startrow).getCell((int) stiles.get(1))
														.getStringCellValue();
										anatheseis.get(anatheseis.size() - 1).add(new String[] { name });
									}
									anatheseis.get(anatheseis.size() - 1).add(new String[] {
											sheet.getRow(startrow).getCell((int) stiles.get(2)).getStringCellValue(),
											sheet.getRow(startrow).getCell((int) stiles.get(3)).getStringCellValue() });
								}
								startrow = startrow + 1;
							}

							// ΚΛΕΙΣΙΜΟ ΑΡΧΕΙΟΥ
							wb.close();

							// ΕΠΕΞΕΡΓΑΣΙΑ ΤΩΝ ΣΕΝΤΟΝΙΩΝ
							// ----------------------------------------------------------------------------------
							// -------------Α ΤΑΞΗ-------------------

							myfile = textField.getText();
							myfile2open = myfile.replace("\\", "\\\\");
							inp = new FileInputStream(myfile2open);
							wb = WorkbookFactory.create(inp);
							sheet = wb.getSheetAt(0);

							// ΕΥΡΕΣΗ ΤΩΝ ΣΤΗΛΩΝ ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ
							stiles.clear();
							startrow = 0;
							for (int i = 0; i < sheet.getLastRowNum(); i++) {
								Row row = sheet.getRow(i);
								if (row == null) {
									// do something with an empty row
									continue;
								}
								for (int y = 0; y < row.getLastCellNum(); y++) {
									Cell cell = row.getCell(y);
									if (cell == null) {
										// do something with an empty cell
										continue;
									}
									if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
										// System.out.println(cell.getStringCellValue());
										if (cell.getStringCellValue().equals(epikefalides1.item(2).getTextContent())) {
											stiles.add(y);
										} else if (cell.getStringCellValue()
												.equals(epikefalides2.item(2).getTextContent())) {
											stiles.add(y);
										}
									}
									if (stiles.size() >= 2) {
										firstcol = y + 1;
										break;
									}
								}
								if (stiles.size() >= 2) {
									lessonsrow = i;
									startrow = i + 1;
									break;
								}
							}
							// ΑΝ ΔΕΝ ΒΡΩ ΟΛΕΣ ΤΙΣ ΣΤΗΛΕΣ
							// ----------------------------------------------------------------
							if (stiles.size() < 2) {
								mystr = "Δυστυχώς υπάρχει πρόβλημα με τα στοιχεία του αρχείου Βαθμολογιών της Α τάξης. Δεν βρέθηκαν όλες οι απαραίτητες στήλες με τα δεδομένα.\n\nΒεβαιωθείτε ότι έχετε επιλέξει το σωστό αρχείο με τις βαθμολογίες.\n\nΑλλιώς επικοινωνείστε με τον προγραμματιστή...";
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;
							}

							if (stiles.size() == 2) {
								System.out.println("");
								System.out.println("arxeio : " + doc.getElementsByTagName("arxeio").item(2)
										.getAttributes().getNamedItem("name").getNodeValue());
								System.out.println("ep1 : " + epikefalides1.item(2).getTextContent() + "\t στήλη : "
										+ stiles.get(0).toString());
								System.out.println("ep2 : " + epikefalides2.item(2).getTextContent() + "\t στήλη : "
										+ stiles.get(1).toString());
								System.out.println("αρχική σειρα δεδομένων : " + startrow);
							}
							// System.out.println(stiles);
							// System.out.println(startrow);
							// System.out.println(startcol);
							// System.out.println( "rows =" +
							// sheet.getLastRowNum());

							// βάζω σε πίνακα τον ΑρΜητρώου και το αρχείο "a"
							// για την Α τάξη (AM,file)= a
							int x = startrow;
							while (x <= sheet.getLastRowNum()) {
								if (sheet.getRow(x).getCell((int) stiles.get(0)).getCellType() == is_numeric) {
									am_map.put(Pair.with(Integer.toString(
											(int) sheet.getRow(x).getCell((int) stiles.get(0)).getNumericCellValue()),
											"file"), "a");
									// βάζω σε πίνακα τον ΑρΜητρώου και tη σειρά
									// toy row (AM,"row")= row
									am_map.put(
											Pair.with(Integer.toString((int) sheet.getRow(x)
													.getCell((int) stiles.get(0)).getNumericCellValue()), "row"),
											Integer.toString(x));
									// System.out.println(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
									// "row") + ", " +
									// am_map.get(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
									// "row")));
								}
								x = x + 1;
							}
							// βάζω σε πίνακα τη στήλη κάθε μαθήματος ("α",
							// Μάθημα )= στήλη
							int j = firstcol;
							while (j < sheet.getRow(lessonsrow).getLastCellNum()) {
								if (sheet.getRow(lessonsrow).getCell(j).getCellType() == is_string) {
									lesson_map.put(
											Pair.with("a", sheet.getRow(lessonsrow).getCell(j).getStringCellValue()),
											Integer.toString(j));
									// System.out.println(j + " " +
									// Pair.with("a",
									// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())
									// + ", " +
									// lesson_map.get(Pair.with("a",
									// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())));
									lastcol = j;
								}
								j = j + 1;
							}
							// System.out.println( "lastcol a: " + lastcol);

							// βάζω σε πίνακα τους ήδη υπάρχοντες βαθμούς ("α",
							// (row, col))= βαθμός
							x = startrow;
							while (x <= sheet.getLastRowNum()) {
								for (j = firstcol; j <= lastcol; j++) {
									if (sheet.getRow(x).getCell(j).getCellType() == is_string) {
										data_map.put(Pair.with("a", Pair.with(x, j)),
												sheet.getRow(x).getCell(j).getStringCellValue());
										// System.out.println(Integer.toString(x)
										// + ","
										// + Integer.toString(j) + " " +
										// sheet.getRow(x).getCell(j).getStringCellValue());
									} else if (sheet.getRow(x).getCell(j).getCellType() == is_numeric) {
										data_map.put(Pair.with("a", Pair.with(x, j)),
												Double.toString(sheet.getRow(x).getCell(j).getNumericCellValue()));
									}
								}
								x = x + 1;
							}
							wb.close();

							// -------------B ΤΑΞΗ-------------------

							myfile = textField_1.getText();
							myfile2open = myfile.replace("\\", "\\\\");
							inp = new FileInputStream(myfile2open);
							wb = WorkbookFactory.create(inp);
							sheet = wb.getSheetAt(0);

							// ΕΥΡΕΣΗ ΤΩΝ ΣΤΗΛΩΝ ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ
							stiles.clear();
							startrow = 0;
							for (int i = 0; i < sheet.getLastRowNum(); i++) {
								Row row = sheet.getRow(i);
								if (row == null) {
									// do something with an empty row
									continue;
								}
								for (int y = 0; y < row.getLastCellNum(); y++) {
									Cell cell = row.getCell(y);
									if (cell == null) {
										// do something with an empty cell
										continue;
									}
									if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
										// System.out.println(cell.getStringCellValue());
										if (cell.getStringCellValue().equals(epikefalides1.item(2).getTextContent())) {
											stiles.add(y);
										} else if (cell.getStringCellValue()
												.equals(epikefalides2.item(2).getTextContent())) {
											stiles.add(y);
										}
									}
									if (stiles.size() >= 2) {
										firstcol = y + 1;
										break;
									}
								}
								if (stiles.size() >= 2) {
									lessonsrow = i;
									startrow = i + 1;
									break;
								}
							}
							// ΑΝ ΔΕΝ ΒΡΩ ΟΛΕΣ ΤΙΣ ΣΤΗΛΕΣ
							// ----------------------------------------------------------------
							if (stiles.size() < 2) {
								mystr = "Δυστυχώς υπάρχει πρόβλημα με τα στοιχεία του αρχείου Βαθμολογιών της Β τάξης. Δεν βρέθηκαν όλες οι απαραίτητες στήλες με τα δεδομένα.\n\nΒεβαιωθείτε ότι έχετε επιλέξει το σωστό αρχείο με τις βαθμολογίες.\n\nΑλλιώς επικοινωνείστε με τον προγραμματιστή...";
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;
							}

							if (stiles.size() == 2) {
								System.out.println("");
								System.out.println("arxeio : " + doc.getElementsByTagName("arxeio").item(2)
										.getAttributes().getNamedItem("name").getNodeValue());
								System.out.println("ep1 : " + epikefalides1.item(2).getTextContent() + "\t στήλη : "
										+ stiles.get(0).toString());
								System.out.println("ep2 : " + epikefalides2.item(2).getTextContent() + "\t στήλη : "
										+ stiles.get(1).toString());
								System.out.println("αρχική σειρα δεδομένων : " + startrow);
							}
							// System.out.println(stiles);
							// System.out.println(startrow);
							// System.out.println(startcol);
							// System.out.println( "rows =" +
							// sheet.getLastRowNum());

							
							x = startrow;
							while (x <= sheet.getLastRowNum()) {
								if (sheet.getRow(x).getCell((int) stiles.get(0)).getCellType() == is_numeric) {
									am_map.put(Pair.with(Integer.toString(
											(int) sheet.getRow(x).getCell((int) stiles.get(0)).getNumericCellValue()),
											"file"), "b");
									am_map.put(
											Pair.with(Integer.toString((int) sheet.getRow(x)
													.getCell((int) stiles.get(0)).getNumericCellValue()), "row"),
											Integer.toString(x));
									// System.out.println(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
									// "row") + ", " +
									// am_map.get(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
									// "row")));
								}
								x = x + 1;
							}
							j = firstcol;
							while (j < sheet.getRow(lessonsrow).getLastCellNum()) {
								if (sheet.getRow(lessonsrow).getCell(j).getCellType() == is_string) {
									lesson_map.put(
											Pair.with("b", sheet.getRow(lessonsrow).getCell(j).getStringCellValue()),
											Integer.toString(j));
									// System.out.println(j + " " +
									// Pair.with("a",
									// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())
									// + ", " +
									// lesson_map.get(Pair.with("a",
									// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())));
									lastcol = j;
								}
								j = j + 1;
							}
							// System.out.println( "lastcol b: " + lastcol);

							x = startrow;
							while (x <= sheet.getLastRowNum()) {
								for (j = firstcol; j <= lastcol; j++) {
									if (sheet.getRow(x).getCell(j).getCellType() == is_string) {
										data_map.put(Pair.with("b", Pair.with(x, j)),
												sheet.getRow(x).getCell(j).getStringCellValue());
										// System.out.println(Integer.toString(x)
										// + ","
										// + Integer.toString(j) + " " +
										// sheet.getRow(x).getCell(j).getStringCellValue());
									} else if (sheet.getRow(x).getCell(j).getCellType() == is_numeric) {
										data_map.put(Pair.with("b", Pair.with(x, j)),
												Double.toString(sheet.getRow(x).getCell(j).getNumericCellValue()));
									}
								}
								x = x + 1;
							}
							wb.close();

							// -------------Γ ΤΑΞΗ-------------------

							myfile = textField_2.getText();
							myfile2open = myfile.replace("\\", "\\\\");
							inp = new FileInputStream(myfile2open);
							wb = WorkbookFactory.create(inp);
							sheet = wb.getSheetAt(0);

							// ΕΥΡΕΣΗ ΤΩΝ ΣΤΗΛΩΝ ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ
							stiles.clear();
							startrow = 0;
							for (int i = 0; i < sheet.getLastRowNum(); i++) {
								Row row = sheet.getRow(i);
								if (row == null) {
									// do something with an empty row
									continue;
								}
								for (int y = 0; y < row.getLastCellNum(); y++) {
									Cell cell = row.getCell(y);
									if (cell == null) {
										// do something with an empty cell
										continue;
									}
									if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
										// System.out.println(cell.getStringCellValue());
										if (cell.getStringCellValue().equals(epikefalides1.item(2).getTextContent())) {
											stiles.add(y);
										} else if (cell.getStringCellValue()
												.equals(epikefalides2.item(2).getTextContent())) {
											stiles.add(y);
										}
									}
									if (stiles.size() >= 2) {
										firstcol = y + 1;
										break;
									}
								}
								if (stiles.size() >= 2) {
									lessonsrow = i;
									startrow = i + 1;
									break;
								}
							}
							// ΑΝ ΔΕΝ ΒΡΩ ΟΛΕΣ ΤΙΣ ΣΤΗΛΕΣ
							// ----------------------------------------------------------------
							if (stiles.size() < 2) {
								mystr = "Δυστυχώς υπάρχει πρόβλημα με τα στοιχεία του αρχείου Βαθμολογιών της Γ τάξης. Δεν βρέθηκαν όλες οι απαραίτητες στήλες με τα δεδομένα.\n\nΒεβαιωθείτε ότι έχετε επιλέξει το σωστό αρχείο με τις βαθμολογίες.\n\nΑλλιώς επικοινωνείστε με τον προγραμματιστή...";
								textArea.setText(mystr);
								JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
										JOptionPane.ERROR_MESSAGE, null);
								return;
							}

							if (stiles.size() == 2) {
								System.out.println("");
								System.out.println("arxeio : " + doc.getElementsByTagName("arxeio").item(2)
										.getAttributes().getNamedItem("name").getNodeValue());
								System.out.println("ep1 : " + epikefalides1.item(2).getTextContent() + "\t στήλη : "
										+ stiles.get(0).toString());
								System.out.println("ep2 : " + epikefalides2.item(2).getTextContent() + "\t στήλη : "
										+ stiles.get(1).toString());
								System.out.println("αρχική σειρα δεδομένων : " + startrow);
							}
							// System.out.println(stiles);
							// System.out.println(startrow);
							// System.out.println(startcol);
							// System.out.println( "rows =" +
							// sheet.getLastRowNum());

							
							x = startrow;
							while (x <= sheet.getLastRowNum()) {
								if (sheet.getRow(x).getCell((int) stiles.get(0)).getCellType() == is_numeric) {
									am_map.put(Pair.with(Integer.toString(
											(int) sheet.getRow(x).getCell((int) stiles.get(0)).getNumericCellValue()),
											"file"), "g");
									am_map.put(
											Pair.with(Integer.toString((int) sheet.getRow(x)
													.getCell((int) stiles.get(0)).getNumericCellValue()), "row"),
											Integer.toString(x));
									// System.out.println(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
									// "row") + ", " +
									// am_map.get(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
									// "row")));
								}
								x = x + 1;
							}
							j = firstcol;
							while (j < sheet.getRow(lessonsrow).getLastCellNum()) {
								if (sheet.getRow(lessonsrow).getCell(j).getCellType() == is_string) {
									lesson_map.put(
											Pair.with("g", sheet.getRow(lessonsrow).getCell(j).getStringCellValue()),
											Integer.toString(j));
									// System.out.println(j + " " +
									// Pair.with("a",
									// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())
									// + ", " +
									// lesson_map.get(Pair.with("a",
									// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())));
									lastcol = j;
								}
								j = j + 1;
							}
							// System.out.println( "lastcol g: " + lastcol);

							x = startrow;
							while (x <= sheet.getLastRowNum()) {
								for (j = firstcol; j <= lastcol; j++) {
									if (sheet.getRow(x).getCell(j).getCellType() == is_string) {
										data_map.put(Pair.with("g", Pair.with(x, j)),
												sheet.getRow(x).getCell(j).getStringCellValue());
										// System.out.println(Integer.toString(x)
										// + ","
										// + Integer.toString(j) + " " +
										// sheet.getRow(x).getCell(j).getStringCellValue());
									} else if (sheet.getRow(x).getCell(j).getCellType() == is_numeric) {
										data_map.put(Pair.with("g", Pair.with(x, j)),
												Double.toString(sheet.getRow(x).getCell(j).getNumericCellValue()));
									}
								}
								x = x + 1;
							}
							wb.close();

							// ΑΝ ΤΟ ΣΧΟΛΕΙΟ ΕΙΝΑΙ ΕΣΠΕΡΙΝΟ
							if (esperino == true) {
								// -------------Δ ΤΑΞΗ-------------------

								myfile = textField_13.getText();
								myfile2open = myfile.replace("\\", "\\\\");
								inp = new FileInputStream(myfile2open);
								wb = WorkbookFactory.create(inp);
								sheet = wb.getSheetAt(0);
								
								// ΕΥΡΕΣΗ ΤΩΝ ΣΤΗΛΩΝ ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ
								stiles.clear();
								startrow = 0;
								for (int i = 0; i < sheet.getLastRowNum(); i++) {
									Row row = sheet.getRow(i);
									if (row == null) {
										// do something with an empty row
										continue;
									}
									for (int y = 0; y < row.getLastCellNum(); y++) {
										Cell cell = row.getCell(y);
										if (cell == null) {
											// do something with an empty cell
											continue;
										}
										if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
											// System.out.println(cell.getStringCellValue());
											if (cell.getStringCellValue().equals(epikefalides1.item(2).getTextContent())) {
												stiles.add(y);
											} else if (cell.getStringCellValue()
													.equals(epikefalides2.item(2).getTextContent())) {
												stiles.add(y);
											}
										}
										if (stiles.size() >= 2) {
											firstcol = y + 1;
											break;
										}
									}
									if (stiles.size() >= 2) {
										lessonsrow = i;
										startrow = i + 1;
										break;
									}
								}
								// ΑΝ ΔΕΝ ΒΡΩ ΟΛΕΣ ΤΙΣ ΣΤΗΛΕΣ
								// ----------------------------------------------------------------
								if (stiles.size() < 2) {
									mystr = "Δυστυχώς υπάρχει πρόβλημα με τα στοιχεία του αρχείου Βαθμολογιών της Δ τάξης. Δεν βρέθηκαν όλες οι απαραίτητες στήλες με τα δεδομένα.\n\nΒεβαιωθείτε ότι έχετε επιλέξει το σωστό αρχείο με τις βαθμολογίες.\n\nΑλλιώς επικοινωνείστε με τον προγραμματιστή...";
									textArea.setText(mystr);
									JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
											JOptionPane.ERROR_MESSAGE, null);
									return;
								}

								if (stiles.size() == 2) {
									System.out.println("");
									System.out.println("arxeio : " + doc.getElementsByTagName("arxeio").item(2)
											.getAttributes().getNamedItem("name").getNodeValue());
									System.out.println("ep1 : " + epikefalides1.item(2).getTextContent() + "\t στήλη : "
											+ stiles.get(0).toString());
									System.out.println("ep2 : " + epikefalides2.item(2).getTextContent() + "\t στήλη : "
											+ stiles.get(1).toString());
									System.out.println("αρχική σειρα δεδομένων : " + startrow);
								}
								// System.out.println(stiles);
								// System.out.println(startrow);
								// System.out.println(startcol);
								// System.out.println( "rows =" +
								// sheet.getLastRowNum());


								x = startrow;
								while (x <= sheet.getLastRowNum()) {
									if (sheet.getRow(x).getCell((int) stiles.get(0)).getCellType() == is_numeric) {
										am_map.put(
												Pair.with(
														Integer.toString((int) sheet.getRow(x)
																.getCell((int) stiles.get(0)).getNumericCellValue()),
														"file"),
												"d");
										am_map.put(
												Pair.with(
														Integer.toString((int) sheet.getRow(x)
																.getCell((int) stiles.get(0)).getNumericCellValue()),
														"row"),
												Integer.toString(x));
										// System.out.println(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
										// "row") + ", " +
										// am_map.get(Pair.with((int)sheet.getRow(x).getCell((int)stiles.get(0)).getNumericCellValue(),
										// "row")));
									}
									x = x + 1;
								}
								j = firstcol;
								while (j < sheet.getRow(lessonsrow).getLastCellNum()) {
									if (sheet.getRow(lessonsrow).getCell(j).getCellType() == is_string) {
										lesson_map.put(
												Pair.with("d",
														sheet.getRow(lessonsrow).getCell(j).getStringCellValue()),
												Integer.toString(j));
										// System.out.println(j + " " +
										// Pair.with("a",
										// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())
										// + ", " +
										// lesson_map.get(Pair.with("a",
										// sheet.getRow(lessonsrow).getCell(j).getStringCellValue())));
										lastcol = j;
									}
									j = j + 1;
								}
								// System.out.println( "lastcol d: " + lastcol);

								x = startrow;
								while (x <= sheet.getLastRowNum()) {
									for (j = firstcol; j <= lastcol; j++) {
										if (sheet.getRow(x).getCell(j).getCellType() == is_string) {
											data_map.put(Pair.with("d", Pair.with(x, j)),
													sheet.getRow(x).getCell(j).getStringCellValue());
											// System.out.println(Integer.toString(x)
											// + ","
											// + Integer.toString(j) + " " +
											// sheet.getRow(x).getCell(j).getStringCellValue());
										} else if (sheet.getRow(x).getCell(j).getCellType() == is_numeric) {
											data_map.put(Pair.with("d", Pair.with(x, j)),
													Double.toString(sheet.getRow(x).getCell(j).getNumericCellValue()));
										}
									}
									x = x + 1;
								}
								wb.close();
							}

						} catch (LeftoverDataException e) {
							// ΕΛΕΓΧΟΣ ΑΝΟΙΓΜΑΤΟΣ ΑΡΧΕΙΟΥ XLS ΤΟY MYSCHOOL
							e.printStackTrace();
							mystr = "Δυστυχώς δεν μπορώ να ανοίξω το αρχέιο xls:\n''" + myfile + "'' !!!"
									+ "\n\nΤα αρχεία xls που δημιουργεί αυτόματα το myschool με το εργαλείο του στην πλατφόρμα .aspx όταν διαβάζονται από το πρόσθετο ApachePOI της java που χειρίζεται τα αρχεία xls προξενούν το παρακάτω λάθος:"
									+ "\nError: " + e.getMessage()
									+ "\n\nΑνοίξτε το με το Excell ή το Libreoffice και Αποθηκεύστε το ως ''Excell 97''.";
							textArea.setText(mystr);
							JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
									JOptionPane.ERROR_MESSAGE, null);
							return;

						} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
							// ΕΛΕΓΧΟΣ ΑΝΟΙΓΜΑΤΟΣ ΑΡΧΕΙΟΥ XLS ΤΟY MYSCHOOL
							e.printStackTrace();
							mystr = "Δυστυχώς δεν μπορώ να ανοίξω το αρχέιο xls: ''" + myfile + "'' !!!" + "\n\nError: "
									+ e.getMessage();
							textArea.setText(mystr);
							JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: Ωχ... Λάθος!!!",
									JOptionPane.ERROR_MESSAGE, null);
							return;

						}

						// System.out.println(am_map.toString());
						// System.out.println(lesson_map.toString());
						// System.out.println(data_map.toString());

						mystr = "Επεξεργάστηκαν τα στοιχεία " + students.size() + " μαθητών.\n\n"
								+ "Επεξεργάστηκαν οι αναθέσεις " + anatheseis.size() + " καθηγητών για "
								+ lesson_map.size() + " μαθήματα.\n\n"
								+ "Επεξεργάστηκαν τα αρχεία \"Κατάσταση Μαθημάτων ανά Τάξη\" για κάθε τάξη.\n\n"
								+ "Δημιουργία των αρχείων για τους καθηγητές. Αυτό μπορεί να διαρκέσει λίγη ... ώρα!";
						System.out.println("\n\n\n" + mystr);
						textArea.setText(mystr);
						n = JOptionPane.showConfirmDialog(contentPane, textArea, "GΘ@2020: ΠΛΗΡΟΦΟΡΗΣΗ",
								JOptionPane.OK_CANCEL_OPTION);
						if (n == JOptionPane.CANCEL_OPTION) {
							return;
						}

						try {
							PrintWriter out;
							if (!prefix.equals("")) {
								File directory = new File(bathmologia_java.mydir + File.separator + dirprefix);
								directory.mkdirs();
								out = new PrintWriter(bathmologia_java.mydir + File.separator + dirprefix
										+ File.separator + unique_indentifier + ".txt");
							} else {
								out = new PrintWriter(
										bathmologia_java.mydir + File.separator + unique_indentifier + ".txt");
							}
							out.write(unique_indentifier);
							out.close();
							System.out.println("\n\n\nΤελευταίο Αρχείο Ελέγχου:");
							System.out.println(bathmologia_java.mydir + File.separator + unique_indentifier + ".txt");
						} catch (FileNotFoundException e) {
							e.printStackTrace();
						}

						try {

							contentPane.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
							pw = new ProgressWindow();

							int num = 0;
							int counter = 0;

							for (int i = 0; i < anatheseis.size(); i++) {
								num = num + anatheseis.get(i).size() - 1;
							}
							// System.out.println(num);

							for (int i = 0; i < anatheseis.size(); i++) {

								// ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ ΓΙΑ ΚΑΘΕ ΚΑΘΗΓΗΤΗ
								// System.out.println(anatheseis.get(i).get(0)[0]);
								String filename = anatheseis.get(i).get(0)[0];
								filename = filename.replaceAll("[^a-zA-Z0-9-\\p{InGreek}]", "_");
								myfile = textField_5.getText() + File.separator + prefix + filename + ".xls";
								wb = new HSSFWorkbook();

								CellStyle border = wb.createCellStyle();
								border.setBorderBottom(CellStyle.BORDER_THIN);
								border.setBorderLeft(CellStyle.BORDER_THIN);
								border.setBorderRight(CellStyle.BORDER_THIN);
								border.setBorderTop(CellStyle.BORDER_THIN);

								CellStyle bold = wb.createCellStyle();
								HSSFFont boldfont = (HSSFFont) wb.createFont();
								((HSSFFont) boldfont).setBold(true);
								bold.setFont((HSSFFont) boldfont);

								CellStyle border_bold = wb.createCellStyle();
								border_bold.setBorderBottom(CellStyle.BORDER_THIN);
								border_bold.setBorderLeft(CellStyle.BORDER_THIN);
								border_bold.setBorderRight(CellStyle.BORDER_THIN);
								border_bold.setBorderTop(CellStyle.BORDER_THIN);
								boldfont = (HSSFFont) wb.createFont();
								((HSSFFont) boldfont).setBold(true);
								border_bold.setFont((HSSFFont) boldfont);

								CellStyle border_bold_center = wb.createCellStyle();
								border_bold_center.setBorderBottom(CellStyle.BORDER_THIN);
								border_bold_center.setBorderLeft(CellStyle.BORDER_THIN);
								border_bold_center.setBorderRight(CellStyle.BORDER_THIN);
								border_bold_center.setBorderTop(CellStyle.BORDER_THIN);
								boldfont = (HSSFFont) wb.createFont();
								((HSSFFont) boldfont).setBold(true);
								border_bold_center.setFont((HSSFFont) boldfont);
								border_bold_center.setAlignment(CellStyle.ALIGN_CENTER);

								CellStyle bold_right = wb.createCellStyle();
								HSSFFont fnt = (HSSFFont) wb.createFont();
								((HSSFFont) fnt).setBold(true);
								bold_right.setFont((HSSFFont) fnt);
								bold_right.setAlignment(CellStyle.ALIGN_RIGHT);

								CellStyle border_center = wb.createCellStyle();
								border_center.setBorderBottom(CellStyle.BORDER_THIN);
								border_center.setBorderLeft(CellStyle.BORDER_THIN);
								border_center.setBorderRight(CellStyle.BORDER_THIN);
								border_center.setBorderTop(CellStyle.BORDER_THIN);
								border_center.setAlignment(CellStyle.ALIGN_CENTER);

								CellStyle border_center_1digit = wb.createCellStyle();
								border_center_1digit.setBorderBottom(CellStyle.BORDER_THIN);
								border_center_1digit.setBorderLeft(CellStyle.BORDER_THIN);
								border_center_1digit.setBorderRight(CellStyle.BORDER_THIN);
								border_center_1digit.setBorderTop(CellStyle.BORDER_THIN);
								border_center_1digit.setAlignment(CellStyle.ALIGN_CENTER);
								if (comboBox.getSelectedItem().toString().equals("ΓΡΑΠΤΟΣ")) {
									border_center_1digit.setDataFormat(wb.createDataFormat().getFormat("0.0"));
								}
								border_center_1digit.setLocked(false);

								CellStyle label = wb.createCellStyle();
								HSSFFont labelfont = (HSSFFont) wb.createFont();
								((HSSFFont) labelfont).setFontHeightInPoints((short) 18);
								((HSSFFont) labelfont).setBold(true);
								label.setFont((HSSFFont) labelfont);
								label.setAlignment(CellStyle.ALIGN_CENTER);

								CellStyle unlockedCell = wb.createCellStyle();
								unlockedCell.setLocked(false);

								for (int j = 1; j < anatheseis.get(i).size(); j++) {

									try {
										if (nosleep == false) {
											Thread.sleep(10);
										}
										pw.lb.setText("Δημιουργία αρχείων για " + filename.replace("_", " "));
										pw.pb.setValue(counter * 100 / num);

										// System.out.println(counter);
										counter++;
									} catch (InterruptedException ex) {
									}
									// ΔΗΜΙΟΥΡΓΙΑ ΦΥΛΛΟΥ ΓΙΑ ΚΑΘΕ ΜΑΘΗΜΑ
									// System.out.println(anatheseis.get(i).get(j)[0]
									// +
									// ", " + anatheseis.get(i).get(j)[1]);
									tmima = anatheseis.get(i).get(j)[0];
									mathima = anatheseis.get(i).get(j)[1];
									tmima_mathima = tmima + "-" + mathima;
									if (tmima_mathima.length() > 31) {
										Random rand = new Random();
										int minimum = 3;
										int maximum = (int) (31 - tmima.length()) / 2;
										if (maximum < 6)
											maximum = 6;
										int randomNum = minimum + rand.nextInt(maximum - minimum);

										tmima_mathima = tmima_mathima.substring(0, 31 - randomNum - 1) + "-"
												+ tmima_mathima.substring(tmima_mathima.length() - randomNum);
									}
									sheetname = WorkbookUtil.createSafeSheetName(tmima_mathima);
									sheet = wb.createSheet(sheetname);
									// ΓΕΜΙΣΜΑ ΤΟΥ ΦΥΛΛΟΥ ΜΕ ΔΕΔΟΜΕΝΑ
									// ΒΡΙΣΚΩ ΤΟΥΣ ΜΑΘΗΤΕΣ
									row2print = 5;
									aa = 1;
									for (int x = 0; x < students.size(); x++) {
										tmimata = ", " + students.get(x)[4];
										// System.out.println(tmimata + " - " +
										// tmima);
										if (tmimata.contains(", " + tmima + ",") == true) {
											if (aa == 1) {
												kodsentoni = am_map.get(Pair.with(students.get(x)[0], "file"));
												sheet.createRow(0).createCell(7).setCellValue(kodsentoni);
												// ΠΡΟΣΘΕΤΩ ΣΤΗΝ (0 based) ΣΤΗΛΗ
												// +1 ΓΙΑ ΕΝΑΡΜΟΝΙΣΗ ΜΕ ΤΟ
												// EXCELL (1 based)
												sheet.createRow(1).createCell(7).setCellValue(
														Integer.valueOf(lesson_map.get(Pair.with(kodsentoni, mathima)))
																+ 1);
												sheet.createRow(2).createCell(7)
														.setCellValue(comboBox.getSelectedItem().toString());
												sheet.createRow(3).createCell(7).setCellValue(unique_indentifier);
											}
											// System.out.println(students.get(x)[4]
											// + "
											// - " + anatheseis.get(i).get(j)[0]
											// + ",");
											myrow = sheet.createRow(row2print);
											myrow.createCell(0);
											myrow.createCell(1);
											myrow.createCell(2);
											myrow.createCell(3);
											myrow.createCell(4);
											myrow.createCell(5);
											myrow.createCell(6);
											myrow.createCell(7);
											sheet.getRow(row2print).getCell(0).setCellValue(aa);
											sheet.getRow(row2print).getCell(1).setCellValue(students.get(x)[0]);
											sheet.getRow(row2print).getCell(2).setCellValue(students.get(x)[1]);
											sheet.getRow(row2print).getCell(3).setCellValue(students.get(x)[2]);
											sheet.getRow(row2print).getCell(4).setCellValue(students.get(x)[3]);

											// ---------------------------------------------------------------------------
											// για debug βάζω τυχαία τιμή από 10
											// έως 20
											Boolean chkdebug = false;
											// chkdebug = true;
											if (chkdebug == true) {
												Random debugrand = new Random();
												int DebugNum = 10 + debugrand.nextInt(11);
												sheet.getRow(row2print).getCell(5).setCellValue(DebugNum);
											}
											// ---------------------------------------------------------------------------

											myref = new CellReference(row2print, 5).formatAsString();
											myformula = "IF( %s =\"\" , \"\" , PROPER(VLOOKUP(INT(ROUND(%s,1)),$I$1:$J$21,2)) & IF(%s-INT(%s)=0,\"\", IF(OR(ROUND(%s-INT(%s),1)=0,ROUND(%s-INT(%s),1)=1),\"\", \" & \" & VLOOKUP(ROUND(%s-INT(%s),1)*10,$I$2:$J11,2) & \" δεκ.\")))"
													.replace("%s", myref);
											sheet.getRow(row2print).getCell(6).setCellFormula(myformula);
											// System.out.println(myformula);
											// int therow =
											// Integer.valueOf(am_map.get(Pair.with(students.get(x)[0],"row")))
											// + 1;
											// ΠΡΟΣΘΕΤΩ ΣΤΗΝ (0 based) ΓΡΑΜΜΗ +1
											// ΓΙΑ
											// ΕΝΑΡΜΟΝΙΣΗ ΜΕ ΤΟ EXCELL (1 based)
											sheet.getRow(row2print).getCell(7).setCellValue(
													Integer.valueOf(am_map.get(Pair.with(students.get(x)[0], "row")))
															+ 1);
											// System.out.println(students.get(x)[0]
											// + "
											// : " +
											// am_map.get(Pair.with(students.get(x)[0],"row")));
											// να βαλω την τιμη απο το σεντονι
											// ---------------
											String existing_value = data_map.get(Pair.with(kodsentoni, Pair.with(
													Integer.valueOf(am_map.get(Pair.with(students.get(x)[0], "row"))),
													Integer.valueOf(lesson_map.get(Pair.with(kodsentoni, mathima))))));
											try {
												Double mynum = Double.parseDouble(existing_value);
												if ((mynum == Math.floor(mynum)) && !Double.isInfinite(mynum)) {
													sheet.getRow(row2print).getCell(5).setCellValue(mynum.intValue());
												} else {
													sheet.getRow(row2print).getCell(5).setCellValue(mynum);
												}
											} catch (NumberFormatException | NullPointerException e) {
											}

											sheet.getRow(row2print).getCell(0).setCellStyle(border_center);
											sheet.getRow(row2print).getCell(1).setCellStyle(border_center);
											sheet.getRow(row2print).getCell(2).setCellStyle(border);
											sheet.getRow(row2print).getCell(3).setCellStyle(border);
											sheet.getRow(row2print).getCell(4).setCellStyle(border);
											sheet.getRow(row2print).getCell(5).setCellStyle(border_center_1digit);
											sheet.getRow(row2print).getCell(6).setCellStyle(border);

											row2print = row2print + 1;
											aa = aa + 1;
										}
									}

									String[] bathmoi = { "μηδέν", "ένα", "δύο", "τρία", "τέσσερα", "πέντε", "έξι",
											"επτά", "οκτώ", "εννέα", "δέκα", "ένδεκα", "δώδεκα", "δεκατρία",
											"δεκατέσσερα", "δεκαπέντε", "δεκαέξι", "δεκαεπτά", "δεκαοκτώ", "δεκαεννέα",
											"είκοσι" };
									for (int k = 0; k < bathmoi.length; k++) {
										myrow = sheet.getRow(k);
										if (myrow == null)
											myrow = sheet.createRow(k);
										mycell = myrow.createCell(8);
										mycell.setCellValue(k);
										mycell = myrow.createCell(9);
										mycell.setCellValue(bathmoi[k]);
									}

									sheet.getRow(0).createCell(0);
									sheet.getRow(0).getCell(0).setCellValue(textField_7.getText());
									sheet.getRow(0).getCell(0).setCellStyle(bold); // style
									sheet.getRow(0).createCell(4);
									sheet.getRow(0).getCell(4)
											.setCellValue(textField_8.getText() + " " + textField_14.getText());
									sheet.getRow(0).getCell(4).setCellStyle(bold_right); // style
									sheet.getRow(1).createCell(0);
									sheet.getRow(1).getCell(0).setCellValue("ΚΑΤΑΣΤΑΣΗ ΒΑΘΜΟΛΟΓΙΑΣ");
									sheet.getRow(1).getCell(0).setCellStyle(label); // style
									sheet.getRow(1).setHeightInPoints(30); // style
									sheet.getRow(2).createCell(0);
									sheet.getRow(2).getCell(0).setCellValue(tmima);
									sheet.getRow(2).getCell(0).setCellStyle(bold); // style
									sheet.getRow(2).createCell(2);
									sheet.getRow(2).getCell(2).setCellValue(mathima);
									sheet.getRow(2).getCell(2).setCellStyle(bold); // style
									sheet.getRow(4).createCell(0);
									sheet.getRow(4).getCell(0).setCellValue("Α/Α");
									sheet.getRow(4).getCell(0).setCellStyle(border_bold_center); // style
									sheet.getRow(4).createCell(1);
									sheet.getRow(4).getCell(1).setCellValue("ΑΜ");
									sheet.getRow(4).getCell(1).setCellStyle(border_bold_center); // style
									sheet.getRow(4).createCell(2);
									sheet.getRow(4).getCell(2).setCellValue("ΕΠΩΝΥΜΟ");
									sheet.getRow(4).getCell(2).setCellStyle(border_bold); // style
									sheet.getRow(4).createCell(3);
									sheet.getRow(4).getCell(3).setCellValue("ΟΝΟΜΑ");
									sheet.getRow(4).getCell(3).setCellStyle(border_bold); // style
									sheet.getRow(4).createCell(4);
									sheet.getRow(4).getCell(4).setCellValue("ΠΑΤΡΩΝΥΜΟ");
									sheet.getRow(4).getCell(4).setCellStyle(border_bold); // style
									sheet.getRow(4).createCell(5);
									sheet.getRow(4).getCell(5).setCellValue("ΒΑΘΜΟΣ");
									sheet.getRow(4).getCell(5).setCellStyle(border_bold_center); // style
									sheet.getRow(4).createCell(6);
									sheet.getRow(4).getCell(6).setCellValue("ΟΛΟΓΡΑΦΩΣ");
									sheet.getRow(4).getCell(6).setCellStyle(border_bold); // style

									sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
									sheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 6));
									sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));

									sheet.setColumnWidth(0, 1200);
									sheet.setColumnWidth(1, 1500);
									sheet.setColumnWidth(2, 4500);
									sheet.setColumnWidth(3, 3500);
									sheet.setColumnWidth(4, 3500);
									sheet.setColumnWidth(5, 2500);
									sheet.setColumnWidth(6, 6500);
									sheet.setColumnHidden(7, true);
									sheet.setColumnHidden(8, true);
									sheet.setColumnHidden(9, true);

									CellRangeAddressList addressList = new CellRangeAddressList(5, 5 + aa - 1, 5, 5);
									DataValidation dataValidation;
									DVConstraint dvConstraint;
									if (comboBox.getSelectedItem().toString().equals("ΠΡΟΦΟΡΙΚΟΣ")) {
										dvConstraint = DVConstraint.createNumericConstraint(
												DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.BETWEEN,
												"0", "20");
									} else {
										dvConstraint = DVConstraint.createNumericConstraint(
												DVConstraint.ValidationType.DECIMAL, DVConstraint.OperatorType.BETWEEN,
												"0", "20");
									}
									dataValidation = new HSSFDataValidation(addressList, dvConstraint);
									dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
									if (comboBox.getSelectedItem().toString().equals("ΠΡΟΦΟΡΙΚΟΣ")) {
										dataValidation.createErrorBox("GΘ@2020: Έλεγχος Δεδομένων",
												"Ο προφορικός βαθμός πρέπει να είναι ακέραιος στο διάστημα 0 ως 20");
									} else {
										dataValidation.createErrorBox("GΘ@2020: Έλεγχος Δεδομένων",
												"Ο γραπτός βαθμός πρέπει να είναι δεκαδκός στο διάστημα 0 ως 20");
									}
									dataValidation.setEmptyCellAllowed(true);
									sheet.addValidationData(dataValidation);

									myrow = sheet.getRow(5 + aa);
									if (myrow == null)
										myrow = sheet.createRow(5 + aa);
									mycell = myrow.createCell(1);
									mycell.setCellValue("Καταχωρίστηκε ....../...../........");
									mycell.setCellStyle(unlockedCell);
									mycell = myrow.createCell(5);
									if (topos.equals("")) {
										mycell.setCellValue("Ο/Η Καθηγητής/τρια");
										mycell.setCellStyle(unlockedCell);
									} else {
										mycell.setCellValue(topos + " " + Hmnia);
										mycell.setCellStyle(unlockedCell);
										myrow = sheet.getRow(6 + aa);
										if (myrow == null)
											myrow = sheet.createRow(6 + aa);
										mycell = myrow.createCell(5);
										mycell.setCellValue("Ο/Η Καθηγητής/τρια");
										mycell.setCellStyle(unlockedCell);
									}

									if (topos.equals("")) {
										myrow = sheet.getRow(5 + aa + 4);
										if (myrow == null)
											myrow = sheet.createRow(5 + aa + 4);
									} else {
										myrow = sheet.getRow(6 + aa + 4);
										if (myrow == null)
											myrow = sheet.createRow(6 + aa + 4);
									}
									mycell = myrow.createCell(5);
									mycell.setCellValue(anatheseis.get(i).get(0)[0]);
									mycell.setCellStyle(unlockedCell);

									// ξεκλειδώνω κελιά γύρω από τη κατάσταση για 20 στήλες και γραμμές
									int mylastrow = aa + 12;
									for (int idxrow = 0; idxrow < aa + 12; idxrow++) {
										myrow = sheet.getRow(idxrow);
										if (myrow == null)
											myrow = sheet.createRow(idxrow);
										for (int idxcol = 10; idxcol < 61; idxcol++) {
											mycell = myrow.createCell(idxcol);
											mycell.setCellStyle(unlockedCell);
										}
									}
									
									for (int idxrow = mylastrow ; idxrow < mylastrow + 50; idxrow++) {
										myrow = sheet.getRow(idxrow);
										if (myrow == null)
											myrow = sheet.createRow(idxrow);
										for (int idxcol = 0; idxcol < 61; idxcol++) {
											mycell = myrow.getCell(idxcol);
											if(mycell == null) mycell = myrow.createCell(idxcol);
											mycell.setCellStyle(unlockedCell);
										}
									}

									sheet.protectSheet(mypassword);

									if (aa == 1) {
										wb.removeSheetAt(wb.getNumberOfSheets() - 1);
									}

								}
								if (wb.getNumberOfSheets() > 0) {
									// System.out.println(" ");
									sheetcount = sheetcount + wb.getNumberOfSheets();
									wbcount = wbcount + 1;
									otp = new FileOutputStream(myfile);
									wb.write(otp);
								}
								wb.close();
							}
							// αν υπάρχει πρόθεμα το συνδέω με τα αρχικά αρχεία
							// 187, αναθεσεις, τμήματα
							// και τα αντιγράφω για να υπάρχει αντιστοιχία με τα
							// αρχεία καθηγητών με το πρόθεμα
							if (!prefix.equals("")) {
								File source = new File(textField.getText());
								File dest = new File(source.getParent() + File.separator + dirprefix + File.separator
										+ prefix + source.getName());
								FileUtils.copyFile(source, dest);

								File source_1 = new File(textField_1.getText());
								File dest_1 = new File(source_1.getParent() + File.separator + dirprefix
										+ File.separator + prefix + source_1.getName());
								FileUtils.copyFile(source_1, dest_1);

								File source_2 = new File(textField_2.getText());
								File dest_2 = new File(source_2.getParent() + File.separator + dirprefix
										+ File.separator + prefix + source_2.getName());
								FileUtils.copyFile(source_2, dest_2);

								File source_3 = new File(textField_3.getText());
								File dest_3 = new File(source_3.getParent() + File.separator + dirprefix
										+ File.separator + prefix + source_3.getName());
								FileUtils.copyFile(source_3, dest_3);

								File source_4 = new File(textField_4.getText());
								File dest_4 = new File(source_4.getParent() + File.separator + dirprefix
										+ File.separator + prefix + source_4.getName());
								FileUtils.copyFile(source_4, dest_4);
								if (esperino == true) {
									File source_13 = new File(textField_13.getText());
									File dest_13 = new File(source_13.getParent() + File.separator + dirprefix
											+ File.separator + prefix + source_13.getName());
									FileUtils.copyFile(source_13, dest_13);
								}
							}

						} catch (IOException e) {
							e.printStackTrace();
						}

						contentPane.setCursor(null);
						pw.frame.setVisible(false);
						pw = null;

						mystr = "Δημιουργήθηκαν " + sheetcount + " Βαθμολογικές Καταστάσεις σε " + wbcount
								+ " αρχεία καθηγητών";
						System.out.println("\n\n\n" + mystr);
						textArea.setText(mystr);
						JOptionPane.showMessageDialog(contentPane, textArea, "GΘ@2020: ΤΕΛΟΣ ΔΙΑΔΙΚΑΣΙΑΣ",
								JOptionPane.INFORMATION_MESSAGE, null);

					}
				};

				worker.start(); // So we don't hold up the dispatch thread.

			}

		});

		JLabel lblPassword = new JLabel("Password για κλείδωμα xls");
		lblPassword.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_lblPassword = new GridBagConstraints();
		gbc_lblPassword.anchor = GridBagConstraints.WEST;
		gbc_lblPassword.insets = new Insets(0, 0, 5, 5);
		gbc_lblPassword.gridx = 0;
		gbc_lblPassword.gridy = 20;
		contentPane.add(lblPassword, gbc_lblPassword);

		textField_9 = new JTextField();
		textField_9.setToolTipText("Πληκτρολογείστε Password για (ξε)κλείδωμα του xls");
		textField_9.setColumns(10);
		GridBagConstraints gbc_textField_9 = new GridBagConstraints();
		gbc_textField_9.gridwidth = 2;
		gbc_textField_9.insets = new Insets(0, 0, 5, 5);
		gbc_textField_9.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_9.gridx = 1;
		gbc_textField_9.gridy = 20;
		contentPane.add(textField_9, gbc_textField_9);

		JSeparator separator_7 = new JSeparator();
		GridBagConstraints gbc_separator_7 = new GridBagConstraints();
		gbc_separator_7.fill = GridBagConstraints.BOTH;
		gbc_separator_7.gridwidth = 3;
		gbc_separator_7.insets = new Insets(0, 0, 5, 5);
		gbc_separator_7.gridx = 0;
		gbc_separator_7.gridy = 19;
		contentPane.add(separator_7, gbc_separator_7);

		JLabel label_11 = new JLabel("Πρόθεμα αρχείων");
		label_11.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_label_11 = new GridBagConstraints();
		gbc_label_11.anchor = GridBagConstraints.WEST;
		gbc_label_11.insets = new Insets(0, 0, 5, 5);
		gbc_label_11.gridx = 0;
		gbc_label_11.gridy = 21;
		contentPane.add(label_11, gbc_label_11);

		textField_10 = new JTextField();
		textField_10.setToolTipText("Πληκτρολογείστε ένα πρόθεμα που θα ξεχωρίσει τα αρχεία σας από παλαιότερα");
		textField_10.setColumns(10);
		GridBagConstraints gbc_textField_10 = new GridBagConstraints();
		gbc_textField_10.gridwidth = 2;
		gbc_textField_10.insets = new Insets(0, 0, 5, 5);
		gbc_textField_10.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_10.gridx = 1;
		gbc_textField_10.gridy = 21;
		contentPane.add(textField_10, gbc_textField_10);

		JSeparator separator_6 = new JSeparator();
		GridBagConstraints gbc_separator_6 = new GridBagConstraints();
		gbc_separator_6.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator_6.gridwidth = 4;
		gbc_separator_6.insets = new Insets(0, 0, 5, 0);
		gbc_separator_6.gridx = 0;
		gbc_separator_6.gridy = 22;
		contentPane.add(separator_6, gbc_separator_6);

		JLabel lblkey = new JLabel("Αρχείο ελέγχου συλλογής");
		lblkey.setHorizontalAlignment(SwingConstants.LEFT);
		GridBagConstraints gbc_lblkey = new GridBagConstraints();
		gbc_lblkey.anchor = GridBagConstraints.WEST;
		gbc_lblkey.insets = new Insets(0, 0, 5, 5);
		gbc_lblkey.gridx = 0;
		gbc_lblkey.gridy = 23;
		contentPane.add(lblkey, gbc_lblkey);

		textField_11 = new JTextField();
		textField_11.setToolTipText(
				"Επιλέξτε ένα Αρχείο ελέγχου αν θέλετε να φιλτράρετε τα αρχεία που σας φέρνουν πίσω οι καθηγητές");
		textField_11.setColumns(10);
		GridBagConstraints gbc_textField_11 = new GridBagConstraints();
		gbc_textField_11.gridwidth = 2;
		gbc_textField_11.insets = new Insets(0, 0, 5, 5);
		gbc_textField_11.fill = GridBagConstraints.HORIZONTAL;
		gbc_textField_11.gridx = 1;
		gbc_textField_11.gridy = 23;
		contentPane.add(textField_11, gbc_textField_11);

		JButton button_7 = new JButton("...");
		button_7.setToolTipText("Επιλέξτε ένα Αρχείο ελέγχου");
		button_7.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser openFile = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Aρχεία κειμένου", "txt");
				openFile.setFileFilter(filter);
				if (bathmologia_java.mydir != null) {
					openFile.setCurrentDirectory(new java.io.File(bathmologia_java.mydir));
				}
				openFile.setDialogTitle("Επιλογή αρχείων");
				int returnVal = openFile.showOpenDialog(null);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textField_11.setText(openFile.getSelectedFile().getAbsolutePath());
					bathmologia_java.mydir = openFile.getSelectedFile().getParent();
				}

			}
		});
		button_7.setFont(new Font("Arial Black", Font.BOLD, 6));
		GridBagConstraints gbc_button_7 = new GridBagConstraints();
		gbc_button_7.insets = new Insets(0, 0, 5, 0);
		gbc_button_7.gridx = 3;
		gbc_button_7.gridy = 23;
		contentPane.add(button_7, gbc_button_7);

		JSeparator separator_5 = new JSeparator();
		GridBagConstraints gbc_separator_5 = new GridBagConstraints();
		gbc_separator_5.fill = GridBagConstraints.HORIZONTAL;
		gbc_separator_5.gridwidth = 4;
		gbc_separator_5.insets = new Insets(2, 0, 5, 0);
		gbc_separator_5.gridx = 0;
		gbc_separator_5.gridy = 24;
		contentPane.add(separator_5, gbc_separator_5);

		GridBagConstraints gbc_btnNewButton = new GridBagConstraints();
		gbc_btnNewButton.gridwidth = 2;
		gbc_btnNewButton.insets = new Insets(0, 0, 5, 5);
		gbc_btnNewButton.gridx = 0;
		gbc_btnNewButton.gridy = 25;
		contentPane.add(btnNewButton, gbc_btnNewButton);
		GridBagConstraints gbc_button_8 = new GridBagConstraints();
		gbc_button_8.gridwidth = 2;
		gbc_button_8.insets = new Insets(0, 0, 5, 0);
		gbc_button_8.gridx = 2;
		gbc_button_8.gridy = 25;
		contentPane.add(button_8, gbc_button_8);
		GridBagConstraints gbc_btnXls = new GridBagConstraints();
		gbc_btnXls.gridwidth = 2;
		gbc_btnXls.insets = new Insets(0, 0, 0, 5);
		gbc_btnXls.gridx = 0;
		gbc_btnXls.gridy = 26;
		contentPane.add(btnXls, gbc_btnXls);

		button_9 = new JButton("Έξοδος");
		button_9.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String espvalue;
				if (esperino == true) {
					espvalue = "1";
				} else {
					espvalue = "";
				}

				Properties props = new Properties();
				props.setProperty("esperino", espvalue);
				props.setProperty("arxA", textField.getText());
				props.setProperty("arxB", textField_1.getText());
				props.setProperty("arxG", textField_2.getText());
				props.setProperty("arxD", textField_13.getText());
				props.setProperty("anatheseis", textField_3.getText());
				props.setProperty("tmimata", textField_4.getText());
				props.setProperty("out", textField_5.getText());
				props.setProperty("in", textField_6.getText());
				props.setProperty("sch_name", textField_7.getText());
				props.setProperty("periodos", textField_8.getText());
				props.setProperty("sch_year", textField_14.getText());
				props.setProperty("poli", textField_12.getText());
				props.setProperty("date", textField_15.getText());
				props.setProperty("pass", textField_9.getText());
				props.setProperty("prothema", textField_10.getText());

				try {
					InputStream in = getClass().getResourceAsStream("epikefalides.xml");
					DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
					DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
					doc = docBuilder.parse(in);
					configFile = doc.getElementsByTagName("configFile");

					String newAppConfigXmlFile = workdir + File.separator + configFile.item(0).getTextContent();
					props.storeToXML(new FileOutputStream(newAppConfigXmlFile), null);
					
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (ParserConfigurationException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (SAXException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				System.exit(0);
			}
		});
		GridBagConstraints gbc_button_9 = new GridBagConstraints();
		gbc_button_9.gridwidth = 2;
		gbc_button_9.gridx = 2;
		gbc_button_9.gridy = 26;
		contentPane.add(button_9, gbc_button_9);
		/*
		 * JSeparator separator_6 = new JSeparator(); GridBagConstraints
		 * gbc_separator_6 = new GridBagConstraints(); gbc_separator_6.anchor =
		 * GridBagConstraints.NORTH; gbc_separator_6.fill =
		 * GridBagConstraints.HORIZONTAL; gbc_separator_6.gridwidth = 4;
		 * gbc_separator_6.insets = new Insets(2, 0, 0, 0);
		 * gbc_separator_6.gridx = 0; gbc_separator_6.gridy = 26;
		 * contentPane.add(separator_6, gbc_separator_6);
		 */
		try {
			configFile = doc.getElementsByTagName("configFile");
			
			String configPath = workdir + File.separator + configFile.item(0).getTextContent();
			Properties props = new Properties();
			props.loadFromXML(new FileInputStream(configPath));
			
			String esperinode = props.getProperty("esperino");
			String arxA = props.getProperty("arxA");
			String arxB = props.getProperty("arxB");
			String arxG = props.getProperty("arxG");
			String arxD = props.getProperty("arxD");
			String anatheseis = props.getProperty("anatheseis");
			String tmimata = props.getProperty("tmimata");
			String out = props.getProperty("out");
			String in = props.getProperty("in");
			String sch_name = props.getProperty("sch_name");
			String periodos = props.getProperty("periodos");
			String sch_year = props.getProperty("sch_year");
			String poli = props.getProperty("poli");
			String date = props.getProperty("date");
			String pass = props.getProperty("pass");
			String prothema = props.getProperty("prothema");
			
			if (esperinode.equals("1"))
				comboBox_2.setSelectedIndex(1);
			textField.setText(arxA);
			textField_1.setText(arxB);
			textField_2.setText(arxG);
			textField_13.setText(arxD);
			textField_3.setText(anatheseis);
			textField_4.setText(tmimata);
			textField_5.setText(out);
			textField_6.setText(in);
			textField_7.setText(sch_name);
			textField_8.setText(periodos);
			textField_14.setText(sch_year);
			textField_12.setText(poli);
			textField_15.setText(date);
			textField_9.setText(pass);
			textField_10.setText(prothema);
			
				if (!arxA.equals("")) {
					File file = new File(arxA);
					bathmologia_java.mydir = file.getParent();
				}
				if (bathmologia_java.mydir == null) {
					if (!arxB.equals("")) {
						File file = new File(arxB);
						bathmologia_java.mydir = file.getParent();
					}
				}
				if (bathmologia_java.mydir == null) {
					if (!arxG.equals("")) {
						File file = new File(arxG);
						bathmologia_java.mydir = file.getParent();
					}
				}
				if (bathmologia_java.mydir == null) {
					if (!arxD.equals("")) {
						File file = new File(arxD);
						bathmologia_java.mydir = file.getParent();
					}
				}
				if (bathmologia_java.mydir == null) {
					if (!anatheseis.equals("")) {
						File file = new File(anatheseis);
						bathmologia_java.mydir = file.getParent();
					}
				}
				if (bathmologia_java.mydir == null) {
					if (!tmimata.equals("")) {
						File file = new File(tmimata);
						bathmologia_java.mydir = file.getParent();
					}
				}
				if (bathmologia_java.mydir == null) {
					if (!out.equals("")) {
						File file = new File(out);
						bathmologia_java.mydir = file.getParent();
					}
				}
				if (bathmologia_java.mydir == null) {
					if (!in.equals("")) {
						File file = new File(in);
						bathmologia_java.mydir = file.getParent();
					}
				}
			// }
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public class ProgressWindow {
		public JProgressBar pb;
		public JLabel lb;
		public JDialog frame;

		public ProgressWindow() {
			frame = new JDialog();
			frame.setTitle("GΘ@2020: ΠΡΟΟΔΟΣ...");
			frame.setSize(500, 110);
			frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);			GridBagLayout gridBagLayout = new GridBagLayout();
			gridBagLayout.columnWidths = new int[] { 0 };
			gridBagLayout.rowHeights = new int[] { 0, 0 };
			gridBagLayout.columnWeights = new double[] { 0.0 };
			gridBagLayout.rowWeights = new double[] { 0.0, 0.0 };
			frame.getContentPane().setLayout(gridBagLayout);

			lb = new JLabel("");
			lb.setHorizontalAlignment(SwingConstants.LEFT);
			GridBagConstraints gbc_lb = new GridBagConstraints();
			gbc_lb.anchor = GridBagConstraints.WEST;
			gbc_lb.insets = new Insets(0, 15, 10, 15);
			gbc_lb.gridx = 0;
			gbc_lb.gridy = 0;
			frame.getContentPane().add(lb, gbc_lb);

			pb = new JProgressBar();
			pb.setStringPainted(true);
			GridBagConstraints gbc_pb = new GridBagConstraints();
			gbc_pb.insets = new Insets(0, 15, 25, 15);
			gbc_pb.fill = GridBagConstraints.BOTH;
			gbc_pb.weighty = 1.0;
			gbc_pb.weightx = 1.0;
			gbc_pb.gridx = 0;
			gbc_pb.gridy = 1;
			frame.getContentPane().add(pb, gbc_pb);

			frame.setLocationRelativeTo(contentPane);
			frame.setVisible(true);

		}
	}

}
