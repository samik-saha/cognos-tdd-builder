package com.cognizant.cognos.tddbuilder;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.JToolBar;
import javax.swing.ProgressMonitor;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;


final class Prompt {
	public String name, parameter, sort, required, format,comment;
}

final class ConditionalVariable {
	public String name, type, logic, values, comment;
}

final class DataItem{
	public String dataItem, nameInPackage, order, sortLevel, aggregation, drill, comments;
}

final class Query{
	public String queryName;
	public DataItem[] dataItems;
	public String[] filters;
}

final class ReportPage{
		public String queryName;
		public Query[] queries;
}

public class MainWindow {
	private JFrame frmCognosTddBuilder;
	protected static Document xmlDocument;
	private XPath xPath;
	private JTextArea logPane;
	private String reportName;
	private JEditorPane specPane;
	private JButton btnExportExcel;
	private JButton btnExportWord;
	JFileChooser fc;
	private ProgressMonitor progressMonitor;
	ReportDataExtracter task;
	private JButton btnImportXml;
	private JMenuItem mntmImportXmlFile;
	String outputFilename;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainWindow window = new MainWindow();
					window.frmCognosTddBuilder.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public MainWindow() {
		// Set look and feel
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (UnsupportedLookAndFeelException | ClassNotFoundException
				| InstantiationException | IllegalAccessException ex) {
			Logger.getLogger(MainWindow.class.getName()).log(Level.SEVERE,
					null, ex);
		}

		// Initialize
		fc = new JFileChooser(".");

		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmCognosTddBuilder = new JFrame();
		frmCognosTddBuilder.setTitle("Cognos TDD Builder");
		frmCognosTddBuilder.setBounds(100, 100, 450, 300);
		frmCognosTddBuilder.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		JMenuBar menuBar = new JMenuBar();
		frmCognosTddBuilder.setJMenuBar(menuBar);

		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);

		mntmImportXmlFile = new JMenuItem("Import XML File");
		mntmImportXmlFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				importXMLEventHandler(e);
			}
		});
		mnFile.add(mntmImportXmlFile);

		JMenuItem mntmExit = new JMenuItem("Exit");
		mnFile.add(mntmExit);

		JMenu mnActions = new JMenu("Actions");
		menuBar.add(mnActions);

		JMenu mnHelp = new JMenu("Help");
		menuBar.add(mnHelp);

		JToolBar toolBar = new JToolBar();
		toolBar.setFloatable(false);
		frmCognosTddBuilder.getContentPane().add(toolBar, BorderLayout.NORTH);

		btnImportXml = new JButton("");
		btnImportXml.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				importXMLEventHandler(arg0);
			}
		});
		btnImportXml.setToolTipText("Import XML File");
		btnImportXml
				.setIcon(new ImageIcon(
						MainWindow.class
								.getResource("/com/cognizant/cognos/tddbuilder/res/document-import.png")));
		toolBar.add(btnImportXml);

		btnExportWord = new JButton("");
		btnExportWord.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				prepareTDDWordDoc();
			}
		});
		btnExportWord.setEnabled(false);
		btnExportWord.setToolTipText("Export Word Document");
		btnExportWord
				.setIcon(new ImageIcon(
						MainWindow.class
								.getResource("/com/cognizant/cognos/tddbuilder/res/page_white_word.png")));
		toolBar.add(btnExportWord);

		btnExportExcel = new JButton("");
		btnExportExcel.setToolTipText("Export Excel Document");
		btnExportExcel.setEnabled(false);
		btnExportExcel
				.setIcon(new ImageIcon(
						MainWindow.class
								.getResource("/com/cognizant/cognos/tddbuilder/res/page_white_excel.png")));
		toolBar.add(btnExportExcel);

		JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.BOTTOM);
		frmCognosTddBuilder.getContentPane().add(tabbedPane,
				BorderLayout.CENTER);

		JScrollPane scrollPane_1 = new JScrollPane();
		tabbedPane.addTab("Spec", null, scrollPane_1, null);

		specPane = new JEditorPane();
		specPane.setEditable(false);
		specPane.setContentType("text/html");
		scrollPane_1.setViewportView(specPane);

		JScrollPane scrollPane = new JScrollPane();
		tabbedPane.addTab("Log", null, scrollPane, null);

		logPane = new JTextArea();
		logPane.setEditable(false);
		scrollPane.setViewportView(logPane);
	}

	private void importXMLEventHandler(ActionEvent e) {
		fc.resetChoosableFileFilters();
		fc.setAcceptAllFileFilterUsed(false);
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
				"XML Files", "xml");
		fc.setFileFilter(filter);
		fc.setSelectedFile(null);
		int returnVal = fc.showOpenDialog(frmCognosTddBuilder);

		if (returnVal == JFileChooser.APPROVE_OPTION) {
			File file = fc.getSelectedFile();
			try {
				DocumentBuilderFactory builderFactory = DocumentBuilderFactory
						.newInstance();
				builderFactory.setValidating(false);
				builderFactory.setFeature(
						"http://xml.org/sax/features/namespaces", false);
				builderFactory.setFeature(
						"http://xml.org/sax/features/validation", false);
				builderFactory
						.setFeature(
								"http://apache.org/xml/features/nonvalidating/load-dtd-grammar",
								false);
				builderFactory
						.setFeature(
								"http://apache.org/xml/features/nonvalidating/load-external-dtd",
								false);

				DocumentBuilder builder = builderFactory.newDocumentBuilder();
				xmlDocument = builder.parse(file);

				log("XML imported successfully from " + file.getCanonicalPath());
				btnExportExcel.setEnabled(true);
				btnExportWord.setEnabled(true);
				displaySpec();

			} catch (IOException | SAXException | ParserConfigurationException ex) {
				ex.printStackTrace();
			}
		}
	}

	public void log(String str) {
		logPane.setText(logPane.getText() + str + "\n");
	}

	private void displaySpec() {
		String htmlText;
		xPath = XPathFactory.newInstance().newXPath();
		try {
			reportName = (String) xPath.compile("/report/reportName/text()")
					.evaluate(xmlDocument, XPathConstants.STRING);
			htmlText = "<b>Report Name:</b> " + reportName;
			htmlText += "<br><b>Report Pages:</b><ul>";
			NodeList rptPages = (NodeList) xPath.compile("//reportPages/page")
					.evaluate(xmlDocument, XPathConstants.NODESET);
			for (int i = 0; i < rptPages.getLength(); i++) {
				String pageName = rptPages.item(i).getAttributes()
						.getNamedItem("name").getNodeValue();
				htmlText += "<li>" + pageName + "</li>";
			}
			htmlText += "</ul>";
			String modelPath=(String) xPath.compile("/report/modelPath/text()")
					.evaluate(xmlDocument, XPathConstants.STRING);
			Pattern p = Pattern.compile("package\\[@name='(.*?)'\\]");
			Matcher m = p.matcher(modelPath);
			String packageName="";
			while(m.find()){
				packageName = m.group(1);
			}
			htmlText += "<b>Package Name:</b> "+packageName;

			specPane.setText(htmlText);
		} catch (XPathExpressionException ex) {
			Logger.getLogger(MainWindow.class.getName()).log(Level.SEVERE,
					null, ex);
		}

	}

	private void prepareTDDWordDoc() {
		fc.resetChoosableFileFilters();
		fc.setAcceptAllFileFilterUsed(false);
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
				"Word document (*.docx)", "docx");
		fc.setFileFilter(filter);
		String suggestedFilename = reportName.replaceAll("\\W+", "_")+".docx";
		fc.setSelectedFile(new File(suggestedFilename));
		int returnVal = fc.showSaveDialog(frmCognosTddBuilder);

		if (returnVal == JFileChooser.APPROVE_OPTION) {
			try {
				enableUserInteraction(false);
				
				outputFilename = fc.getSelectedFile().getCanonicalPath();
				ExportedOutputInterface exportedWordOutput = new WordOutput(
						this, outputFilename);

				progressMonitor = new ProgressMonitor(getFrame(), "Builing TDD",
						"", 0, 100);

				task = new ReportDataExtracter(this, exportedWordOutput);
				task.addPropertyChangeListener(new PropertyChangeListener() {

					@Override
					public void propertyChange(PropertyChangeEvent evt) {
						if ("progress" == evt.getPropertyName()) {
							int progress = (Integer) evt.getNewValue();
							progressMonitor.setProgress(progress);
							String message = String.format("Completed %d%%.\n",
									progress);
							progressMonitor.setNote(message);
							if (progressMonitor.isCanceled() || task.isDone()) {
								if (progressMonitor.isCanceled()) {
									task.cancel(true);
								} else {
									progressMonitor.close();
								}
							}
						}

					}
				});

				task.execute();
			} catch (IOException e) {
				e.printStackTrace();
			}

		}

	}
	

	protected void enableUserInteraction(boolean flag){
		btnImportXml.setEnabled(flag);
		btnExportExcel.setEnabled(flag);
		btnExportWord.setEnabled(flag);
		mntmImportXmlFile.setEnabled(flag);
	}

	public JFrame getFrame() {
		return frmCognosTddBuilder;
	}
}
