package com.cognizant.cognos.tddbuilder;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.NodeList;

public class ReportDataExtracter extends SwingWorker<Void, Void> {
	ExportedOutputInterface exportedOutput;
	Prompt[] prompts;
	XPath xPath;
	MainWindow mainWindow;
	String outputFilename;

	public ReportDataExtracter(MainWindow mainWindow,
			ExportedOutputInterface exportedOutput) {
		this.exportedOutput = exportedOutput;
		this.mainWindow = mainWindow;
		xPath = XPathFactory.newInstance().newXPath();
	}

	@Override
	protected Void doInBackground() throws Exception {
		mainWindow.log("Looking for prompts...");
		extractPromptDetails();
		setProgress(30);
		Thread.sleep(2000);
		outputFilename = exportedOutput.writeToFile();
		setProgress(90);
		Thread.sleep(2000);
		setProgress(100);
		return null;
	}

	@Override
	protected void done() {
		mainWindow.enableUserInteraction(true);
		if (outputFilename != null){
	        Object[] options = { "Open File", "Cancel" };
			int n = JOptionPane.showOptionDialog(mainWindow.getFrame(),
					"TDD generation complete. Would you like to open the file?","Information", JOptionPane.OK_CANCEL_OPTION,
					JOptionPane.INFORMATION_MESSAGE, null, options, options[0]);
			if (n==JOptionPane.OK_OPTION){
				try {
					Desktop.getDesktop().open(new File(outputFilename));
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
			
	}

	private void extractPromptDetails() {
		try {
			NodeList selectValueNodes;
			selectValueNodes = (NodeList) xPath.compile(
					"//promptPages//selectValue").evaluate(
					MainWindow.xmlDocument, XPathConstants.NODESET);
			mainWindow.log(selectValueNodes.getLength() + " prompts found.");
			prompts = new Prompt[selectValueNodes.getLength()];
			for (int i = 0; i < selectValueNodes.getLength(); i++) {
				NamedNodeMap attributes = selectValueNodes.item(i)
						.getAttributes();
				Prompt prompt = new Prompt();
				prompt.name = attributes.getNamedItem("name") != null ? attributes
						.getNamedItem("name").getNodeValue() : "";
				prompt.parameter = attributes.getNamedItem("parameter") != null ? attributes
						.getNamedItem("parameter").getNodeValue() : "";
				prompt.required = attributes.getNamedItem("required") != null ? attributes
						.getNamedItem("required").getNodeValue() : "";
				prompt.format = attributes.getNamedItem("selectValueUI") != null ? attributes
						.getNamedItem("selectValueUI").getNodeValue() : "";
				prompts[i] = prompt;
			}

			exportedOutput.writePromptDetails(prompts);

		} catch (XPathExpressionException e) {
			e.printStackTrace();
		}
	}
}
