package com.cognizant.cognos.tddbuilder;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class ReportDataExtracter extends SwingWorker<Void, Void> {
	ExportedOutputInterface exportedOutput;
	Prompt[] prompts;
	ConditionalVariable[] conditionalVariables;
	Query[] queries;
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
		extractConditionalVariables();
		setProgress(40);
		extractReportPages();
		extractReportQueries();
		outputFilename = exportedOutput.writeToFile();
		setProgress(90);

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
			NodeList promptNodes;
			promptNodes = (NodeList) xPath.compile(
					"//promptPages//*[@parameter]").evaluate(
					MainWindow.xmlDocument, XPathConstants.NODESET);
			mainWindow.log(promptNodes.getLength() + " prompts found.");
			prompts = new Prompt[promptNodes.getLength()];
			for (int i = 0; i < promptNodes.getLength(); i++) {
				Node promptNode = promptNodes.item(i);
				NamedNodeMap attributes = promptNodes.item(i)
						.getAttributes();
				Prompt prompt = new Prompt();
				prompt.name = attributes.getNamedItem("name") != null ? attributes
						.getNamedItem("name").getNodeValue() : "NA";
				prompt.parameter = attributes.getNamedItem("parameter") != null ? attributes
						.getNamedItem("parameter").getNodeValue() : "";
				prompt.required = attributes.getNamedItem("required") != null ? attributes
						.getNamedItem("required").getNodeValue() : "true";
				
				/* Extract Prompt Format */
			    String nodeName = promptNode.getNodeName();
			    switch (nodeName){
			    	case "selectValue":
			    		String selectValueUI=attributes.getNamedItem("selectValueUI") != null ? attributes
								.getNamedItem("selectValueUI").getNodeValue() : "";
			    		if (selectValueUI.equals("radioGroup")){
			    			prompt.format = "Radio Button Group Prompt";
			    		}
			    		else
			    			prompt.format = "Dropdown List Prompt";
			    		break;
			    	case "textBox":
			    		prompt.format = "Text Box Prompt";
			    		break;
			    	case "selectWithSearch":
			    		prompt.format = "Select And Search Prompt";
			    		break;
			    	case "selectDate":
			    		prompt.format = "Date Prompt";
			    		break;
			    	case "selectTime":
			    		prompt.format = "Time Prompt";
			    		break;
			    	case "selectDateTime":
			    		prompt.format = "Date/Time Prompt";
			    		break;
			    	case "selectInterval":
			    		prompt.format = "Interval Prompt";
			    		break;
			    	case "selectWithTree":
			    		prompt.format = "Tree Prompt";
			    		break;
			    	case "generatedPrompt":
			    		prompt.format = "Generated Prompt";
			    		break;
			    	default: 
			    		break;
			    }
			    
			    /* Sort Order */
			    //Check for sortList node
			    Node sortList = (Node) xPath.compile("//promptPages//"+promptNode.getNodeName()+"[@parameter=\""+prompt.parameter+"\"]/sortList").evaluate(promptNode, XPathConstants.NODE);
			    prompt.sort="";
			    if(sortList != null){
			    	NodeList sortItemList = sortList.getChildNodes();
			    	for(int j=0; j<sortItemList.getLength();j++){
			    		Node sortItem = sortItemList.item(j);
			    		String refDataItem =sortItem.getAttributes().getNamedItem("refDataItem").getTextContent();
			    		String sortOrder=sortItem.getAttributes().getNamedItem("sortOrder") != null?
			    				sortItem.getAttributes().getNamedItem("sortOrder").getTextContent():"ascending";
			    		prompt.sort+=refDataItem+" ("+sortOrder+")\n";
			    	}
			    }
			    else {
			    	prompt.sort = "NA";
			    }
			    
			    
			    prompt.comment="";
			    prompt.comment+=attributes.getNamedItem("cascadeOn") != null ? "Cascade on "+attributes
						.getNamedItem("cascadeOn").getNodeValue()+"\n" : "";
						
				prompt.comment+=attributes.getNamedItem("prePopulateIfParentOptional") != null ? "Pre-populated":"";
			    

				prompts[i] = prompt;
				
			
			}

			exportedOutput.writePromptDetails(prompts);

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.getMessage());
		}
	}

	private void extractConditionalVariables(){
		Node reportExprNode, variableValuesNode;
		try {
			NodeList reportVariableNodes;
			reportVariableNodes = (NodeList) xPath.compile(
					"//reportVariables/reportVariable").evaluate(
					MainWindow.xmlDocument, XPathConstants.NODESET);
			mainWindow.log(reportVariableNodes.getLength() + " report variables found.");
			conditionalVariables = new ConditionalVariable[reportVariableNodes.getLength()];
			for (int i = 0; i < reportVariableNodes.getLength(); i++) {
				Node reportVariableNode = reportVariableNodes.item(i);
				NamedNodeMap attributes = reportVariableNode.getAttributes();
				ConditionalVariable conditionalVariable = new ConditionalVariable();
				conditionalVariable.name = attributes.getNamedItem("name").getTextContent();
				conditionalVariable.type = attributes.getNamedItem("type").getTextContent();
				NodeList childNodes = reportVariableNode.getChildNodes();
				reportExprNode = null;
				variableValuesNode = null;
				for(int j = 0; j < childNodes.getLength(); j++){
					if(childNodes.item(j).getNodeName() == "reportExpression"){
						reportExprNode = childNodes.item(j);
					}
					else if (childNodes.item(j).getNodeName() == "variableValues"){
						variableValuesNode = childNodes.item(j);
					}
				}
				
				conditionalVariable.logic = reportExprNode != null? reportExprNode.getTextContent():"";
				
				conditionalVariable.values = "";
				if (variableValuesNode != null){
					NodeList variableValues = variableValuesNode.getChildNodes();
					for (int k = 0; k < variableValues.getLength(); k++){
						Node variableValue = variableValues.item(k);
						if(variableValue.getNodeName() == "variableValue"){
							String value = variableValue.getAttributes()!=null?
									variableValue.getAttributes().getNamedItem("value").getNodeValue():"";

							conditionalVariable.values += (conditionalVariable.values !=""? ", ":"")+value;
						}
					}
				}
				
				conditionalVariables[i]=conditionalVariable;
			}
			
			exportedOutput.writeConditionalVariableDetails(conditionalVariables);
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	private void extractReportPages(){
		try{
			NodeList pageNodes;
			pageNodes = (NodeList)xPath.compile("//reportPages/page/@name").evaluate(MainWindow.xmlDocument,XPathConstants.NODESET);
			
			for(int i = 0; i < pageNodes.getLength(); i++){
				//System.out.println(pageNodes.item(i).getTextContent());
			}
		}
		catch(Exception e){
			
		}
	}
	
	private void extractReportQueries(){
		try{
			NodeList queryNodes;
			queryNodes = (NodeList)xPath.compile("/report/queries/query").evaluate(MainWindow.xmlDocument,XPathConstants.NODESET);
			
			queries = new Query[queryNodes.getLength()];
			for(int i = 0; i < queryNodes.getLength(); i++){
				Node queryNode = queryNodes.item(i);
				Query query = new Query();
				query.queryName = queryNode.getAttributes().getNamedItem("name").getTextContent();
				NodeList dataItemNodes;
				dataItemNodes = (NodeList) xPath.compile("/report/queries/query[@name=\""+
						query.queryName + "\"]/selection/dataItem").evaluate(MainWindow.xmlDocument,XPathConstants.NODESET);
				query.dataItems = new DataItem[dataItemNodes.getLength()];
				for (int j = 0; j < dataItemNodes.getLength(); j++){
					Node dataItemNode = dataItemNodes.item(j);
					DataItem dataItem = new DataItem();
					dataItem.dataItem = dataItemNode.getAttributes().getNamedItem("name").getTextContent();
					NodeList childNodes = dataItemNode.getChildNodes();
					dataItem.nameInPackage = "";
					for (int k = 0; k < childNodes.getLength();k++)
						if  (childNodes.item(k).getNodeName() == "expression")
							dataItem.nameInPackage = childNodes.item(k).getTextContent();
					query.dataItems[j]=dataItem;
				}
				NodeList filterNodes = (NodeList) xPath.compile("/report/queries/query[@name=\""+
						query.queryName + "\"]/detailFilters/detailFilter").evaluate(MainWindow.xmlDocument,XPathConstants.NODESET);
				query.filters = new String[filterNodes.getLength()];
				for (int j = 0; j < filterNodes.getLength(); j++){
					Node filterNode = filterNodes.item(j);
					NodeList childNodes = filterNode.getChildNodes();
					for (int k = 0; k < childNodes.getLength();k++)
						if  (childNodes.item(k).getNodeName() == "filterExpression")
							query.filters[j]=childNodes.item(k).getTextContent();
				}
				queries[i] = query;
			}
			exportedOutput.writeQueryDetails(queries);
		}
		catch(Exception e){
			
		}
	}

}
