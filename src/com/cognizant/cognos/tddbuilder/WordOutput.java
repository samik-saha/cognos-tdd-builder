package com.cognizant.cognos.tddbuilder;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;


public class WordOutput implements ExportedOutputInterface{
	MainWindow mainWindow;
	XWPFDocument doc;
	String filename;
	XWPFNumbering numbering;
	
	public WordOutput(MainWindow mainWindow, String filename) {
		this.mainWindow=mainWindow;
		this.filename=filename;
		
		try {
			doc = new XWPFDocument(new FileInputStream("template.docx"));
		}
		catch(IOException ex){
			System.err.println(ex.getMessage());
			doc = new XWPFDocument();
		}
		
		numbering = doc.createNumbering();
	}

	public void createSimpleTable() throws Exception {

        XWPFTable table = doc.createTable(3, 3);

        table.getRow(1).getCell(1).setText("EXAMPLE OF TABLE");

        // table cells have a list of paragraphs; there is an initial
        // paragraph created when the cell is created. If you create a
        // paragraph in the document to put in the cell, it will also
        // appear in the document following the table, which is probably
        // not the desired result.
        XWPFParagraph p1 = table.getRow(0).getCell(0).getParagraphs().get(0);

        XWPFRun r1 = p1.createRun();
        r1.setBold(true);
        r1.setText("The quick brown fox");
        r1.setItalic(true);
        r1.setFontFamily("Courier");
        r1.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
        r1.setTextPosition(100);

        table.getRow(2).getCell(2).setText("only text");

        FileOutputStream out = new FileOutputStream("simpleTable.docx");
        doc.write(out);
        out.close();
    }

    /**
     * Create a table with some row and column styling. I "manually" add the
     * style name to the table, but don't check to see if the style actually
     * exists in the document. Since I'm creating it from scratch, it obviously
     * won't exist. When opened in MS Word, the table style becomes "Normal".
     * I manually set alternating row colors. This could be done using Themes,
     * but that's left as an exercise for the reader. The cells in the last
     * column of the table have 10pt. "Courier" font.
     * I make no claims that this is the "right" way to do it, but it worked
     * for me. Given the scarcity of XWPF examples, I thought this may prove
     * instructive and give you ideas for your own solutions.

     * @throws Exception
     */
    public void createStyledTable() throws Exception {

    	// -- OR --
        // open an existing empty document with styles already defined
        //XWPFDocument doc = new XWPFDocument(new FileInputStream("base_document.docx"));

    	// Create a new table with 6 rows and 3 columns
    	int nRows = 6;
    	int nCols = 3;
        XWPFTable table = doc.createTable(nRows, nCols);

        // Set the table style. If the style is not defined, the table style
        // will become "Normal".
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTString styleStr = tblPr.addNewTblStyle();
        styleStr.setVal("StyledTable");

        // Get a list of the rows in the table
        List<XWPFTableRow> rows = table.getRows();
        int rowCt = 0;
        int colCt = 0;
        for (XWPFTableRow row : rows) {
        	// get table row properties (trPr)
        	CTTrPr trPr = row.getCtRow().addNewTrPr();
        	// set row height; units = twentieth of a point, 360 = 0.25"
        	CTHeight ht = trPr.addNewTrHeight();
        	ht.setVal(BigInteger.valueOf(360));

	        // get the cells in this row
        	List<XWPFTableCell> cells = row.getTableCells();
            // add content to each cell
        	for (XWPFTableCell cell : cells) {
        		// get a table cell properties element (tcPr)
        		CTTcPr tcpr = cell.getCTTc().addNewTcPr();
        		// set vertical alignment to "center"
        		CTVerticalJc va = tcpr.addNewVAlign();
        		va.setVal(STVerticalJc.CENTER);

        		// create cell color element
        		CTShd ctshd = tcpr.addNewShd();
                ctshd.setColor("auto");
                ctshd.setVal(STShd.CLEAR);
                if (rowCt == 0) {
                	// header row
                	ctshd.setFill("A7BFDE");
                }
            	else if (rowCt % 2 == 0) {
            		// even row
                	ctshd.setFill("D3DFEE");
            	}
            	else {
            		// odd row
                	ctshd.setFill("EDF2F8");
            	}

                // get 1st paragraph in cell's paragraph list
                XWPFParagraph para = cell.getParagraphs().get(0);
                // create a run to contain the content
                XWPFRun rh = para.createRun();
                // style cell as desired
                if (colCt == nCols - 1) {
                	// last column is 10pt Courier
                	rh.setFontSize(10);
                	rh.setFontFamily("Courier");
                }
                if (rowCt == 0) {
                	// header row
                    rh.setText("header row, col " + colCt);
                	rh.setBold(true);
                    para.setAlignment(ParagraphAlignment.CENTER);
                }
            	else if (rowCt % 2 == 0) {
            		// even row
                    rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
            	}
            	else {
            		// odd row
                    rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
            	}
                colCt++;
        	} // for cell
        	colCt = 0;
        	rowCt++;
        } // for row

        // write the file
        FileOutputStream out = new FileOutputStream("styledTable.docx");
        doc.write(out);
        out.close();
    }




	@Override
	public void writePromptDetails(Prompt[] prompts) {
		XWPFParagraph para;
    	XWPFRun rh;
		
		if(prompts.length>0){
			mainWindow.log("Writing prompt details to document...");
			int nRows = prompts.length + 1;
	    	int nCols = 7;
	    	
	    	para = doc.createParagraph();
	    	para.setStyle("Heading2");
	    	rh = para.createRun();
	    	rh.setText("Report Prompts");
	    	
	        XWPFTable table = doc.createTable(nRows, nCols);

	        // Set the table style. If the style is not defined, the table style
	        // will become "Normal".
	        CTTblPr tblPr = table.getCTTbl().getTblPr();
	        CTString styleStr = tblPr.addNewTblStyle();
	        styleStr.setVal("StyledTable");
	        
	        // Get a list of the rows in the table
	        XWPFTableRow row = table.getRow(0);
	        
	        List<XWPFTableCell> cells = row.getTableCells();
	        for (XWPFTableCell cell : cells) {
	    		// get a table cell properties element (tcPr)
	    		CTTcPr tcpr = cell.getCTTc().addNewTcPr();
	    		// set vertical alignment to "center"
	    		CTVerticalJc va = tcpr.addNewVAlign();
	    		va.setVal(STVerticalJc.CENTER);

	    		// create cell color element
	    		CTShd ctshd = tcpr.addNewShd();
	            ctshd.setColor("auto");
	            ctshd.setVal(STShd.CLEAR);
	            ctshd.setFill("A7BFDE");
	        }
	        
	        /* Write table headers */
	        table.getRow(0).getCell(0).setText("Parameter Name (Case Sensitive)");
	        table.getRow(0).getCell(1).setText("Name in Package");
	        table.getRow(0).getCell(2).setText("Sort");
	        table.getRow(0).getCell(3).setText("Man/Opt");
	        table.getRow(0).getCell(4).setText("Format");
	        table.getRow(0).getCell(5).setText("Test Value");
	        table.getRow(0).getCell(6).setText("Comment (Pre selection, cascade,...)");
	        
	        
	        for (int i=0; i<nRows-1;i++){
	        	table.getRow(i+1).getCell(0).setText(prompts[i].parameter);
	        	table.getRow(i+1).getCell(1).setText(prompts[i].name);
	        	
	        	/* Write sorting details */
	        	// Replace null comments with blank text
	        	prompts[i].sort=prompts[i].sort !=null?prompts[i].sort:"";
	        	// Split text based on newline characters
	        	String[] multilineSort=prompts[i].sort.split("\\r?\\n");
	        	// Remove automatically created default paragraph
	        	table.getRow(i+1).getCell(2).removeParagraph(0);
	        	// Create a paragraph for each line
	        	for (int j=0; j<multilineSort.length; j++){
	        		para = table.getRow(i+1).getCell(2).addParagraph();
	        		rh = para.createRun();
	        		rh.setText(multilineSort[j]);
	        	}
	        	
	        	table.getRow(i+1).getCell(3).setText(prompts[i].required=="true"?"Mandatory":"Optional");
	        	
	        	table.getRow(i+1).getCell(4).setText(prompts[i].format);
	        	
	        	/* write comments */
	        	// Replace null comments with blank text
	        	prompts[i].comment=prompts[i].comment !=null?prompts[i].comment:"";
	        	// Split text based on newline characters
	        	String[] multilineComment=prompts[i].comment.split("\\r?\\n");
	        	// Remove automatically created default paragraph
	        	table.getRow(i+1).getCell(6).removeParagraph(0);
	        	// Create a paragraph for each line
	        	for (int j=0; j<multilineComment.length; j++){
	        		para = table.getRow(i+1).getCell(6).addParagraph();
	        		rh = para.createRun();
	        		rh.setText(multilineComment[j]);
	        	}
	        }
	        mainWindow.log("Prompt details has been written to document.");
		}
		else{
			mainWindow.log("No prompt details is wrriten to document.");
		}
		
		
	}
	
	
	
	
	
	@Override
	public String writeToFile(){
        // write the file
        FileOutputStream out;
		try {
			mainWindow.log("Creating MS Word Document " + filename);
			out = new FileOutputStream(filename);
	        doc.write(out);
	        out.close();
	        return filename;
			
		} catch (IOException e) {
			e.printStackTrace();
			mainWindow.log("Error writing to file "+filename);
			JOptionPane.showMessageDialog(mainWindow.getFrame(),
				    "Could not write to the file "+filename+"!",
				    "Error",
				    JOptionPane.ERROR_MESSAGE);
			
			return null;
		} 

	}

	@Override
	public void writeConditionalVariableDetails(
			ConditionalVariable[] conditionalVariables) {
		XWPFParagraph para;
    	XWPFRun rh;
		
		if(conditionalVariables.length>0){
			mainWindow.log("Writing report variable details to document...");
			int nRows = conditionalVariables.length + 1;
	    	int nCols = 5;
	    	
	    	doc.createParagraph();
	    	para = doc.createParagraph();
	    	para.setStyle("Heading2");
	    	rh = para.createRun();
	    	rh.setText("Conditional Variables");
	    	
	        XWPFTable table = doc.createTable(nRows, nCols);

	        // Set the table style. If the style is not defined, the table style
	        // will become "Normal".
	        CTTblPr tblPr = table.getCTTbl().getTblPr();
	        CTString styleStr = tblPr.addNewTblStyle();
	        styleStr.setVal("StyledTable");
	        
	        // Get a list of the rows in the table
	        XWPFTableRow row = table.getRow(0);
	        
	        List<XWPFTableCell> cells = row.getTableCells();
	        for (XWPFTableCell cell : cells) {
	    		// get a table cell properties element (tcPr)
	    		CTTcPr tcpr = cell.getCTTc().addNewTcPr();
	    		// set vertical alignment to "center"
	    		CTVerticalJc va = tcpr.addNewVAlign();
	    		va.setVal(STVerticalJc.CENTER);

	    		// create cell color element
	    		CTShd ctshd = tcpr.addNewShd();
	            ctshd.setColor("auto");
	            ctshd.setVal(STShd.CLEAR);
	            ctshd.setFill("A7BFDE");
	        }
	        
	        /* Write table headers */
	        table.getRow(0).getCell(0).setText("Name");
	        table.getRow(0).getCell(1).setText("Type");
	        table.getRow(0).getCell(2).setText("Logic");
	        table.getRow(0).getCell(3).setText("Values");
	        table.getRow(0).getCell(4).setText("Comment");
	        
	        for (int i=0; i<nRows-1;i++){
	        	table.getRow(i+1).getCell(0).setText(conditionalVariables[i].name);
	        	table.getRow(i+1).getCell(1).setText(conditionalVariables[i].type);
	        	table.getRow(i+1).getCell(2).setText(conditionalVariables[i].logic);
	        	table.getRow(i+1).getCell(3).setText(conditionalVariables[i].values);
	        	
	        }
	        mainWindow.log("Report variable details has been written to document.");
		}
		else{
			mainWindow.log("No report variable details is wrriten to document.");
		}
		
	}

	@Override
	public void writeReportPageDetails(ReportPage[] reportPage) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void writeQueryDetails(Query[] queries) {
		XWPFParagraph para;
    	XWPFRun rh;
    	doc.createParagraph();
    	para = doc.createParagraph();
    	para.setStyle("Heading2");
    	rh = para.createRun();
    	rh.setText("Report Queries");
		
		for(int i = 0; i < queries.length; i++){
			writeSingleQueryDetails(queries[i]);
		}	
	}
	
	private void writeSingleQueryDetails(Query query){
		XWPFParagraph para;
    	XWPFRun rh;
    	
		if(query.dataItems.length>0){
			mainWindow.log("Writing query details to document for "+query.queryName+"...");
			int nRows = query.dataItems.length + 1;
	    	int nCols = 3;
	    	
	    	doc.createParagraph();
	    	para = doc.createParagraph();
	    	para.setStyle("Heading3");
	    	rh = para.createRun();
	    	rh.setText("Query Name: " + query.queryName);
	    	
	        XWPFTable table = doc.createTable(nRows, nCols);

	        // Set the table style. If the style is not defined, the table style
	        // will become "Normal".
	        CTTblPr tblPr = table.getCTTbl().getTblPr();
	        CTString styleStr = tblPr.addNewTblStyle();
	        styleStr.setVal("StyledTable");
	        
	        // Get a list of the rows in the table
	        XWPFTableRow row = table.getRow(0);
	        
	        List<XWPFTableCell> cells = row.getTableCells();
	        for (XWPFTableCell cell : cells) {
	    		// get a table cell properties element (tcPr)
	    		CTTcPr tcpr = cell.getCTTc().addNewTcPr();
	    		// set vertical alignment to "center"
	    		CTVerticalJc va = tcpr.addNewVAlign();
	    		va.setVal(STVerticalJc.CENTER);

	    		// create cell color element
	    		CTShd ctshd = tcpr.addNewShd();
	            ctshd.setColor("auto");
	            ctshd.setVal(STShd.CLEAR);
	            ctshd.setFill("A7BFDE");
	        }
	        
	        /* Write table headers */
	        table.getRow(0).getCell(0).setText("DataItem");
	        table.getRow(0).getCell(1).setText("Name in Package");
	        table.getRow(0).getCell(2).setText("Comment");
	        
	        for (int i=0; i<nRows-1;i++){
	        	table.getRow(i+1).getCell(0).setText(query.dataItems[i].dataItem);
	        	table.getRow(i+1).getCell(1).setText(query.dataItems[i].nameInPackage);
	        }
	        
	        /* Write query filters */
	        
	        if (query.filters.length > 0){
	        	para = doc.createParagraph();
		    	rh = para.createRun();
		    	rh.setBold(true);
		    	rh.setText("Filters:");
		    	
		    	
		        for (int i = 0; i < query.filters.length; i++){
		        	para = doc.createParagraph();
		        	rh = para.createRun();
		        	rh.setText(query.filters[i]);
		        }
	        }
	        
	        mainWindow.log("Query details for " + query.queryName + " has been written to document.");
		}
		else{
			mainWindow.log("Query " + query.queryName + "does not have any data item.");
		}
	}

}


