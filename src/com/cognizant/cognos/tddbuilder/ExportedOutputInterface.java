package com.cognizant.cognos.tddbuilder;

public interface ExportedOutputInterface {
	public void writePromptDetails(Prompt prompts[]);
	public void writeConditionalVariableDetails(ConditionalVariable conditionalVariable[]);
	public void writeReportPageDetails(ReportPage reportPage[]);
	public String writeToFile();

}
