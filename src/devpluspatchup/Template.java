/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package devpluspatchup;

/**
 *
 * @author patel
 */
public class Template {

	String testName;
	String testDesc;
	String testStep;
	String testStepDesc;
	String testExpect;
	String testStatus;
	String testPreCond;
	String testPriority;
	String testModule;

	Template(String testNameCol, String testDescCol, String testStepCol, String testStepDescCol,
			String testExpectCol, String testStatusCol, String testPreCondCol, String testPriority, String testModule) {
		super();
		this.testName = testNameCol;
		this.testDesc = testDescCol;
		this.testStep = testStepCol;
		this.testStepDesc = testStepDescCol;
		this.testExpect = testExpectCol;
		this.testStatus = testStatusCol;
		this.testPreCond = testPreCondCol;
		this.testPriority = testPriority;
		this.testModule = testModule;
	}

}
