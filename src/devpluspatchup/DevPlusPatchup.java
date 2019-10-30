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
import java.awt.Dimension;
import java.awt.Toolkit;
import java.io.File;
import javax.swing.JFileChooser;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author Patel
 */
public class DevPlusPatchup extends javax.swing.JFrame {

	/**
	 * 
	 */
	private String testType;
	int testPriorityCol;
	int testNameCol;
	int testDescCol;
	int testStepCol;
	int testStepDescCol;
	int testExpectCol;
	int testPreCondCol;
	int testStatusCol;
	int testTypeCol;
	int testRunModeCol;
	int testTestTypeCol;
	int testModuleCol;
	private final String testCase = "Test Case";
	private final String testScenerio = "Test Scenerio";
	ArrayList<ArrayList<String>> aas = new ArrayList<>();
	boolean isXLS;
	Workbook workbook;
	int rowNo = 0;
	ArrayList<Template> templateList = new ArrayList<>();

	/**
	 * Creates new form MainForm
	 */
	public DevPlusPatchup() {
		initComponents();

		Dimension dmnsn = Toolkit.getDefaultToolkit().getScreenSize();
		setBounds(dmnsn.width / 2 - getSize().width / 2, dmnsn.height / 2
				- getSize().height / 2, getSize().width, getSize().height);

		buttonGroup1.add(testCaseRadioButton);
		buttonGroup1.add(testScenerioRadioButton);
		testCaseRadioButton.setSelected(true);
		testType = testCase;
	}

	/**
	 * This method is called from within the constructor to initialize the form.
	 * WARNING: Do NOT modify this code.
	 */
	@SuppressWarnings("unchecked")
	// <editor-fold defaultstate="collapsed" desc="Generated
	// Code">//GEN-BEGIN:initComponents
	private void initComponents() {

		buttonGroup1 = new javax.swing.ButtonGroup();
		jPanel1 = new javax.swing.JPanel();
		jLabel1 = new javax.swing.JLabel();
		openFileTextField = new javax.swing.JTextField();
		inputFileButton = new javax.swing.JButton();
		convertButton = new javax.swing.JButton();
		jLabel2 = new javax.swing.JLabel();
		saveFileTextField = new javax.swing.JTextField();
		outputFileButton = new javax.swing.JButton();
		jLabel3 = new javax.swing.JLabel();
		sheetNoTextField = new javax.swing.JTextField();
		testCaseRadioButton = new javax.swing.JRadioButton();
		testScenerioRadioButton = new javax.swing.JRadioButton();
		teamNameTextField = new javax.swing.JTextField();
		jLabel4 = new javax.swing.JLabel();
		jLabel5 = new javax.swing.JLabel();
		runModeTextField = new javax.swing.JTextField();

		setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
		setTitle("DevPlus Patchup");
		setResizable(false);

		jPanel1.setBackground(new java.awt.Color(255, 255, 255));
		jPanel1.setBorder(javax.swing.BorderFactory.createEtchedBorder(
				new java.awt.Color(0, 0, 102),
				new java.awt.Color(204, 204, 255)));
		jPanel1.setForeground(new java.awt.Color(0, 0, 204));
		jPanel1.setFont(new java.awt.Font("Times New Roman", 0, 11)); // NOI18N

		jLabel1.setText("Select Test Scenerio/Test Case file");

		inputFileButton.setText("Browse");
		inputFileButton.addActionListener(new java.awt.event.ActionListener() {
			public void actionPerformed(java.awt.event.ActionEvent evt) {
				inputFileButtonActionPerformed(evt);
			}
		});

		convertButton.setText("Convert to DevPlus Template");
		convertButton.addActionListener(new java.awt.event.ActionListener() {
			public void actionPerformed(java.awt.event.ActionEvent evt) {
				convertButtonActionPerformed(evt);
			}
		});

		jLabel2.setText("Select output folder");

		outputFileButton.setText("Browse");
		outputFileButton.addActionListener(new java.awt.event.ActionListener() {
			public void actionPerformed(java.awt.event.ActionEvent evt) {
				outputFileButtonActionPerformed(evt);
			}
		});

		jLabel3.setText("Select Sheet Number");

		sheetNoTextField.setText("1");

		testCaseRadioButton.setText("Test Case");

		testScenerioRadioButton.setText("Test Scenerio");

		jLabel4.setText("Team Name");

		jLabel5.setText("Run Mode");

		runModeTextField.setText("Manual");

		javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(
				jPanel1);
		jPanel1.setLayout(jPanel1Layout);
		jPanel1Layout
				.setHorizontalGroup(jPanel1Layout
						.createParallelGroup(
								javax.swing.GroupLayout.Alignment.LEADING)
						.addGroup(
								jPanel1Layout
										.createSequentialGroup()
										.addGroup(
												jPanel1Layout
														.createParallelGroup(
																javax.swing.GroupLayout.Alignment.LEADING)
														.addGroup(
																jPanel1Layout
																		.createSequentialGroup()
																		.addContainerGap()
																		.addComponent(
																				saveFileTextField,
																				javax.swing.GroupLayout.PREFERRED_SIZE,
																				398,
																				javax.swing.GroupLayout.PREFERRED_SIZE)
																		.addPreferredGap(
																				javax.swing.LayoutStyle.ComponentPlacement.RELATED)
																		.addComponent(
																				outputFileButton))
														.addGroup(
																jPanel1Layout
																		.createSequentialGroup()
																		.addContainerGap()
																		.addComponent(
																				openFileTextField,
																				javax.swing.GroupLayout.PREFERRED_SIZE,
																				398,
																				javax.swing.GroupLayout.PREFERRED_SIZE)
																		.addPreferredGap(
																				javax.swing.LayoutStyle.ComponentPlacement.RELATED)
																		.addComponent(
																				inputFileButton))
														.addGroup(
																jPanel1Layout
																		.createSequentialGroup()
																		.addGap(49,
																				49,
																				49)
																		.addComponent(
																				jLabel1))
														.addGroup(
																jPanel1Layout
																		.createSequentialGroup()
																		.addGap(43,
																				43,
																				43)
																		.addGroup(
																				jPanel1Layout
																						.createParallelGroup(
																								javax.swing.GroupLayout.Alignment.LEADING)
																						.addGroup(
																								jPanel1Layout
																										.createSequentialGroup()
																										.addComponent(
																												jLabel5)
																										.addPreferredGap(
																												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
																										.addComponent(
																												runModeTextField,
																												javax.swing.GroupLayout.PREFERRED_SIZE,
																												123,
																												javax.swing.GroupLayout.PREFERRED_SIZE))
																						.addGroup(
																								jPanel1Layout
																										.createSequentialGroup()
																										.addComponent(
																												jLabel3)
																										.addGap(4,
																												4,
																												4)
																										.addComponent(
																												sheetNoTextField,
																												javax.swing.GroupLayout.PREFERRED_SIZE,
																												39,
																												javax.swing.GroupLayout.PREFERRED_SIZE)
																										.addGap(41,
																												41,
																												41)
																										.addComponent(
																												testCaseRadioButton)
																										.addGap(18,
																												18,
																												18)
																										.addComponent(
																												testScenerioRadioButton))
																						.addGroup(
																								jPanel1Layout
																										.createSequentialGroup()
																										.addComponent(
																												jLabel4)
																										.addPreferredGap(
																												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
																										.addComponent(
																												teamNameTextField,
																												javax.swing.GroupLayout.PREFERRED_SIZE,
																												241,
																												javax.swing.GroupLayout.PREFERRED_SIZE))
																						.addComponent(
																								jLabel2)))
														.addGroup(
																jPanel1Layout
																		.createSequentialGroup()
																		.addGap(152,
																				152,
																				152)
																		.addComponent(
																				convertButton)))
										.addContainerGap(
												javax.swing.GroupLayout.DEFAULT_SIZE,
												Short.MAX_VALUE)));
		jPanel1Layout
				.setVerticalGroup(jPanel1Layout
						.createParallelGroup(
								javax.swing.GroupLayout.Alignment.LEADING)
						.addGroup(
								javax.swing.GroupLayout.Alignment.TRAILING,
								jPanel1Layout
										.createSequentialGroup()
										.addContainerGap(21, Short.MAX_VALUE)
										.addComponent(jLabel1)
										.addPreferredGap(
												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
										.addGroup(
												jPanel1Layout
														.createParallelGroup(
																javax.swing.GroupLayout.Alignment.BASELINE)
														.addComponent(
																openFileTextField,
																javax.swing.GroupLayout.PREFERRED_SIZE,
																javax.swing.GroupLayout.DEFAULT_SIZE,
																javax.swing.GroupLayout.PREFERRED_SIZE)
														.addComponent(
																inputFileButton))
										.addGap(5, 5, 5)
										.addGroup(
												jPanel1Layout
														.createParallelGroup(
																javax.swing.GroupLayout.Alignment.BASELINE)
														.addComponent(jLabel3)
														.addComponent(
																sheetNoTextField,
																javax.swing.GroupLayout.PREFERRED_SIZE,
																javax.swing.GroupLayout.DEFAULT_SIZE,
																javax.swing.GroupLayout.PREFERRED_SIZE)
														.addComponent(
																testCaseRadioButton)
														.addComponent(
																testScenerioRadioButton))
										.addPreferredGap(
												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
										.addGroup(
												jPanel1Layout
														.createParallelGroup(
																javax.swing.GroupLayout.Alignment.BASELINE)
														.addComponent(
																teamNameTextField,
																javax.swing.GroupLayout.PREFERRED_SIZE,
																javax.swing.GroupLayout.DEFAULT_SIZE,
																javax.swing.GroupLayout.PREFERRED_SIZE)
														.addComponent(jLabel4))
										.addPreferredGap(
												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
										.addGroup(
												jPanel1Layout
														.createParallelGroup(
																javax.swing.GroupLayout.Alignment.BASELINE)
														.addComponent(jLabel5)
														.addComponent(
																runModeTextField,
																javax.swing.GroupLayout.PREFERRED_SIZE,
																javax.swing.GroupLayout.DEFAULT_SIZE,
																javax.swing.GroupLayout.PREFERRED_SIZE))
										.addPreferredGap(
												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
										.addComponent(jLabel2)
										.addGap(3, 3, 3)
										.addGroup(
												jPanel1Layout
														.createParallelGroup(
																javax.swing.GroupLayout.Alignment.BASELINE)
														.addComponent(
																saveFileTextField,
																javax.swing.GroupLayout.PREFERRED_SIZE,
																javax.swing.GroupLayout.DEFAULT_SIZE,
																javax.swing.GroupLayout.PREFERRED_SIZE)
														.addComponent(
																outputFileButton))
										.addPreferredGap(
												javax.swing.LayoutStyle.ComponentPlacement.RELATED)
										.addComponent(convertButton)
										.addGap(5, 5, 5)));

		javax.swing.GroupLayout layout = new javax.swing.GroupLayout(
				getContentPane());
		getContentPane().setLayout(layout);
		layout.setHorizontalGroup(layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGroup(
				layout.createSequentialGroup()
						.addComponent(jPanel1,
								javax.swing.GroupLayout.PREFERRED_SIZE,
								javax.swing.GroupLayout.DEFAULT_SIZE,
								javax.swing.GroupLayout.PREFERRED_SIZE)
						.addGap(0, 0, Short.MAX_VALUE)));
		layout.setVerticalGroup(layout.createParallelGroup(
				javax.swing.GroupLayout.Alignment.LEADING).addGroup(
				layout.createSequentialGroup()
						.addComponent(jPanel1,
								javax.swing.GroupLayout.PREFERRED_SIZE,
								javax.swing.GroupLayout.DEFAULT_SIZE,
								javax.swing.GroupLayout.PREFERRED_SIZE)
						.addGap(0, 0, Short.MAX_VALUE)));

		pack();
		setLocationRelativeTo(null);
	}// </editor-fold>//GEN-END:initComponents

	private void inputFileButtonActionPerformed(java.awt.event.ActionEvent evt) {// GEN-FIRST:event_inputFileButtonActionPerformed
		JFileChooser fileChooser = new JFileChooser();
		int result = fileChooser.showOpenDialog(this);
		if (result == JFileChooser.APPROVE_OPTION) {
			openFileTextField.setText(fileChooser.getSelectedFile()
					.getAbsolutePath());
		}
	}// GEN-LAST:event_inputFileButtonActionPerformed

	private void outputFileButtonActionPerformed(java.awt.event.ActionEvent evt) {// GEN-FIRST:event_outputFileButtonActionPerformed
		JFileChooser fileChooser = new JFileChooser();
		int result = fileChooser.showSaveDialog(this);
		if (result == JFileChooser.APPROVE_OPTION) {
			saveFileTextField.setText(fileChooser.getSelectedFile()
					.getAbsolutePath());
		}
	}// GEN-LAST:event_outputFileButtonActionPerformed

	private void convertButtonActionPerformed(java.awt.event.ActionEvent evt) {// GEN-FIRST:event_convertButtonActionPerformed

		try {
			String s = saveFileTextField.getText();
			String o = openFileTextField.getText();
			if (!s.isEmpty() && !o.isEmpty()) {
				String path;

				if (testCaseRadioButton.isSelected()) {
					testType = testCase;
				} else if (testScenerioRadioButton.isSelected()) {
					testType = testScenerio;
				}

				int c = processReadingFile();

				if (isXLS) {
					path = s.contains(".xls") ? s : s + ".xls";
				} else {
					path = s.contains(".xlsx") ? s : s + ".xlsx";
				}

				if (c == 0) {
					writeFile(path);
				}
			} else {
				JOptionPane.showMessageDialog(this,
						"Please select input/output file path!");
			}
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, e.getMessage());
			Logger.getLogger(DevPlusPatchup.class.getName()).log(Level.SEVERE,
					null, e);
		} finally {
			testType = testCase;
			testPriorityCol = 0;
			testNameCol = 0;
			testDescCol = 0;
			testStepCol = 0;
			testStepDescCol = 0;
			testExpectCol = 0;
			testPreCondCol = 0;
			testStatusCol = 0;
			testTypeCol = 0;
			testRunModeCol = 0;
			testTestTypeCol = 0;
			testModuleCol = 0;
			aas = new ArrayList<>();
			rowNo = 0;
			templateList = new ArrayList<>();
			workbook = null;
			writeWorkbook = null;
		}
	}// GEN-LAST:event_convertButtonActionPerformed

	/**
	 * @param args
	 *            the command line arguments
	 */
	public static void main(String args[]) {
		/* Set the Nimbus look and feel */
		// <editor-fold defaultstate="collapsed" desc=" Look and feel setting
		// code
		// (optional) ">
		/*
		 * If Nimbus (introduced in Java SE 6) is not available, stay with the
		 * default look and feel. For details see
		 * http://download.oracle.com/javase
		 * /tutorial/uiswing/lookandfeel/plaf.html
		 */
		try {
			for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager
					.getInstalledLookAndFeels()) {
				if ("Nimbus".equals(info.getName())) {
					javax.swing.UIManager.setLookAndFeel(info.getClassName());
					break;
				}
			}
		} catch (ClassNotFoundException ex) {
			java.util.logging.Logger.getLogger(DevPlusPatchup.class.getName())
					.log(java.util.logging.Level.SEVERE, null, ex);
		} catch (InstantiationException ex) {
			java.util.logging.Logger.getLogger(DevPlusPatchup.class.getName())
					.log(java.util.logging.Level.SEVERE, null, ex);
		} catch (IllegalAccessException ex) {
			java.util.logging.Logger.getLogger(DevPlusPatchup.class.getName())
					.log(java.util.logging.Level.SEVERE, null, ex);
		} catch (javax.swing.UnsupportedLookAndFeelException ex) {
			java.util.logging.Logger.getLogger(DevPlusPatchup.class.getName())
					.log(java.util.logging.Level.SEVERE, null, ex);
		}
		// </editor-fold>

		/* Create and display the form */
		java.awt.EventQueue.invokeLater(new Runnable() {
			public void run() {
				new DevPlusPatchup().setVisible(true);
			}
		});
	}

	// Variables declaration - do not modify//GEN-BEGIN:variables
	private javax.swing.ButtonGroup buttonGroup1;
	private javax.swing.JButton convertButton;
	private javax.swing.JButton inputFileButton;
	private javax.swing.JLabel jLabel1;
	private javax.swing.JLabel jLabel2;
	private javax.swing.JLabel jLabel3;
	private javax.swing.JLabel jLabel4;
	private javax.swing.JLabel jLabel5;
	private javax.swing.JPanel jPanel1;
	private javax.swing.JTextField openFileTextField;
	private javax.swing.JButton outputFileButton;
	private javax.swing.JTextField runModeTextField;
	private javax.swing.JTextField saveFileTextField;
	private javax.swing.JTextField sheetNoTextField;
	private javax.swing.JTextField teamNameTextField;
	private javax.swing.JRadioButton testCaseRadioButton;
	private javax.swing.JRadioButton testScenerioRadioButton;

	// End of variables declaration//GEN-END:variables

	private int processReadingFile() {

		FileInputStream inputStream = null;
		String s = openFileTextField.getText();
		try {
			inputStream = new FileInputStream(new File(s));
		} catch (FileNotFoundException ex) {
			Logger.getLogger(DevPlusPatchup.class.getName()).log(Level.SEVERE,
					null, ex);
		}

		if (s.substring(s.lastIndexOf("."), s.length())
				.equalsIgnoreCase(".xls")) {

			try {
				workbook = new HSSFWorkbook(new FileInputStream(s));
				isXLS = true;
			} catch (Exception ex) {
				Logger.getLogger(DevPlusPatchup.class.getName()).log(
						Level.SEVERE, null, ex);
				JOptionPane.showMessageDialog(this, "Enter valid file");
			}

		} else if (s.substring(s.lastIndexOf("."), s.length())
				.equalsIgnoreCase(".xlsx")) {

			try {
				workbook = new XSSFWorkbook(inputStream);
				isXLS = false;
			} catch (Exception ex) {
				Logger.getLogger(DevPlusPatchup.class.getName()).log(
						Level.SEVERE, null, ex);
				JOptionPane.showMessageDialog(this, "Enter valid file");
			}
		} else {
			JOptionPane
					.showMessageDialog(this,
							"Enter valid file format.\n Only .xls and .xlsx are accepted.");
		}
		Sheet sheet = workbook.getSheetAt(Integer.parseInt(sheetNoTextField
				.getText()));

		for (int i = 0; i < sheet.getLastRowNum() + 25; i++) {
			Row nextRow = sheet.getRow(i);
			ArrayList<String> as = new ArrayList<>();
			as.add("");
			int n = 25;
			try {
				n = nextRow.getLastCellNum();
			} catch (NullPointerException e) {
				n = 25;
				Logger.getLogger(DevPlusPatchup.class.getName()).log(
						Level.SEVERE, null, e);
			}
			for (int j = 0; j < n; j++) {
				String cellValue;
				try {
					Cell cell = nextRow.getCell(j);
					cellValue = cell.toString();
				} catch (NullPointerException npe) {
					cellValue = "";
				}
				as.add(cellValue);
			}
			aas.add(as);
			System.out.println();
		}
		// printTable();
		int c = processData();
		if (c == 0) {
			prepareData();
		}
		try {
			inputStream.close();
		} catch (IOException ex) {
			Logger.getLogger(DevPlusPatchup.class.getName()).log(Level.SEVERE,
					null, ex);
		}
		return c;
	}

	/*
	 * private void printTable() { System.out.println("File contents:"); for
	 * (ArrayList<String> as : aas) { for (String s : as) { System.out.print(s +
	 * "-"); } System.out.println(); } System.out.println(
	 * "----------------------------------------------------------------------------"
	 * ); }
	 */
	private int processData() {
		int j = 0;
		boolean nameColDone = false;
		boolean descColDone = false;
		boolean stepColDone = false;
		boolean stepDescColDone = false;
		boolean expectColDone = false;

		for (ArrayList<String> as : aas) {
			int i = 0;

			for (String s : as) {

				if (!s.isEmpty()) {
					switch (s.toLowerCase()) {
					case "test name":
					case "testname":
					case "id":
					case "testcaseid":
					case "test case id":
					case "testcasename":
					case "test case name":
					case "test id":
					case "testid":
					case "test scenario id":
						if(!nameColDone) {
						testNameCol = i;
						nameColDone = true;
						}
						break;
					case "test case description":
					case "test scenario description":
					case "test description":						
						testDescCol = i;
						descColDone = true;
						break;
					case "step name":
					case "step no":
					case "step no.":
					case "step-no":
					case "step-name":
						testStepCol = i;
						stepColDone = true;
						break;
					case "step description":
						testStepDescCol = i;
						stepDescColDone = true;
						break;
					case "expected result":
						testExpectCol = i;
						expectColDone = true;
						break;
					case "pre-condition":
					case "pre condition":
					case "pre-conditions":
					case "pre conditions":
						testPreCondCol = i;
						break;
					case "status":
					case "test status":
						testStatusCol = i;
						break;
					case "test type":
						testTypeCol = i;
						break;
					case "run mode":
						testRunModeCol = i;
						break;
					case "test priority":
					case "priority":
						testPriorityCol = i;
						break;
					case "application module":
					case "module":
						testModuleCol = i;
						break;
					}
				}
				i++;
			}
			if (testType.equals(testCase)) {
				if (nameColDone && descColDone && stepColDone
						&& stepDescColDone && expectColDone) {
					rowNo = j;
					return 0;
				} else if (nameColDone || descColDone || stepColDone
						|| stepDescColDone || expectColDone) {
					JOptionPane.showMessageDialog(this,
							"Prepare all the Necessary Columns!\n"
									+ "Columns Present : "
									+ "\ntest case id/name : " + nameColDone
									+ "\ntest case description : "
									+ descColDone + "\nstep name : "
									+ stepColDone + "\nstep description : "
									+ stepDescColDone + "\nexpected result : "
									+ expectColDone);
					return -1;
				}
			} else if (testType.equals(testScenerio)) {
				if (nameColDone && descColDone) {
					rowNo = j;
					return 0;
				} else if (nameColDone || descColDone) {
					JOptionPane.showMessageDialog(this,
							"Prepare all the Necessary Columns!\n"
									+ "Columns Present : "
									+ "\ntest scenerio id/name : "
									+ nameColDone
									+ "\ntest scenerio description : "
									+ descColDone);
					return -1;
				}
			}
			j++;
		}
		return -1;
	}

	private void prepareData() {

		Template t = new Template("Test Name", "Test Description", "Step Name",
				"Description", "Expected Result", "Test Status",
				"Test Pre-Requisites", "Test Priority", "Module");

		templateList.add(t);
		try {
			for (int i = rowNo + 1; i < aas.size(); i++) {
				ArrayList<String> as = aas.get(i);

				String tmpTestName = allocateString(as, testNameCol);
				String tmpTestDesc = allocateString(as, testDescCol);
				String tmpTestStep = allocateString(as, testStepCol);
				String tmpTestStepDesc = allocateString(as, testStepDescCol);
				String tmpTestExpect = allocateString(as, testExpectCol);
				String tmpTestStatus = allocateString(as, testStatusCol);
				String tmpTestPreCond = allocateString(as, testPreCondCol);
				String tmpTestPriority = allocateString(as, testPriorityCol);
				String tmpTestModule = allocateString(as, testModuleCol);

				templateList.add(new Template(tmpTestName, tmpTestDesc,
						tmpTestStep, tmpTestStepDesc, tmpTestExpect,
						tmpTestStatus, tmpTestPreCond, tmpTestPriority,
						tmpTestModule));
			}
		} catch (Exception ex) {
			Logger.getLogger(DevPlusPatchup.class.getName()).log(Level.SEVERE,
					null, ex);
		}
		// printPreparedData();
	}

	private String allocateString(ArrayList<String> as, int colNum) {
		return as.size() > colNum ? as.get(colNum) : "";
	}

	/*
	 * private void printPreparedData() { for (Template t : templateList) {
	 * System.out.println(t.toString()); } }
	 */
	Workbook writeWorkbook;

	private void writeFile(String FILE_PATH) {

		if (isXLS) {
			writeWorkbook = new HSSFWorkbook();
		} else {
			writeWorkbook = new XSSFWorkbook();
		}

		Sheet studentsSheet = writeWorkbook.createSheet("Tests");

		int rowIndex = 3;
		boolean isFirst = true;

		for (Template tmp : templateList) {
			Row row = studentsSheet.createRow(rowIndex++);

			row.createCell(9).setCellValue(tmp.testName);
			row.createCell(10).setCellValue(tmp.testPreCond);
			row.createCell(13).setCellValue(tmp.testDesc);
			row.createCell(14).setCellValue(tmp.testModule);
			row.createCell(16).setCellValue(tmp.testStep);
			row.createCell(17).setCellValue(tmp.testStepDesc);
			row.createCell(18).setCellValue(tmp.testExpect);

			if (!tmp.testName.isEmpty()) {

				if (tmp.testStatus.isEmpty()) {
					row.createCell(4).setCellValue("New");
				} else {
					row.createCell(4).setCellValue(tmp.testStatus);
				}

				if (tmp.testPriority.isEmpty()) {
					row.createCell(7).setCellValue("High");
				} else {
					row.createCell(7).setCellValue(tmp.testPriority);
				}

				if (isFirst) {
					row.createCell(2).setCellValue("Subject");
					row.createCell(3).setCellValue("Test Type");
					row.createCell(5).setCellValue("Run Mode");
					row.createCell(6).setCellValue("Automation Engine");
					row.createCell(8).setCellValue("Team Name");
					row.createCell(11).setCellValue("Test Post-Conditions");
					row.createCell(12).setCellValue("Tags");
					row.createCell(15).setCellValue("Call Case ID");
					row.createCell(19).setCellValue("Parameters");
					row.createCell(20).setCellValue("Requirement Coverage");
					row.createCell(21).setCellValue("Pre-Conditions");
					row.createCell(22).setCellValue("Reverse-Pre-Conditions");
					row.createCell(23).setCellValue("Linked Defects");
					row.createCell(24).setCellValue("Linked Wiki");
					isFirst = false;
				} else {

					row.createCell(3).setCellValue(testType);
					row.createCell(5).setCellValue(runModeTextField.getText());
					row.createCell(8).setCellValue(teamNameTextField.getText());
				}
			}
		}

		// write this writeWorkbook in excel file.
		try {
			FileOutputStream fos = new FileOutputStream(FILE_PATH);
			writeWorkbook.write(fos);
			fos.close();

			JOptionPane.showMessageDialog(null, FILE_PATH
					+ " is successfully written");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
