/**
 * Owner: ShaownS
 * File: SpreadSheetReader.java
 * Package: org.fusecensus.poireader
 * Project: FUSECensus
 * Email: ssarker@ncsu.edu
 */
package org.fusecensus.poireader;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.PushbackInputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * 
 */
public class SpreadSheetReader {
	
	public SpreadSheetReader() {
		
	}
	
	private static Workbook createWorkBook(InputStream inputStreamForWorkbook) throws IOException, InvalidFormatException {

		if (POIFSFileSystem.hasPOIFSHeader(inputStreamForWorkbook)) {
			NPOIFSFileSystem nfs = new NPOIFSFileSystem(inputStreamForWorkbook);
		    return WorkbookFactory.create(nfs);
		}
		return WorkbookFactory.create(inputStreamForWorkbook);
	}
	
	private static void addFormulaStat(Workbook wb, String fileName, Cell c) {
		// Get the first sheet.
		Sheet s = wb.getSheetAt(0);
		
		// Should not be called for anything other than formula cell.
		if (c.getCellType() != Cell.CELL_TYPE_FORMULA) {
			return;
		}
		
		// Get the last row number in the sheet.
		int lastRowNum = s.getLastRowNum();
		
		// Append a row below that.
		Row addedRow = s.createRow(lastRowNum + 1);
		
		// Add file name.
		addedRow.createCell(0).setCellValue(fileName);
		
		// Add cell location.
		CellReference currentCellRef = new CellReference(c);
		addedRow.createCell(1).setCellValue(currentCellRef.formatAsString());
		
		// Add cell formula.
		addedRow.createCell(2).setCellValue(c.getCellFormula());
	}
	
	private static SXSSFWorkbook getOutputWorkbook() {
		// Create the output Excel 2007 (.xlsx) spreadsheet and add a header.
		// We have to make this workbook a streaming one, so that we don't
		// run out of heap memory trying to contain the output file. We keep
		// 50000 rows in memory and then flush them in the file when another one
		// is created. Caller must release resources by calling dispose on the
		// returned workbook.
		SXSSFWorkbook wb = new SXSSFWorkbook(50000);
		Sheet resultSheet = wb.createSheet("Formula Census");
		Row headerRow = resultSheet.createRow(0);
		if (headerRow != null) {
			headerRow.createCell(0).setCellValue("FUSE URN ID");
			headerRow.createCell(1).setCellValue("Cell Location");
			headerRow.createCell(2).setCellValue("Formula in Cell");
			
			// Freeze the header row.
			resultSheet.createFreezePane( 0, 1, 0, 1 );
		}
		return wb;
	}
	
	public static void main(String[] args) throws Exception {
		// Must have at least one parameter.
		if (args.length == 0) {
			System.out.println("Must give the directory containing spreadsheet files");
			System.exit(1);
		}
		
		// Keep a count of the files processed so far.
		int fileCount = 0;
		int damagedFileCount = 0;
		int unreadFileCount = 0;
		
		try {
			SXSSFWorkbook resultWb = getOutputWorkbook();			
			DirectoryReader dirReader = new DirectoryReader(args[0]);			
			while (dirReader.hasMoreFilesInDirectory()) {
				String inputFile = dirReader.getNextFilename();
				InputStream fileStream = null;
				if (inputFile != null) {	// Sanity check.
					try {
						// File name is absolute, just get the name part.
						fileStream = new FileInputStream(inputFile);
				    	
				    	if (!fileStream.markSupported()) {
			                fileStream = new PushbackInputStream(fileStream, 8);
				    	} 
				    	
						try {
							// Create the workbook from the file.
							Workbook inputWb = createWorkBook(fileStream);						
							for (int sheetNo = 0; sheetNo < inputWb.getNumberOfSheets(); sheetNo++) {						
								Sheet currentSheet = inputWb.getSheetAt(sheetNo);					    	
								for (Row currentRow : currentSheet) {
									// No empty rows.
									if (currentRow != null) {								
										for (Cell currentCell : currentRow) {
											// Only formula cells.
											if (currentCell != null && 
												currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
												addFormulaStat(resultWb, inputFile, currentCell);
											}
										}
									}							
								}						
							}
						} catch (Exception e) {
							if (e instanceof InvalidFormatException) {
								System.out.println("File format not supported. " + e.getMessage());
							} else if (e instanceof EncryptedDocumentException) {
								System.out.println("File is encrypted. " + e.getMessage());
							} else if (e instanceof IOException) {
								System.out.println("IO exception " + e.getMessage());
							} else if (e instanceof IllegalArgumentException || e instanceof IndexOutOfBoundsException) {
								System.out.println("File has wrong format or file is damaged and needs repair." + e.getMessage());
								damagedFileCount++;								
							} else if (e instanceof RuntimeException) {
								System.out.println("Runtime exception occurred. " + e.getMessage());
							} else {
								// We didn't anticipate this.
								e.printStackTrace();
							}
							unreadFileCount++;				
						} finally {
							fileStream.close();
						}
					} 
					catch (EncryptedDocumentException e) {
						System.out.println("File is encrypted. " + e.getMessage());
					}
					
					// Increase the counter.
					fileCount++;
					
					// We want to write a file for every 10,000 files we process or 
					// we have reached the last file in the directory. This is done to
					// avoid the too many files open error, process limit is 10,240 for OSX.
					if (fileCount % 10000 == 0 || !dirReader.hasMoreFilesInDirectory()) {
						// Get the output file suffix.
						int resultFileSuffix = (int)fileCount/10000;
						if (!dirReader.hasMoreFilesInDirectory()) {
							// This is the last chunk of files less than 10,000 files.
							resultFileSuffix = resultFileSuffix + 1;
						}
						
						// Write the result to the stream.
					    FileOutputStream resultStream = new FileOutputStream("./Output/FUSE_formula_Census_" + Integer.toString(resultFileSuffix) + ".xlsx");
					    resultWb.write(resultStream);
					    resultWb.close();
					    resultStream.close();
					    
					    // Release temporary resources when done.
					    resultWb.dispose();
					    
					    // Create a new one for the next files only if we have more files to process.
					    if (dirReader.hasMoreFilesInDirectory()) {
							resultWb = getOutputWorkbook();
						}
					}
				}
			}
		} catch (MessageException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println("IO exception " + e.getMessage());			
		}
		
		// Writer for the statistics file. Will be created anew for every run.
		String stat = "Total files read: " + Integer.toString(fileCount) + System.lineSeparator() +
						"Total damaged files: " + Integer.toString(damagedFileCount) + System.lineSeparator() +
						"Total unread files: " + Integer.toString(unreadFileCount) + System.lineSeparator();
		
		// Print on the console to denote end of processing.
		System.out.println(stat);
		
		// Dump the text file.
		PrintWriter processedFilesStatWriter = null;
		processedFilesStatWriter = new PrintWriter("./Output/Stats.txt");
		processedFilesStatWriter.write(stat);
		processedFilesStatWriter.close();		
	}
}
