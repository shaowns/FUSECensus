/**
 * Owner: ShaownS
 * File: DirectoryReader.java
 * Package: org.fusecensus.poireader
 * Project: FUSECensus
 * Email: ssarker@ncsu.edu
 */
package org.fusecensus.poireader;

import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;

/**
 * 
 */
public class DirectoryReader {
	
	public DirectoryReader(String dir) throws IOException, MessageException {
		// Basic checks.
		File folder = new File(dir);
		
		// Check if directory exists and is indeed a folder.
		if (!folder.exists()) {
			throw new MessageException("Directory doest not exist " + dir);
		}
		
		if (!folder.isDirectory()) {
			throw new MessageException("Directory is a file " + dir);
		}
		
		Path directoryPath = FileSystems.getDefault().getPath(dir);
		DirectoryStream<Path> stream = Files.newDirectoryStream(directoryPath);
		// Remember the iterator.
		this.filesIterator = stream.iterator();
	}
	
	public boolean hasMoreFilesInDirectory() {
		return filesIterator.hasNext();
	}
	
	public String getNextFilename() {
		while (filesIterator.hasNext()) {
			Path currentPath = filesIterator.next();
			File fileName = currentPath.toFile();
			if (fileName.isFile()) {
				return fileName.getAbsolutePath();
			}			
		}
		return null;
	}
	
	Iterator<Path> filesIterator;
}
