/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package org.vektor;

import javax.xml.parsers.DocumentBuilderFactory;

import org.vektor.util.XmlMailMergeDataSet;

import com.aspose.words.Document;

public class OrderGenerator {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.

		// Use DocumentBuilder from the javax.xml.parsers package and Document
		// class from the org.w3c.dom package to read
		// the XML data file and store it in memory.
		javax.xml.parsers.DocumentBuilder db = DocumentBuilderFactory
				.newInstance().newDocumentBuilder();

		// Parse the XML data.
		org.w3c.dom.Document xmlData = db.parse(OrderGenerator.class
				.getResourceAsStream("../../data/Orders.xml"));

		// Open a template document.
		Document doc = new Document(
				OrderGenerator.class
						.getResourceAsStream("../../data/OrderTemplate.doc"));

		doc.getMailMerge().setUseNonMergeFields(true);

		// Note that this class also works with a single repeatable region (and
		// any nested regions).
		// To merge multiple regions at the same time from a single XML data
		// source, use the XmlMailMergeDataSet class.
		// e.g doc.getMailMerge().executeWithRegions(new
		// XmlMailMergeDataSet(xmlData));
		
		XmlMailMergeDataSet dataSet = new XmlMailMergeDataSet(xmlData);
		doc.getMailMerge().executeWithRegions(dataSet);

		// Save the output document.
		doc.save("output/Order.docx");
	}

}