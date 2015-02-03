/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

import com.aspose.words.Document;

public class SimpleMailMerge {
	public static void main(String[] args) throws Exception {
		Document doc = new Document(SimpleMailMerge.class.getResourceAsStream("data/Template.doc"));
		// Fill the fields in the document with user data.
		doc.getMailMerge().execute(
				new String[] { "FullName", "Company", "Address", "Address2",
						"City" },
				new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "",
						"London" });
		// Saves the document to disk.
		doc.save("output/MailMerge Result Out.docx");
	}
}