import { Document, Packer, Paragraph, HeadingLevel } from "docx";
import { saveAs } from "file-saver";

function saveDocumentToFile(doc, fileName) {
  // Create new instance of Packer for the docx module

  // Create a mime type that will associate the new file with Microsoft Word
  const mimeType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  // Create a Blob containing the Document instance and the mimeType
  Packer.toBlob(doc).then((blob) => {
    const docblob = blob.slice(0, blob.size, mimeType);
    // Save the file using saveAs from the file-saver package
    saveAs(docblob, fileName);
  });
}

function generateWordDocument(event) {
  event.preventDefault();
  // Create a new instance of Document for the docx module
  let doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: "myCustomStyle",
          name: "My Custom Style",
          basedOn: "Normal",
          run: {
            color: "FF0000",
            italics: true,
            bold: true,
            size: 26,
            font: "Calibri"
          },
          paragraph: {
            spacing: { line: 276, before: 150, after: 150 }
          }
        }
      ]
    },
    sections: [
      {
        children: [
          new Paragraph({ text: "Title", heading: HeadingLevel.TITLE }),
          new Paragraph({ text: "Heading 1", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: "Heading 2", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({
            text:
              "Aliquam gravida quam sapien, quis dapibus eros malesuada vel. Praesent tempor aliquam iaculis. Nam ut neque ex. Curabitur pretium laoreet nunc, ut ornare augue aliquet sed. Pellentesque laoreet sem risus. Cras sodales libero convallis, convallis ex sed, ultrices neque. Sed quis ullamcorper mi. Ut a leo consectetur, scelerisque nibh sit amet, egestas mauris. Donec augue sapien, vestibulum in urna et, cursus feugiat enim. Ut sit amet placerat quam, id tincidunt nulla. Cras et lorem nibh. Suspendisse posuere orci nec ligula mattis vestibulum. Suspendisse in vestibulum urna, non imperdiet enim. Vestibulum vel dolor eget neque iaculis ultrices."
          }),
          new Paragraph({
            text: "This is a paragraph styled with my custom style",
            style: "myCustomStyle"
          })
        ]
      }
    ]
  });
  // Call saveDocumentToFile with the document instance and a filename
  saveDocumentToFile(doc, "New Document.docx");
}

// Listen for clicks on Generate Word Document button
document.getElementById("generate").addEventListener(
  "click",
  function (event) {
    generateWordDocument(event);
  },
  false
);
