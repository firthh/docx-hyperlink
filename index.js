const officegen = require("officegen");
const fs = require("fs");

function createDocument(filename, linkLocation) {
  let docx = officegen("docx");

  docx.on("finalize", function () {
    console.log(`finished creating ${filename}`);
  });

  docx.on("error", function (err) {
    console.log(err);
  });

  let pObj = docx.createP();

  pObj.addText("Simple document");

  pObj = docx.createP();

  pObj.addText("with an external link ");
  pObj.addText("external link", { link: linkLocation });

  let out = fs.createWriteStream(filename);

  out.on("error", function (err) {
    console.log(err);
  });

  // Async call to generate the output file:
  docx.generate(out);
}

createDocument("simplelink.docx", "https://github.com");
createDocument("withQueryParams.docx", "https://github.com?foo=bar&test=test");
createDocument(
  "withEscapedQueryParams.docx",
  "https://github.com?foo=bar&amp;test=test"
);
