const fs = require("fs");
const officegen = require("officegen");
const { LOGO } = require("./constants");
const { getJsonForCreditNoteDocx } = require("./helper");

const download = async (req, res) => {
  try {
    const creditNote = req.body;
    res.writeHead(200, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-disposition": "attachment; filename=surprise.pptx",
    });
    const data = getJsonForCreditNoteDocx(creditNote);

    const docx = officegen({
      type: "docx",
      orientation: "portrait",
      pageMargins: {
        top: 500,
        left: 500,
        bottom: 500,
        right: 500,
      },
    });

    docx.on("error", (e) => console.log(e));
    docx.createByJson(data);
    // fs.writeFileSync("./data.json", JSON.stringify(data, null, 2));
    // Async call to generate the output file:
    setHeaderAndFooter(docx);
    return docx.generate(res);
    // return res.json({ success: true });
  } catch (error) {
    console.log(error);
    return res.json({ error: true, success: false, message: error.message });
  }
};

const setHeaderAndFooter = (docx) => {
  const header = docx.getHeader().createP({ align: "left" });
  header.addImage(LOGO, { cy: 84.661417323, cx: 114.8976378, align: "left" });

  // TODO
  // align this on the right side of the document's header
  // header.addText(
  //   "Fefifo Malaysia Sdn Bhd\n64-3, Jalan 27/70a\nDesa Sri Hartamas\n50480 Kuala Lumpur, Malaysia\nCompany No. 834529-T"
  // );
  return docx;
};

module.exports = { download };
