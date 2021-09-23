const fs = require("fs");
const officegen = require("officegen");
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
        top: 1000,
        left: 1000,
        bottom: 1000,
        right: 1000,
      },
    });

    docx.on("error", (e) => console.log(e));
    docx.createByJson(data);
    // fs.writeFileSync("./data.json", JSON.stringify(data, null, 2));
    // Async call to generate the output file:
    return docx.generate(res);
    // return res.json({ success: true });
  } catch (error) {
    console.log(error);
    return res.json({ error: true, success: false, message: error.message });
  }
};

module.exports = { download };
