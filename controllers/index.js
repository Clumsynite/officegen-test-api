const fs = require("fs");
const officegen = require("officegen");
const { LOGO } = require("./constants");
const { getJsonForCreditNoteDocx, getJSONForSalesRejectsBreakdown } = require("./helper");

const download = async (req, res) => {
  try {
    const creditNote = req.body;
    res.writeHead(200, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-disposition": "attachment; filename=surprise.pptx",
    });
    const data = getJsonForCreditNoteDocx(creditNote); // page - 1

    const docx = officegen({
      type: "docx",
      orientation: "portrait",
      pageMargins: {
        top: 400,
        left: 400,
        bottom: 400,
        right: 400,
      },
    });

    docx.on("error", (e) => console.log(e));
    docx.createByJson(data);
    const terms = [
      "Please deduct this amount from the Original Invoice No. stated above.",
      "This is part of a consolidated invoice of accounts receivable due to Fefifo, the exclusively authorised agent for Agropreneurs (sellers).",
      "All payments collected are on behalf of respective Agropreneurs (sellers) for crops produced and received by the buyer.",
      "Breakdown of participating Agropreneurs (sellers) for this sales rejects breakdown enclosed hereafter.",
      "Payment made by bank transfer, please quote our original invoice no. and email the bank slip after payment.",
    ];
    terms.forEach((term) => {
      const p = docx.createListOfDots({ border: "single" });
      p.addText(term, { font_size: 8, italic: true });
    });

    const tableStyle = {
      borders: false,
      tableFontFamily: "Calibri",
      tableSize: 10,
      align: "left",
      tableAlign: "left",
    };
    // docx.createP().addText({ font_size: 14 });
    const table = [
      [
        {
          val: "",
          opts: { font_size: 14, bold: true, cellColWidth: 10000 },
        },
        {
          val: "",
          opts: { font_size: 14, bold: true, cellColWidth: 10000 },
        },
      ],
      [
        {
          val: "Please make payment by bank transfer to the following:",
          opts: { sz: 18, bold: true, gridSpan: 2 },
        },
      ],
      [
        { val: "Bank Acct Name: Fefifo Malaysia Sdn Bhd", opts: { sz: 18 } },
        { val: "Bank Acct No.: 8010169358", opts: { sz: 18 } },
      ],
      [
        { val: "Bank Name: CIMB Bank Berhad", opts: { sz: 18 } },
        { val: "Swift Code: CIBBMYKL", opts: { sz: 18 } },
      ],
    ];
    docx.createP();
    docx.createTable(table, tableStyle);
    docx.createP();
    docx.createTable(
      [
        [{ val: "Thank you,", opts: { cellColWidth: 10000, sz: 18 } }],
        [
          { val: "FEFIFO MALAYSIA SDN BHD ", opts: { cellColWidth: 10000, sz: 18, b: true } },
          { val: "This is system generated. No signature required.", opts: { cellColWidth: 10000, sz: 18 } },
        ],
      ],
      {
        borders: false,
        tableFontFamily: "Calibri",
        tableSize: 10,
        align: "left",
        tableAlign: "left",
      }
    );
    docx.putPageBreak();

    // TODO
    // sales rejects breakdown
    const saleRejectsBreakDown = getJSONForSalesRejectsBreakdown(creditNote); // page - 2
    docx.createByJson(saleRejectsBreakDown);

    setHeaderAndFooter(docx);
    return docx.generate(res);
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
  const footer = docx.getFooter().createP();
  footer.addText("Fefifo Malaysia Sdn Bhd");
  return docx;
};

module.exports = { download };
