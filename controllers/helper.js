const { LOGO } = require("./constants");

const getDetailsTable = (data) => [
  [
    {
      val: "Credit Note:.",
      opts: {
        cellColWidth: 5000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: data.creditNoteNo,
      opts: {
        cellColWidth: 10000,
        align: "left",
        sz: "16",
      },
    },
    {
      val: "Issue Date:",
      opts: {
        cellColWidth: 5000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: data.issueDate,
      opts: {
        cellColWidth: 10000,
        align: "left",
        sz: "16",
      },
    },
  ],
  [
    {
      val: "Bill To:",
      opts: {
        cellColWidth: 7000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: data.orderedBy.name,
      opts: {
        cellColWidth: 7000,
        align: "left",
        sz: "16",
      },
    },
    {
      val: "Original Invoice No.:",
      opts: {
        cellColWidth: 7000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: data.invoiceNo,
      opts: {
        cellColWidth: 7000,
        align: "left",
        sz: "16",
      },
    },
  ],
  [
    {
      val: "Attn:",
      opts: {
        cellColWidth: 7000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: data.contact.name,
      opts: {
        cellColWidth: 7000,
        align: "left",
        sz: "16",
        b: true,
      },
    },
    {
      val: "Original Invoice Date:",
      opts: {
        cellColWidth: 7000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: data.invoiceDate,
      opts: {
        cellColWidth: 7000,
        align: "left",
        sz: "16",
      },
    },
  ],
  [
    {
      val: "Contact",
      opts: {
        cellColWidth: 7000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: `+${data.contact.mobile.fullNumber}`,
      opts: {
        cellColWidth: 7000,
        align: "left",
        sz: "16",
      },
    },
    {
      val: "Buyer's Ref.:",
      opts: {
        cellColWidth: 7000,
        align: "left",
        b: true,
        sz: "16",
      },
    },
    {
      val: "",
      opts: {
        cellColWidth: 7000,
        align: "left",
        sz: "16",
      },
    },
  ],
];

const tableRow = (index, desc, qty, unit, price, amount) => {
  const opts = {
    sz: "16",
    align: "center",
    vAlign: "center",
  };
  return [
    {
      val: index,
      opts: { ...opts },
    },
    {
      val: desc,
      opts: { ...opts, align: "left" },
    },
    {
      val: qty,
      opts: { ...opts },
    },
    {
      val: unit,
      opts: { ...opts },
    },
    {
      val: price,
      opts: { ...opts },
    },
    {
      val: amount,
      opts: { ...opts },
    },
  ];
};

const getJsonForCreditNoteDocx = (creditNote) => {
  try {
    const detailsTable = getDetailsTable(creditNote);

    const tableStyle = {
      // tableColWidth: 20000,
      align: "left",
      tableSize: 24,
      tableAlign: "left",
      borders: true,
      tableFontFamily: "Calibri",
    };

    const darkOpts = {
      b: true,
      sz: "18",
      color: "ffffff",
      shd: {
        fill: "000000",
        themeFill: "text1",
        themeFillTint: "99",
      },
      fontFamily: "Arial",
    };

    const tableHeading = [
      {
        val: "S/N",
        opts: { ...darkOpts, cellColWidth: 2000 },
      },
      {
        val: "Description",
        opts: { ...darkOpts, cellColWidth: 20000 },
      },
      {
        val: "Qty",
        opts: { ...darkOpts, cellColWidth: 4000 },
      },
      {
        val: "UOM",
        opts: { ...darkOpts, cellColWidth: 4000 },
      },
      {
        val: "U/Price (MYR)",
        opts: { ...darkOpts, cellColWidth: 4000 },
      },
      {
        val: "Amount (MYR)",
        opts: { ...darkOpts, cellColWidth: 4000 },
      },
    ];

    let creditNoteTable = [tableHeading];

    for (let i = 0; i < creditNote.deliveryDetails.length; i++) {
      const cropModel = creditNote.deliveryDetails[i];
      const heading = [{ val: cropModel.cropModelCode, opts: { gridSpan: 6, sz: "16" } }];
      let gradeAContract = tableRow(
        1,
        `${cropModel.cropLabel || "[Crop Type] [Seed Variant] - Grade A (Contracted)"}`,
        cropModel.creditNote.gradeAContractedQty,
        cropModel.unitValue,
        cropModel.gradeAPriceContract,
        cropModel.creditNote.gradeAContractedTotal !== 0
          ? `(${Math.abs(cropModel.creditNote.gradeAContractedTotalTotal)})`
          : 0
      );
      let gradeAUncontracted = tableRow(
        2,
        `${cropModel.cropLabel || "[Crop Type] [Seed Variant] - Grade A (Uncontracted)"}`,
        cropModel.creditNote.gradeAUncontractedQty,
        cropModel.unitValue,
        cropModel.gradeAPriceSpot,
        cropModel.creditNote.gradeAUncontractedTotal !== 0
          ? `(${Math.abs(cropModel.creditNote.gradeAUncontractedTotal)})`
          : 0
      );
      let nonGradeA = tableRow(
        3,
        `${cropModel.cropLabel || "[Crop Type] [Seed Variant] - Non Grade A"}`,
        cropModel.creditNote.nonGradeAQty,
        cropModel.unitValue,
        cropModel.nonGradeAPrice,
        cropModel.creditNote.nonGradeATotal !== 0 ? `(${Math.abs(cropModel.creditNote.nonGradeATotal)})` : 0
      );
      let appComm = tableRow(
        4,
        `${cropModel.cropLabel || "Applicable Commission"}`,
        cropModel.creditNote.appCommQty,
        cropModel.unitValue,
        cropModel.appComm,
        cropModel.creditNote.appCommTotal
      );

      creditNoteTable.push(heading);
      creditNoteTable.push(gradeAContract);
      creditNoteTable.push(gradeAUncontracted);
      creditNoteTable.push(nonGradeA);
      creditNoteTable.push(appComm);
    }

    creditNoteTable.push([
      { val: "Total", opts: { sz: "16", b: true, align: "right", gridSpan: 5 } },
      {
        val: creditNote.creditNoteTotal < 0 ? `(${Math.abs(creditNote.creditNoteTotal)})` : creditNote.creditNoteTotal,
        opts: { sz: "16", b: true, align: "center" },
      },
    ]);
    // fs.writeFileSync("creditNoteTable.json", JSON.stringify(creditNoteTable, null, 2));

    const data = [
      {
        type: "text",
        val: "Credit Note",
        opt: { bold: true, font_size: 16 },
        lopt: { align: "center" },
      },
      { type: "table", val: detailsTable, opt: tableStyle },
      { type: "linebreak" },
      // { type: "table", val: [tableHeading], opt: { ...tableStyle, tableColor: "525353" } },
      { type: "table", val: creditNoteTable, opt: tableStyle },
    ];

    return data;
  } catch (error) {
    console.log("error in getJSONForCreditNote", error);
  }
};

module.exports = { getJsonForCreditNoteDocx };
