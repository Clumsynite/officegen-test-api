const fs = require("fs");

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

// PAGE 1
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

    const data = [
      {
        type: "text",
        val: "Credit Note",
        opt: { bold: true, font_size: 14 },
        lopt: { align: "center" },
      },
      { type: "table", val: detailsTable, opt: tableStyle },
      { type: "text", val: "" },
      { type: "table", val: creditNoteTable, opt: tableStyle },
      { type: "text", val: "" },
      { type: "text", val: "Notes/Terms:", opt: { font_size: 8, italic: true } },
    ];

    return data;
  } catch (error) {
    console.log("error in getJSONForCreditNote", error);
  }
};

const harvestRow = (harvest) => {
  const opts = {
    sz: "16",
    vAlign: "center",
    align: "center",
  };
  return [
    {
      val: `${harvest.agroProfile.name.fullName}\r
      ${harvest.agroProfile.code}\r
      ${harvest.agroProfile.assignedCropPlotCode}\r
      ${harvest.agroProfile.assignedCropModelCode}
    `,
      opts: { ...opts, align: "left" },
    },
    {
      val: `Cycle: ${harvest.cycleNo}\r
    Section: ${harvest.sectionNo}\r
    Harvest: ${harvest.harvestPeriod}`,
      opts: { ...opts, align: "left" },
    },
    { val: `${harvest.gradeAYield} kg`, opts: { ...opts } },
    { val: `${harvest.nonGradeAYield} kg`, opts: { ...opts } },
    { val: `${harvest.gradeArejectAmt || 0} kg`, opts: { ...opts } },
    { val: `${harvest.nonGradeArejectAmt || 0} kg`, opts: { ...opts } },
  ];
};

const getJSONForSalesRejectsBreakdown = (creditNote) => {
  const tableHeading = [
    { val: "", opts: { cellColWidth: 20000 } },
    { val: "", opts: { cellColWidth: 7000 } },
    { val: "", opts: { cellColWidth: 7000 } },
    { val: "", opts: { cellColWidth: 7000 } },
    { val: "", opts: { cellColWidth: 7000 } },
    { val: "", opts: { cellColWidth: 7000 } },
  ];
  const rejectsTable = [tableHeading];
  const opts = {
    sz: "16",
  };
  const rejectsLabel = [{ val: "Assigned Rejects", opts: { ...opts } }];
  const labeledHeading = [
    { val: "", opts: { ...opts } },
    { val: "Harvest Details", opts: { ...opts, b: true } },
    { val: "Grade A Yield", opts: { ...opts, b: true } },
    { val: "Non-Grade A Yield", opts: { ...opts, b: true } },
    { val: "Grade A Reject", opts: { ...opts, b: true } },
    { val: "Non-Grade A Reject", opts: { ...opts, b: true } },
  ];
  for (let i = 0; i < creditNote.deliveryDetails.length; i++) {
    const cropModel = creditNote.deliveryDetails[i];
    const heading = [{ val: cropModel.cropModelCode, opts: { gridSpan: 6, sz: "18", b: true, align: "left" } }];
    rejectsTable.push(heading);

    for (let harvest of cropModel.harvestPeriods) {
      rejectsTable.push(rejectsLabel);
      rejectsTable.push(labeledHeading);
      rejectsTable.push(harvestRow(harvest));
    }
  }

  const data = [
    { type: "text", val: "SALES REJECTS BREAKDOWN", opt: { bold: true, font_size: 14 }, lopt: { align: "center" } },
    { type: "table", val: rejectsTable, opt: { borders: false, tableFontFamily: "Calibri" } },
  ];
  return data;
};

module.exports = { getJsonForCreditNoteDocx, getJSONForSalesRejectsBreakdown };
