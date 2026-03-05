const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  PageBreak,
  UnderlineType,
  Footer,
  Table,
  TableRow,
  TableCell,
  WidthType,
} = require("docx");

const { ChartJSNodeCanvas } = require("chartjs-node-canvas");
require("chart.js/auto");
const fs = require("fs");

const chartCanvas = new ChartJSNodeCanvas({
  width: 800,
  height: 500,
});

exports.generateReport = async (req, res) => {
  try {
    const d = req.body;
    const photos = req.files || [];

    // ===== HELPER FUNCTIONS =====

    const heading = (text) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 360 },
        children: [
          new TextRun({
            text: text.toUpperCase(),
            font: "Times New Roman",
            size: 28,
            bold: true,
            underline: { type: UnderlineType.SINGLE },
          }),
        ],
      });

    const normalText = (text, center = false) =>
      new Paragraph({
        alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { line: 360 },
        children: [
          new TextRun({
            text: String(text),
            font: "Times New Roman",
            size: 24,
          }),
        ],
      });

    const blank = () =>
      new Paragraph({
        text: "",
        spacing: { line: 360 },
      });

    const children = [];

    // ================= PAGE 1 =================

    children.push(heading(d.collegeName));
    children.push(heading(d.departmentName));
    children.push(heading(`Camp Report – ${d.campLocation}`));
    children.push(heading(`Date: ${d.reportDateShort}`));
    children.push(blank());

    children.push(
      normalText(
        `Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and with ${d.projectName} conducted a dental screening and treatment camp at ${d.campLocation} on ${d.reportDateLong}.`
      )
    );

    children.push(
      normalText(
        `Dr R. Palanivel Pandian organised this program. The Camp started at ${d.startTime} and ended at ${d.endTime}. A team of dentists including ${d.staffCount} staff member, ${d.postgraduateCount} postgraduate and ${d.internCount} interns provided oral health care to the people.`
      )
    );

    children.push(
      normalText(
        `A total of ${d.totalPatients} people attended the dental camp and ${d.treatmentCount} people were treated along with oral health education and oral hygiene instructions.`
      )
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= PHOTOS =================

    children.push(heading("Photos"));
    children.push(blank());

    for (let i = 0; i < photos.length; i += 4) {
      const row1 = [];
      const row2 = [];

      for (let j = i; j < i + 2 && j < photos.length; j++) {
        const img = fs.readFileSync(photos[j].path);

        row1.push(
          new ImageRun({
            data: img,
            transformation: { width: 300, height: 200 },
          })
        );
      }

      for (let j = i + 2; j < i + 4 && j < photos.length; j++) {
        const img = fs.readFileSync(photos[j].path);

        row2.push(
          new ImageRun({
            data: img,
            transformation: { width: 300, height: 200 },
          })
        );
      }

      if (row1.length > 0) {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: row1,
          })
        );
      }

      if (row2.length > 0) {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: row2,
          })
        );
      }

      children.push(new Paragraph({ children: [new PageBreak()] }));
    }

    // ================= CAMP STATISTICS =================

    const campDataRows = [
      ["Male", d.maleCount],
      ["Female", d.femaleCount],
    ];

    const campTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 60, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({ children: [normalText("Gender", true)] }),
            new TableCell({ children: [normalText("No of Patients", true)] }),
          ],
        }),
        ...campDataRows.map(
          (row) =>
            new TableRow({
              children: row.map(
                (val) =>
                  new TableCell({
                    children: [normalText(val, true)],
                  })
              ),
            })
        ),
      ],
    });

    const campChart = await chartCanvas.renderToBuffer({
      type: "bar",
      data: {
        labels: ["Male", "Female"],
        datasets: [
          {
            label: "No of Patients",
            data: [parseInt(d.maleCount), parseInt(d.femaleCount)],
            backgroundColor: "lightblue",
          },
        ],
      },
      options: { plugins: { legend: { display: false } } },
    });

    children.push(heading("Camp Statistics"));
    children.push(campTable);
    children.push(blank());

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: campChart,
            transformation: { width: 500, height: 300 },
          }),
        ],
      })
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= SCREENING STATISTICS =================

    const screeningDataRows = [
      ["Dental Caries", d.dentalCaries],
      ["Root Stump", d.rootStump],
      ["Gingivitis", d.gingivitis],
      ["Periodontitis", d.periodontitis],
      ["Missing", d.missing],
      ["Consultation", d.consultation],
      ["Others", d.others],
    ];

    const screeningTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 70, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({ children: [normalText("Diagnosis", true)] }),
            new TableCell({ children: [normalText("No of Patients", true)] }),
          ],
        }),
        ...screeningDataRows.map(
          (row) =>
            new TableRow({
              children: row.map(
                (val) =>
                  new TableCell({
                    children: [normalText(val, true)],
                  })
              ),
            })
        ),
      ],
    });

    const screeningChart = await chartCanvas.renderToBuffer({
      type: "bar",
      data: {
        labels: screeningDataRows.map((r) => r[0]),
        datasets: [
          {
            label: "No of Patients",
            data: screeningDataRows.map((r) => r[1]),
            backgroundColor: "lightblue",
          },
        ],
      },
      options: { plugins: { legend: { display: false } } },
    });

    children.push(heading("Screening Statistics"));
    children.push(screeningTable);
    children.push(blank());

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: screeningChart,
            transformation: { width: 500, height: 300 },
          }),
        ],
      })
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= TREATMENT STATISTICS =================

    const treatmentDataRows = [["Scaling", d.scaling]];

    const treatmentTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 60, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({ children: [normalText("Treatment", true)] }),
            new TableCell({ children: [normalText("No of Patients", true)] }),
          ],
        }),
        ...treatmentDataRows.map(
          (row) =>
            new TableRow({
              children: row.map(
                (val) =>
                  new TableCell({
                    children: [normalText(val, true)],
                  })
              ),
            })
        ),
      ],
    });

    const treatmentChart = await chartCanvas.renderToBuffer({
      type: "bar",
      data: {
        labels: ["Scaling"],
        datasets: [
          {
            label: "No of Patients",
            data: [d.scaling],
            backgroundColor: "lightblue",
          },
        ],
      },
      options: { plugins: { legend: { display: false } } },
    });

    children.push(heading("Treatment Statistics"));
    children.push(treatmentTable);
    children.push(blank());

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: treatmentChart,
            transformation: { width: 500, height: 300 },
          }),
        ],
      })
    );

    // ================= FOOTER =================

    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 360 },
          children: [
            new TextRun({
              text:
                "HEAD OF THE DEPARTMENT                                      PRINCIPAL",
              font: "Times New Roman",
              size: 28,
              bold: true,
            }),
          ],
        }),
      ],
    });

    // ================= DOC CREATION =================

    const doc = new Document({
      sections: [
        {
          footers: { default: footer },
          children,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=Camp_Report.docx"
    );

    res.send(buffer);
  } catch (err) {
    console.log(err);
    res.status(500).send("Error generating report");
  }
};
