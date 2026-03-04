const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  PageBreak,
  UnderlineType,
  Footer
} = require("docx");

const { ChartJSNodeCanvas } = require("chartjs-node-canvas");
require("chart.js/auto");
const fs = require("fs");

const chartCanvas = new ChartJSNodeCanvas({
  width: 800,
  height: 500
});

exports.generateReport = async (req, res) => {
  try {
    const d = req.body;
    const photos = req.files || [];

    // ================= COMMON FORMAT =================

    // HEADINGS → SIZE 14 (28), BOLD, UNDERLINE, CAPS, SPACING ADDED
    const heading = (text) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 400 }, // space after heading
        children: [
          new TextRun({
            text: text.toUpperCase(),
            font: "Times New Roman",
            size: 28,
            bold: true,
            underline: { type: UnderlineType.SINGLE }
          })
        ]
      });

    // NORMAL TEXT → SIZE 14, SPACING ADDED
    const normalText = (text) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 100, after: 200 },
        children: [
          new TextRun({
            text,
            font: "Times New Roman",
            size: 28
          })
        ]
      });

    const blank = () => new Paragraph({ text: "" });

    const children = [];

    // ================= PAGE 1 =================
    children.push(heading(d.collegeName));
    children.push(heading(d.departmentName));
    children.push(heading(`Camp Report – ${d.campLocation}`));
    children.push(heading(`Date: ${d.reportDateShort}`));
    children.push(blank());

    children.push(normalText(
      `The Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and with ${d.projectName} conducted a dental treatment camp at ${d.campLocation} on ${d.reportDateLong}.`
    ));

    children.push(normalText(
      `The Camp started at ${d.startTime} and concluded at ${d.endTime}. A team of dentists including ${d.staffCount} staff, ${d.postgraduateCount} postgraduate and ${d.internCount} interns rendered oral health care for the public.`
    ));

    children.push(normalText(
      `A total of ${d.totalPatients} people attended the dental camp and ${d.treatmentCount} people were treated. Oral cavity examination was done with oral health talk and oral hygiene instructions.`
    ));

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= PAGE 2 (PHOTOS) =================
    children.push(heading("Photos"));
    children.push(blank());

    for (let photo of photos) {
      const img = fs.readFileSync(photo.path);
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
          children: [
            new ImageRun({
              data: img,
              transformation: { width: 400, height: 250 }
            })
          ]
        })
      );
    }

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= PAGE 3 (CAMP STATISTICS) =================
    const campChart = await chartCanvas.renderToBuffer({
      type: "bar",
      data: {
        labels: ["Male", "Female"],
        datasets: [{
          data: [parseInt(d.maleCount), parseInt(d.femaleCount)],
          backgroundColor: "blue"
        }]
      },
      options: { plugins: { legend: { display: false } } }
    });

    children.push(heading("Camp Statistics"));
    children.push(normalText(`Total Number of Patients    ${d.totalPatients}`));
    children.push(normalText(`Male    ${d.maleCount}`));
    children.push(normalText(`Female    ${d.femaleCount}`));
    children.push(blank());

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 }, // space before graph
        children: [
          new ImageRun({
            data: campChart,
            transformation: { width: 500, height: 300 }
          })
        ]
      })
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= PAGE 4 (SCREENING STATISTICS) =================
    const screeningChart = await chartCanvas.renderToBuffer({
      type: "bar",
      data: {
        labels: ["Dental Caries","Root Stump","Gingivitis","Periodontitis","Missing","Consultation","Others"],
        datasets: [{
          data: [
            d.dentalCaries,
            d.rootStump,
            d.gingivitis,
            d.periodontitis,
            d.missing,
            d.consultation,
            d.others
          ],
          backgroundColor: "blue"
        }]
      },
      options: { plugins: { legend: { display: false } } }
    });

    children.push(heading("Screening Statistics"));
    children.push(normalText(`Dental Caries    ${d.dentalCaries}`));
    children.push(normalText(`Root Stump    ${d.rootStump}`));
    children.push(normalText(`Gingivitis    ${d.gingivitis}`));
    children.push(normalText(`Periodontitis    ${d.periodontitis}`));
    children.push(normalText(`Missing    ${d.missing}`));
    children.push(normalText(`Consultation    ${d.consultation}`));
    children.push(normalText(`Others    ${d.others}`));
    children.push(blank());

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [
          new ImageRun({
            data: screeningChart,
            transformation: { width: 500, height: 300 }
          })
        ]
      })
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ================= PAGE 5 (TREATMENT STATISTICS) =================
    const treatmentChart = await chartCanvas.renderToBuffer({
      type: "bar",
      data: {
        labels: ["Scaling"],
        datasets: [{
          data: [d.scaling],
          backgroundColor: "blue"
        }]
      },
      options: { plugins: { legend: { display: false } } }
    });

    children.push(heading("Treatment Statistics"));
    children.push(normalText(`Scaling    ${d.scaling}`));
    children.push(blank());

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [
          new ImageRun({
            data: treatmentChart,
            transformation: { width: 500, height: 300 }
          })
        ]
      })
    );

    // ================= FOOTER =================
    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200 },
          children: [
            new TextRun({
              text: "HEAD OF THE DEPARTMENT          PRINCIPAL",
              font: "Times New Roman",
              size: 28,
              bold: true
            })
          ]
        })
      ]
    });

    // ================= CREATE DOCUMENT =================
    const doc = new Document({
      sections: [
        {
          footers: { default: footer },
          children
        }
      ]
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Disposition", "attachment; filename=Camp_Report.docx");
    res.send(buffer);

  } catch (err) {
    console.log(err);
    res.status(500).send("Error generating report");
  }
};