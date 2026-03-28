/**
 * Example: Branded .docx report using shared helpers.
 *
 * This demonstrates the standard pattern for generating reports.
 * The report uses whatever brand is configured in brand-config.js.
 *
 * Run: node references/generate-example.js
 */

const path = require("path");

// -- Import shared helpers --
const SKILL_REF = __dirname;
const { createTable, PAGE_W } = require(path.join(SKILL_REF, "table-templates"));
const { heading1, heading2, bodyText, bulletPoint, metaLine, spacer, pageBreak } = require(path.join(SKILL_REF, "typography"));
const { titleBlock, buildDocument, packAndWrite } = require(path.join(SKILL_REF, "page-setup"));
const { ORG_NAME } = require(path.join(SKILL_REF, "brand-config"));

// -- Report content --
const children = [];

// Title block
children.push(...titleBlock("Quarterly Sales Report", "Q1 2026 Performance Summary"));

// Metadata
children.push(metaLine("Report Date: ", "March 28, 2026"));
children.push(metaLine("Prepared by: ", ORG_NAME + " Analytics Team"));
children.push(spacer());

// Section 1: Summary table
children.push(heading1("Top 5 Products by Revenue"));
children.push(createTable(
  ["Rank", "Product Code", "Product Name", "Revenue ($)"],
  [
    ["1", "PRD-001", "Enterprise Platform License", "1,245,000"],
    ["2", "PRD-012", "Professional Services Bundle", "892,500"],
    ["3", "PRD-003", "Data Analytics Add-on", "634,200"],
    ["4", "PRD-007", "API Integration Package", "521,800"],
    ["5", "PRD-015", "Training & Certification", "378,900"],
  ],
  [700, 1200, 5460, 2000],
  { aligns: ["center", "center", "left", "right"] }
));
children.push(pageBreak());

// Section 2: Observations
children.push(heading1("Key Observations"));

children.push(heading2("Growth Trends"));
children.push(bodyText(
  "Enterprise Platform License revenue grew 23% quarter-over-quarter, driven by " +
  "three new enterprise contracts signed in February. The Professional Services Bundle " +
  "continues to show strong attach rates with new license sales."
));

children.push(heading2("Notable Findings"));
children.push(bulletPoint("Enterprise licenses account for 34% of total Q1 revenue"));
children.push(bulletPoint("API Integration Package saw the highest growth rate at 41% QoQ"));
children.push(bulletPoint("Training & Certification is a leading indicator for upsell opportunities"));
children.push(spacer());

// Section 3: Methodology
children.push(heading1("Methodology"));
children.push(bulletPoint("Source: Internal CRM and billing system exports"));
children.push(bulletPoint("Period: January 1 — March 31, 2026"));
children.push(bulletPoint("Revenue figures are net of discounts and refunds"));

// -- Build and write --
const doc = buildDocument(children, { headerText: ORG_NAME + " — Quarterly Report" });

const outputPath = path.join(__dirname, "example-output.docx");
packAndWrite(doc, outputPath);
