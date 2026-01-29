// Minimal pptxgenjs scaffold for a branded base deck.
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

// TODO: set author/company/title
pres.author = "Company Name";
pres.company = "Company Name";
pres.title = "Company Template";

// TODO: add slides and shapes here.
// Keep colors/fonts minimal; apply real theme with potxkit.

pres.writeFile({ fileName: "template-base.pptx" });
