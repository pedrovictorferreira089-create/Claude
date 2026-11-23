const L = require("./lib");
const { D, section, save, numbering, styles } = L;
const out = process.argv[2];
const capa = require("./capa");
const corpo = require("./artigo");
const children = [...capa, new D.Paragraph({ children: [new D.PageBreak()] }), ...corpo];
const doc = new D.Document({ creator: "", title: "", description: "", lastModifiedBy: "", styles, numbering, sections: [section(children)] });
save(doc, out).then(() => console.log("ok", out));
