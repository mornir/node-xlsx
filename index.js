const ExcelJS = require("exceljs")
const sanityClient = require("@sanity/client")
const toPlainText = require("./toPlainText")

const workbook = new ExcelJS.Workbook()

const worksheet = workbook.addWorksheet("My Sheet")

worksheet.columns = [
  { header: "Begriff", key: "term", width: 32 },
  { header: "Definition", key: "definition", width: 62 },
]

const client = sanityClient({
  projectId: "nipfx4rq",
  dataset: "production",
  apiVersion: "2021-06-19", // use current UTC date - see "specifying API version"!
  useCdn: false, // `false` if you want to ensure fresh data
})

const query = /* groq */ `*[_type == 'entry' && status == "definition"] {deTitle, 'definition': content.de.definition}`

client
  .fetch(query)
  .then((entries) => {
    entries.forEach((entry) => {
      worksheet.addRow({
        term: entry.deTitle,
        definition: toPlainText(entry.definition),
      })
    })
    return workbook.xlsx.writeFile("test.xlsx")
  })
  .then(() => {
    console.log("success")
  })
  .catch((e) => {
    console.log(e)
  })
